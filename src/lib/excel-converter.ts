import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// ========== Types ==========
export interface MappingEntry {
  code: string;
  systemName: string;
}

export interface MappingTable {
  [originalName: string]: MappingEntry;
}

export interface DataRow {
  date: string;
  code: string;
  originalStoreName: string;
  storeName: string;
  productName: string;
  boxQty: number;
  afternoon: string;
  isMappingFailed: boolean;
}

export interface ValidationRow {
  date: string;
  dayName: string;
  extractedBox: number;
  originalBox: number;
  result: string;
  mappingFailureStores: number;
  mappingFailureRows: number;
}

export interface StoreDailyRow {
  date: string;
  code: string;
  storeName: string;
  boxSum: number;
}

export interface ConversionResult {
  data: DataRow[];
  validation: ValidationRow[];
  storeDaily: StoreDailyRow[];
  mappingFailures: string[];
}

// ========== WASM Types ==========
interface WasmDataRow {
  date: string;
  code: string;
  original_store_name: string;
  store_name: string;
  product_name: string;
  box_qty: number;
  afternoon: string;
  mapping_failed: string;
}

interface WasmValidationRow {
  date: string;
  day_name: string;
  extracted_box: number;
  original_box: number;
  result: string;
  mapping_failure_stores: number;
  mapping_failure_rows: number;
}

interface WasmStoreDailyRow {
  date: string;
  code: string;
  store_name: string;
  box_sum: number;
}

interface WasmConversionResult {
  data: WasmDataRow[];
  validation: WasmValidationRow[];
  store_daily: WasmStoreDailyRow[];
  mapping_failures: string[];
  success: boolean;
  error?: string;
}

interface WasmModule {
  convert_excel: (originData: Uint8Array, mappingData: Uint8Array, filename: string) => WasmConversionResult;
  default: (path?: string) => Promise<unknown>;
}

// ========== WASM Module Manager ==========
class WasmModuleManager {
  private instance: WasmModule | null = null;
  private ready = false;
  private initPromise: Promise<boolean> | null = null;

  async initialize(): Promise<boolean> {
    if (this.ready) return true;
    if (this.initPromise) return this.initPromise;

    this.initPromise = this._doInitialize();
    return this.initPromise;
  }

  private async _doInitialize(): Promise<boolean> {
    try {
      // Dynamic import of WASM module
      const wasm = await import('@/wasm/excel_converter_wasm.js') as unknown as WasmModule;
      
      // Initialize with public path for wasm file
      const wasmPath = import.meta.env.BASE_URL + 'excel_converter_wasm_bg.wasm';
      await wasm.default(wasmPath);
      
      this.instance = wasm;
      this.ready = true;
      console.log('WASM module loaded successfully');
      return true;
    } catch (e) {
      console.warn('WASM not available, using JS fallback:', e instanceof Error ? e.message : e);
      return false;
    }
  }

  isReady(): boolean {
    return this.ready;
  }

  convert(originData: Uint8Array, mappingData: Uint8Array, filename: string): WasmConversionResult {
    if (!this.instance) {
      throw new Error('WASM module not initialized');
    }
    const result = this.instance.convert_excel(originData, mappingData, filename);
    if (!result.success) {
      throw new Error(result.error || '변환 실패');
    }
    return result;
  }
}

const wasmModule = new WasmModuleManager();

// ========== WASM Result Mapper ==========
function mapWasmResultToJs(wasmResult: WasmConversionResult): ConversionResult {
  return {
    data: wasmResult.data.map(r => ({
      date: r.date,
      code: r.code,
      originalStoreName: r.original_store_name,
      storeName: r.store_name,
      productName: r.product_name,
      boxQty: r.box_qty,
      afternoon: r.afternoon || '',
      isMappingFailed: r.mapping_failed === 'Y'
    })),
    validation: wasmResult.validation.map(r => ({
      date: r.date,
      dayName: r.day_name,
      extractedBox: r.extracted_box,
      originalBox: r.original_box,
      result: r.result,
      mappingFailureStores: r.mapping_failure_stores,
      mappingFailureRows: r.mapping_failure_rows
    })),
    storeDaily: wasmResult.store_daily.map(r => ({
      date: r.date,
      code: r.code,
      storeName: r.store_name,
      boxSum: r.box_sum
    })),
    mappingFailures: wasmResult.mapping_failures
  };
}

// ========== File Utilities ==========
async function readFileAsUint8Array(file: File): Promise<Uint8Array> {
  const buffer = await file.arrayBuffer();
  return new Uint8Array(buffer);
}

// ========== Constants (for JS fallback) ==========
const DAY_NAMES = ['월', '화', '수', '목', '금'];
const MAX_PRODUCTS_PER_BLOCK = 25;
const PRODUCT_DATA_START_OFFSET = 4;
const BOX_TOTAL_ROW = 8;
const BOX_TOTAL_COL = 6;

const STORE_BLOCK_LAYOUTS = [
  { nameCol: 2, afternoonCol: 3, productCol: 5, boxCol: 6 },
  { nameCol: 11, afternoonCol: 12, productCol: 14, boxCol: 15 },
  { nameCol: 20, afternoonCol: 21, productCol: 23, boxCol: 24 }
];

// ========== Utilities (for JS fallback) ==========
function extractStoreName(cellValue: unknown): string | null {
  if (!cellValue || typeof cellValue !== 'string') return null;
  if (!cellValue.includes('※')) return null;

  const match = cellValue.match(/※\s*(.+?)\s*:\s*\d*/);
  return match ? match[1].trim() : null;
}

function extractDateFromFileName(fileName: string): Date {
  const dateMatch = fileName.match(/\((\d+)\.(\d+)~(\d+)\.(\d+)\)/);
  const yearMatch = fileName.match(/(\d+)년\s*(\d+)월/);

  if (dateMatch && yearMatch) {
    let year = parseInt(yearMatch[1]);
    if (year < 100) year += 2000;
    const month = parseInt(dateMatch[1]) - 1;
    const day = parseInt(dateMatch[2]);
    return new Date(year, month, day);
  }

  return new Date();
}

function formatDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getCellValue(worksheet: ExcelJS.Worksheet, row: number, col: number): unknown {
  const cell = worksheet.getCell(row, col);
  const value = cell.value;

  if (value && typeof value === 'object') {
    if ('result' in value) return value.result;
    if ('text' in value) return value.text;
  }

  return value;
}

// ========== Main Functions ==========
export async function initializeWasm(): Promise<boolean> {
  return wasmModule.initialize();
}

export function isWasmReady(): boolean {
  return wasmModule.isReady();
}

export async function parseMappingTable(file: File): Promise<MappingTable> {
  const workbook = new ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  const mapping: MappingTable = {};

  const headers: string[] = [];
  worksheet.getRow(1).eachCell((cell, colNumber) => {
    headers[colNumber] = String(cell.value || '');
  });

  const codeCol = headers.findIndex(h => h === '코드') + 1;
  const originalNameCol = headers.findIndex(h => h === '원본 사업장명') + 1;
  const systemNameCol = headers.findIndex(h => h === '사업장명') + 1;

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const originalName = String(row.getCell(originalNameCol).value || '').trim();
    if (originalName) {
      mapping[originalName] = {
        code: String(row.getCell(codeCol).value || ''),
        systemName: String(row.getCell(systemNameCol).value || '')
      };
    }
  });

  return mapping;
}

export async function convertExcel(
  originFile: File,
  mappingFile: File
): Promise<{ result: ConversionResult; mode: 'WASM' | 'JS' }> {
  // Try WASM first
  if (wasmModule.isReady()) {
    try {
      const [originData, mappingData] = await Promise.all([
        readFileAsUint8Array(originFile),
        readFileAsUint8Array(mappingFile)
      ]);

      const wasmResult = wasmModule.convert(originData, mappingData, originFile.name);
      return {
        result: mapWasmResultToJs(wasmResult),
        mode: 'WASM'
      };
    } catch (e) {
      console.warn('WASM conversion failed, falling back to JS:', e);
    }
  }

  // Fallback to JS
  const mapping = await parseMappingTable(mappingFile);
  const result = await convertExcelWithJs(originFile, mapping);
  return { result, mode: 'JS' };
}

async function convertExcelWithJs(
  originFile: File,
  mapping: MappingTable
): Promise<ConversionResult> {
  const workbook = new ExcelJS.Workbook();
  const buffer = await originFile.arrayBuffer();
  await workbook.xlsx.load(buffer);

  const baseDate = extractDateFromFileName(originFile.name);
  const dayDates: Record<string, Date> = {};
  DAY_NAMES.forEach((day, idx) => {
    const date = new Date(baseDate);
    date.setDate(date.getDate() + idx);
    dayDates[day] = date;
  });

  const allData: DataRow[] = [];
  const mappingFailures: { day: string; storeName: string }[] = [];

  for (const [dayName, date] of Object.entries(dayDates)) {
    const worksheet = workbook.getWorksheet(dayName);
    if (!worksheet) continue;

    const storeBlocks = findStoreBlocks(worksheet);

    for (const block of storeBlocks) {
      const { storeName } = block;
      let code: string, systemName: string, isMappingFailed: boolean;

      const entry = mapping[storeName];
      if (!entry) {
        mappingFailures.push({ day: dayName, storeName });
        code = 'MAPPING_FAILED';
        systemName = '[매핑실패] ' + storeName;
        isMappingFailed = true;
      } else if (!entry.code || !entry.systemName) {
        mappingFailures.push({ day: dayName, storeName });
        code = 'MAPPING_FAILED';
        systemName = '[매핑실패-빈값] ' + storeName;
        isMappingFailed = true;
      } else {
        code = entry.code;
        systemName = entry.systemName;
        isMappingFailed = false;
      }

      const products = extractProducts(worksheet, block);

      for (const product of products) {
        allData.push({
          date: formatDate(date),
          code,
          originalStoreName: storeName,
          storeName: systemName,
          productName: product.productName,
          boxQty: product.boxQty,
          afternoon: product.afternoon,
          isMappingFailed
        });
      }
    }
  }

  const validation = buildValidation(workbook, dayDates, allData, mappingFailures);
  const storeDaily = buildStoreDaily(allData);
  const uniqueMappingFailures = [...new Set(mappingFailures.map(f => f.storeName))];

  return {
    data: allData,
    validation,
    storeDaily,
    mappingFailures: uniqueMappingFailures
  };
}

interface StoreBlock {
  storeName: string;
  row: number;
  colNo: number;
  colAfternoon: number;
  colProduct: number;
  colBox: number;
}

function findStoreBlocks(worksheet: ExcelJS.Worksheet): StoreBlock[] {
  const blocks: StoreBlock[] = [];
  const rowCount = worksheet.rowCount;

  for (let row = 1; row <= rowCount; row++) {
    for (const layout of STORE_BLOCK_LAYOUTS) {
      const cellValue = getCellValue(worksheet, row, layout.nameCol);
      const storeName = extractStoreName(cellValue);

      if (storeName) {
        blocks.push({
          storeName,
          row,
          colNo: layout.nameCol,
          colAfternoon: layout.afternoonCol,
          colProduct: layout.productCol,
          colBox: layout.boxCol
        });
      }
    }
  }

  return blocks;
}

interface Product {
  productName: string;
  boxQty: number;
  afternoon: string;
}

function extractProducts(worksheet: ExcelJS.Worksheet, block: StoreBlock): Product[] {
  const products: Product[] = [];
  const startRow = block.row + PRODUCT_DATA_START_OFFSET;

  for (let row = startRow; row < startRow + MAX_PRODUCTS_PER_BLOCK; row++) {
    const noVal = getCellValue(worksheet, row, block.colNo);
    if (noVal === null || noVal === undefined || isNaN(parseInt(String(noVal)))) {
      break;
    }

    const productName = getCellValue(worksheet, row, block.colProduct);
    if (!productName) continue;

    const boxQty = parseInt(String(getCellValue(worksheet, row, block.colBox))) || 0;
    if (boxQty === 0) continue;

    const afternoonVal = getCellValue(worksheet, row, block.colAfternoon);

    products.push({
      productName: String(productName).trim(),
      boxQty,
      afternoon: afternoonVal ? String(afternoonVal).trim() : ''
    });
  }

  return products;
}

function buildValidation(
  workbook: ExcelJS.Workbook,
  dayDates: Record<string, Date>,
  allData: DataRow[],
  mappingFailures: { day: string; storeName: string }[]
): ValidationRow[] {
  const validationData: ValidationRow[] = [];

  for (const [dayName, date] of Object.entries(dayDates)) {
    const worksheet = workbook.getWorksheet(dayName);
    if (!worksheet) continue;

    const dateStr = formatDate(date);
    const extractedBox = allData
      .filter(row => row.date === dateStr)
      .reduce((sum, row) => sum + row.boxQty, 0);

    const originalBox = parseInt(String(getCellValue(worksheet, BOX_TOTAL_ROW, BOX_TOTAL_COL))) || 0;

    let result: string;
    if (originalBox <= 0) {
      result = '원본 데이터 없음';
    } else if (extractedBox === originalBox) {
      result = '일치';
    } else {
      result = `불일치 (차이: ${extractedBox - originalBox})`;
    }

    const dayFailures = mappingFailures.filter(f => f.day === dayName).map(f => f.storeName);
    const uniqueDayFailures = [...new Set(dayFailures)];
    const mappingFailureRows = allData.filter(row => row.date === dateStr && row.isMappingFailed).length;

    validationData.push({
      date: dateStr,
      dayName,
      extractedBox,
      originalBox,
      result,
      mappingFailureStores: uniqueDayFailures.length,
      mappingFailureRows
    });
  }

  return validationData;
}

function buildStoreDaily(allData: DataRow[]): StoreDailyRow[] {
  const storeDaily: Record<string, StoreDailyRow> = {};

  allData.forEach(row => {
    const key = `${row.date}_${row.code}_${row.storeName}`;

    if (!storeDaily[key]) {
      storeDaily[key] = {
        date: row.date,
        code: row.code,
        storeName: row.storeName,
        boxSum: 0
      };
    }
    storeDaily[key].boxSum += row.boxQty;
  });

  return Object.values(storeDaily);
}

export async function downloadResult(result: ConversionResult, originalFileName: string): Promise<void> {
  const workbook = new ExcelJS.Workbook();

  // 데이터 시트
  const dataSheet = workbook.addWorksheet('데이터');
  dataSheet.columns = [
    { header: '일자', key: 'date', width: 12 },
    { header: '코드', key: 'code', width: 15 },
    { header: '원본 사업장명', key: 'originalStoreName', width: 20 },
    { header: '사업장명', key: 'storeName', width: 25 },
    { header: '품목명', key: 'productName', width: 30 },
    { header: 'Box 입수', key: 'boxQty', width: 10 },
    { header: '오후 진열', key: 'afternoon', width: 10 }
  ];

  result.data.forEach(row => {
    const excelRow = dataSheet.addRow(row);
    if (row.isMappingFailed) {
      excelRow.eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFCCCC' }
        };
      });
    }
  });

  // 헤더 스타일
  dataSheet.getRow(1).font = { bold: true };
  dataSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
  };

  // 검증 시트
  const validationSheet = workbook.addWorksheet('검증');
  const hasMappingFailures = result.validation.some(
    row => row.mappingFailureStores > 0 || row.mappingFailureRows > 0
  );

  if (hasMappingFailures) {
    validationSheet.columns = [
      { header: '일자', key: 'date', width: 12 },
      { header: '요일', key: 'dayName', width: 8 },
      { header: '추출 Box 합계', key: 'extractedBox', width: 15 },
      { header: '원본 Box 합계', key: 'originalBox', width: 15 },
      { header: '검증 결과', key: 'result', width: 20 },
      { header: '매핑실패 매장수', key: 'mappingFailureStores', width: 15 },
      { header: '매핑실패 데이터수', key: 'mappingFailureRows', width: 15 }
    ];
  } else {
    validationSheet.columns = [
      { header: '일자', key: 'date', width: 12 },
      { header: '요일', key: 'dayName', width: 8 },
      { header: '추출 Box 합계', key: 'extractedBox', width: 15 },
      { header: '원본 Box 합계', key: 'originalBox', width: 15 },
      { header: '검증 결과', key: 'result', width: 20 }
    ];
  }

  result.validation.forEach(row => validationSheet.addRow(row));
  validationSheet.getRow(1).font = { bold: true };
  validationSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
  };

  // 매장별 상세 시트
  const storeDailySheet = workbook.addWorksheet('매장별 상세');
  storeDailySheet.columns = [
    { header: '일자', key: 'date', width: 12 },
    { header: '코드', key: 'code', width: 15 },
    { header: '사업장명', key: 'storeName', width: 25 },
    { header: 'Box 합계', key: 'boxSum', width: 12 }
  ];

  result.storeDaily.forEach(row => storeDailySheet.addRow(row));
  storeDailySheet.getRow(1).font = { bold: true };
  storeDailySheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
  };

  // 매핑실패 시트
  if (result.mappingFailures.length > 0) {
    const failureSheet = workbook.addWorksheet('매핑실패 매장 리스트');
    failureSheet.columns = [
      { header: '매장명', key: 'storeName', width: 30 }
    ];
    result.mappingFailures.forEach(name => failureSheet.addRow({ storeName: name }));
    failureSheet.getRow(1).font = { bold: true };
    failureSheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0E0E0' }
    };
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const outputFileName = originalFileName.replace(/\.xlsx?$/i, '_result.xlsx');
  saveAs(blob, outputFileName);
}
