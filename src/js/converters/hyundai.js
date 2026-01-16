/**
 * 현대차 간식서비스 엑셀 변환 모듈
 * WASM 가속 지원 (fallback: JS)
 */

import { ExcelCore, StatusManager, FileInputManager } from '../core.js?v=6';

// ========== 상수 ==========
const CONVERTER_CONFIG = {
    id: 'hyundai',
    name: '현대차',
    description: '현대차 간식서비스 원본 엑셀을 시스템 입력용 형식으로 변환'
};

const DAY_NAMES = ['월', '화', '수', '목', '금'];

const MAX_PRODUCTS_PER_BLOCK = 25;
const PRODUCT_DATA_START_OFFSET = 4;
const BOX_TOTAL_ROW = 8;
const BOX_TOTAL_COL = 6;

// 엑셀 시트의 매장 블록 컬럼 레이아웃 (1-indexed)
const STORE_BLOCK_LAYOUTS = [
    { nameCol: 2, afternoonCol: 3, productCol: 5, boxCol: 6 },
    { nameCol: 11, afternoonCol: 12, productCol: 14, boxCol: 15 },
    { nameCol: 20, afternoonCol: 21, productCol: 23, boxCol: 24 }
];

const MAPPING_TABLE_COLUMNS = {
    ORIGINAL_NAME: '원본 사업장명',
    CODE: '코드',
    SYSTEM_NAME: '사업장명'
};

const OUTPUT_COLUMNS = {
    DATE: '일자',
    CODE: '코드',
    ORIGINAL_STORE_NAME: '원본 사업장명',
    STORE_NAME: '사업장명',
    PRODUCT_NAME: '품목명',
    BOX_QTY: 'Box 입수',
    AFTERNOON: '오후 진열',
    MAPPING_FAILED_FLAG: '_isMappingFailed'
};

const DOM_IDS = {
    WASM_STATUS: 'hyundai-wasm-status',
    ORIGIN_FILE: 'hyundai-originFile',
    ORIGIN_FILE_NAME: 'hyundai-originFileName',
    MAPPING_FILE: 'hyundai-mappingFile',
    MAPPING_FILE_NAME: 'hyundai-mappingFileName',
    CONVERT_BTN: 'hyundai-convertBtn',
    STATUS: 'hyundai-status'
};

// ========== WASM 모듈 관리 ==========
const WasmModule = {
    instance: null,
    ready: false,

    async initialize() {
        try {
            const wasm = await import('../../wasm/excel_converter_wasm.js');
            await wasm.default();
            this.instance = wasm;
            this.ready = true;
            console.log('WASM module loaded successfully');
            return true;
        } catch (e) {
            console.warn('WASM not available, using JS fallback:', e.message);
            return false;
        }
    },

    convert(originData, mappingData, filename) {
        const result = this.instance.convert_excel(originData, mappingData, filename);
        if (!result.success) {
            throw new Error(result.error || '변환 실패');
        }
        return result;
    }
};

// ========== 파일 유틸리티 ==========
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(new Uint8Array(reader.result));
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// ========== 파싱 유틸리티 ==========
const Parser = {
    /**
     * 셀 값에서 매장명 추출 (※ 매장명: 숫자 형식)
     */
    extractStoreName(cellValue) {
        if (!cellValue || typeof cellValue !== 'string') return null;
        if (!cellValue.includes('※')) return null;

        const match = cellValue.match(/※\s*(.+?)\s*:\s*\d*/);
        return match ? match[1].trim() : null;
    },

    /**
     * 파일명에서 시작 날짜 추출
     */
    extractDateFromFileName(fileName) {
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
    },

    /**
     * 매핑 테이블 파싱
     */
    parseMappingTable(workbook) {
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = ExcelCore.sheetToJson(sheet);

        const mapping = {};
        data.forEach(row => {
            const originalName = row[MAPPING_TABLE_COLUMNS.ORIGINAL_NAME];
            if (originalName) {
                mapping[originalName] = {
                    code: row[MAPPING_TABLE_COLUMNS.CODE] || '',
                    systemName: row[MAPPING_TABLE_COLUMNS.SYSTEM_NAME] || ''
                };
            }
        });

        return mapping;
    }
};

// ========== 매장 블록 추출 ==========
const StoreBlockExtractor = {
    /**
     * 시트에서 모든 매장 블록 찾기
     */
    findAll(sheet) {
        const blocks = [];
        const range = ExcelCore.decodeRange(sheet['!ref'] || 'A1');

        for (let row = 1; row <= range.e.r + 1; row++) {
            this._extractBlocksFromRow(sheet, row, blocks);
        }

        return blocks;
    },

    _extractBlocksFromRow(sheet, row, blocks) {
        for (const layout of STORE_BLOCK_LAYOUTS) {
            const cellValue = ExcelCore.getCellValue(sheet, row, layout.nameCol);
            const storeName = Parser.extractStoreName(cellValue);

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
    },

    /**
     * 블록에서 상품 목록 추출
     */
    extractProducts(sheet, block) {
        const products = [];
        const startRow = block.row + PRODUCT_DATA_START_OFFSET;

        for (let row = startRow; row < startRow + MAX_PRODUCTS_PER_BLOCK; row++) {
            const product = this._extractProductFromRow(sheet, block, row);
            if (!product) break;
            if (product.boxQty > 0) {
                products.push(product);
            }
        }

        return products;
    },

    _extractProductFromRow(sheet, block, row) {
        const noVal = ExcelCore.getCellValue(sheet, row, block.colNo);
        if (noVal === null || isNaN(parseInt(noVal))) {
            return null;
        }

        const productName = ExcelCore.getCellValue(sheet, row, block.colProduct);
        if (!productName) {
            return { boxQty: 0 };
        }

        const boxQty = parseInt(ExcelCore.getCellValue(sheet, row, block.colBox)) || 0;
        const afternoonVal = ExcelCore.getCellValue(sheet, row, block.colAfternoon);

        return {
            storeName: block.storeName,
            productName: String(productName).trim(),
            boxQty,
            afternoon: afternoonVal ? String(afternoonVal).trim() : ''
        };
    }
};

// ========== 데이터 변환 ==========
const DataConverter = {
    /**
     * 원본 워크북을 결과 데이터로 변환
     */
    convert(originWorkbook, mapping, fileName) {
        const baseDate = Parser.extractDateFromFileName(fileName);
        const dayDates = this._buildDayDates(baseDate);

        const allData = [];
        const mappingFailures = [];

        for (const [dayName, date] of Object.entries(dayDates)) {
            if (!originWorkbook.SheetNames.includes(dayName)) continue;

            const sheet = originWorkbook.Sheets[dayName];
            this._processSheet(sheet, dayName, date, mapping, allData, mappingFailures);
        }

        return {
            data: allData,
            validation: this._buildValidation(originWorkbook, dayDates, allData, mappingFailures),
            storeDaily: this._buildStoreDaily(allData),
            mappingFailures: this._buildUniqueMappingFailures(mappingFailures)
        };
    },

    _buildDayDates(baseDate) {
        const dayDates = {};
        DAY_NAMES.forEach((day, idx) => {
            const date = new Date(baseDate);
            date.setDate(date.getDate() + idx);
            dayDates[day] = date;
        });
        return dayDates;
    },

    _processSheet(sheet, dayName, date, mapping, allData, mappingFailures) {
        const storeBlocks = StoreBlockExtractor.findAll(sheet);

        for (const block of storeBlocks) {
            const mappingResult = this._resolveMappingForStore(block.storeName, mapping);

            if (mappingResult.failed) {
                mappingFailures.push({ day: dayName, storeName: block.storeName });
            }

            const products = StoreBlockExtractor.extractProducts(sheet, block);

            for (const product of products) {
                allData.push(this._createDataRow(date, block.storeName, mappingResult, product));
            }
        }
    },

    _resolveMappingForStore(storeName, mapping) {
        const entry = mapping[storeName];

        if (!entry) {
            return {
                code: 'MAPPING_FAILED',
                systemName: '[매핑실패] ' + storeName,
                failed: true
            };
        }

        if (!entry.code || !entry.systemName) {
            return {
                code: 'MAPPING_FAILED',
                systemName: '[매핑실패-빈값] ' + storeName,
                failed: true
            };
        }

        return {
            code: entry.code,
            systemName: entry.systemName,
            failed: false
        };
    },

    _createDataRow(date, storeName, mappingResult, product) {
        return {
            [OUTPUT_COLUMNS.DATE]: ExcelCore.formatDate(date),
            [OUTPUT_COLUMNS.CODE]: mappingResult.code,
            [OUTPUT_COLUMNS.ORIGINAL_STORE_NAME]: storeName,
            [OUTPUT_COLUMNS.STORE_NAME]: mappingResult.systemName,
            [OUTPUT_COLUMNS.PRODUCT_NAME]: product.productName,
            [OUTPUT_COLUMNS.BOX_QTY]: product.boxQty,
            [OUTPUT_COLUMNS.AFTERNOON]: product.afternoon,
            [OUTPUT_COLUMNS.MAPPING_FAILED_FLAG]: mappingResult.failed
        };
    },

    _buildValidation(originWorkbook, dayDates, allData, mappingFailures) {
        const validationData = [];

        for (const [dayName, date] of Object.entries(dayDates)) {
            if (!originWorkbook.SheetNames.includes(dayName)) continue;

            const sheet = originWorkbook.Sheets[dayName];
            const dateStr = ExcelCore.formatDate(date);

            const extractedBox = allData
                .filter(row => row[OUTPUT_COLUMNS.DATE] === dateStr)
                .reduce((sum, row) => sum + row[OUTPUT_COLUMNS.BOX_QTY], 0);

            const originalBox = this._getOriginalBoxTotal(sheet);
            const matchResult = this._evaluateMatch(extractedBox, originalBox);

            const dayFailures = mappingFailures
                .filter(f => f.day === dayName)
                .map(f => f.storeName);
            const uniqueDayFailures = [...new Set(dayFailures)];

            const mappingFailureRows = allData
                .filter(row => row[OUTPUT_COLUMNS.DATE] === dateStr && row['매핑실패'] === 'Y')
                .length;

            validationData.push({
                '일자': dateStr,
                '요일': dayName,
                '추출 Box 합계': extractedBox,
                '원본 Box 합계': originalBox,
                '검증 결과': matchResult,
                '매핑실패 매장수': uniqueDayFailures.length,
                '매핑실패 데이터수': mappingFailureRows
            });
        }

        return validationData;
    },

    _getOriginalBoxTotal(sheet) {
        const val = ExcelCore.getCellValue(sheet, BOX_TOTAL_ROW, BOX_TOTAL_COL);
        return parseInt(val) || 0;
    },

    _evaluateMatch(extracted, original) {
        if (original <= 0) return '원본 데이터 없음';
        if (extracted === original) return '일치';
        return `불일치 (차이: ${extracted - original})`;
    },

    _buildStoreDaily(allData) {
        const storeDaily = {};

        allData.forEach(row => {
            const key = `${row[OUTPUT_COLUMNS.DATE]}_${row[OUTPUT_COLUMNS.CODE]}_${row[OUTPUT_COLUMNS.STORE_NAME]}`;

            if (!storeDaily[key]) {
                storeDaily[key] = {
                    '일자': row[OUTPUT_COLUMNS.DATE],
                    '코드': row[OUTPUT_COLUMNS.CODE],
                    '사업장명': row[OUTPUT_COLUMNS.STORE_NAME],
                    'Box 합계': 0
                };
            }
            storeDaily[key]['Box 합계'] += row[OUTPUT_COLUMNS.BOX_QTY];
        });

        return Object.values(storeDaily);
    },

    _buildUniqueMappingFailures(mappingFailures) {
        const uniqueNames = [...new Set(mappingFailures.map(f => f.storeName))];
        return uniqueNames.map(s => ({ '매장명': s }));
    }
};

// ========== 결과 워크북 생성 ==========
const ResultWorkbookBuilder = {
    build(result) {
        const workbook = ExcelCore.createWorkbook();

        ExcelCore.addSheet(workbook, result.data, '데이터');
        ExcelCore.addSheet(workbook, this._prepareValidationData(result.validation), '검증');
        ExcelCore.addSheet(workbook, result.storeDaily, '매장별 상세');

        if (result.mappingFailures.length > 0) {
            ExcelCore.addSheet(workbook, result.mappingFailures, '매핑실패 매장 리스트');
        }

        return workbook;
    },

    _prepareValidationData(validation) {
        const hasMappingFailures = validation.some(
            row => row['매핑실패 매장수'] > 0 || row['매핑실패 데이터수'] > 0
        );

        if (hasMappingFailures) {
            return validation;
        }

        return validation.map(row => {
            const { '매핑실패 매장수': _, '매핑실패 데이터수': __, ...rest } = row;
            return rest;
        });
    }
};

// ========== WASM 결과 변환 ==========
const WasmResultMapper = {
    toJsFormat(wasmResult) {
        return {
            data: wasmResult.data.map(r => ({
                [OUTPUT_COLUMNS.DATE]: r.date,
                [OUTPUT_COLUMNS.CODE]: r.code,
                [OUTPUT_COLUMNS.ORIGINAL_STORE_NAME]: r.original_store_name,
                [OUTPUT_COLUMNS.STORE_NAME]: r.store_name,
                [OUTPUT_COLUMNS.PRODUCT_NAME]: r.product_name,
                [OUTPUT_COLUMNS.BOX_QTY]: r.box_qty,
                [OUTPUT_COLUMNS.AFTERNOON]: r.afternoon || '',
                [OUTPUT_COLUMNS.MAPPING_FAILED_FLAG]: r.mapping_failed === 'Y'
            })),
            validation: wasmResult.validation.map(r => ({
                '일자': r.date,
                '요일': r.day_name,
                '추출 Box 합계': r.extracted_box,
                '원본 Box 합계': r.original_box,
                '검증 결과': r.result,
                '매핑실패 매장수': r.mapping_failure_stores,
                '매핑실패 데이터수': r.mapping_failure_rows
            })),
            storeDaily: wasmResult.store_daily.map(r => ({
                '일자': r.date,
                '코드': r.code,
                '사업장명': r.store_name,
                'Box 합계': r.box_sum
            })),
            mappingFailures: wasmResult.mapping_failures.map(s => ({ '매장명': s }))
        };
    }
};

// ========== 변환 실행기 ==========
const ConversionExecutor = {
    async execute(originFile, mappingFile) {
        const startTime = performance.now();

        const result = WasmModule.ready
            ? await this._executeWithWasm(originFile, mappingFile)
            : await this._executeWithJs(originFile, mappingFile);

        const workbook = ResultWorkbookBuilder.build(result.data);
        const outputFileName = originFile.name.replace(/\.xlsx?$/i, '_result.xlsx');
        await ExcelCore.downloadExcel(workbook, outputFileName);

        const elapsed = ((performance.now() - startTime) / 1000).toFixed(2);

        return {
            count: result.data.data.length,
            elapsed,
            mode: result.mode
        };
    },

    async _executeWithWasm(originFile, mappingFile) {
        const [originData, mappingData] = await Promise.all([
            readFileAsArrayBuffer(originFile),
            readFileAsArrayBuffer(mappingFile)
        ]);

        const wasmResult = WasmModule.convert(originData, mappingData, originFile.name);
        return {
            data: WasmResultMapper.toJsFormat(wasmResult),
            mode: 'WASM'
        };
    },

    async _executeWithJs(originFile, mappingFile) {
        const [originWorkbook, mappingWorkbook] = await Promise.all([
            ExcelCore.readFile(originFile),
            ExcelCore.readFile(mappingFile)
        ]);

        const mapping = Parser.parseMappingTable(mappingWorkbook);
        const result = DataConverter.convert(originWorkbook, mapping, originFile.name);

        return {
            data: result,
            mode: 'JS'
        };
    }
};

// ========== UI ==========
const UI = {
    state: {
        originFile: null,
        mappingFile: null
    },

    initialize(container) {
        this._renderTemplate(container);
        this._injectStyles();
        this._initWasmStatus();
        this._setupFileInputs();
        this._setupConvertButton();
    },

    _renderTemplate(container) {
        container.innerHTML = `
            <div class="converter-form">
                <div class="wasm-status" id="${DOM_IDS.WASM_STATUS}"></div>

                <div class="file-input-wrapper">
                    <label>1. 원본 엑셀 파일</label>
                    <input type="file" id="${DOM_IDS.ORIGIN_FILE}" class="file-input" accept=".xlsx,.xls">
                    <div class="file-name" id="${DOM_IDS.ORIGIN_FILE_NAME}"></div>
                </div>

                <div class="file-input-wrapper">
                    <label>2. 매핑 테이블 파일</label>
                    <input type="file" id="${DOM_IDS.MAPPING_FILE}" class="file-input" accept=".xlsx,.xls">
                    <div class="file-name" id="${DOM_IDS.MAPPING_FILE_NAME}"></div>
                </div>

                <button class="btn" id="${DOM_IDS.CONVERT_BTN}" disabled>변환하기</button>

                <div class="status" id="${DOM_IDS.STATUS}"></div>

                <div class="info">
                    <h3>사용 방법</h3>
                    <ul>
                        <li>원본 엑셀 파일과 매핑 테이블을 선택하세요</li>
                        <li>변환하기 버튼을 클릭하면 결과 파일이 다운로드됩니다</li>
                        <li>결과 파일명: 원본파일명_result.xlsx</li>
                    </ul>
                    <h3 style="margin-top: 15px;">매핑 테이블 형식</h3>
                    <ul>
                        <li>컬럼: 코드, 원본 사업장명, 사업장명</li>
                    </ul>
                </div>
            </div>
        `;
    },

    _injectStyles() {
        if (document.getElementById('hyundai-wasm-styles')) return;

        const style = document.createElement('style');
        style.id = 'hyundai-wasm-styles';
        style.textContent = `
            .wasm-status {
                padding: 8px 12px;
                border-radius: 6px;
                font-size: 12px;
                margin-bottom: 15px;
                display: flex;
                align-items: center;
                gap: 6px;
            }
            .wasm-status.loading { background: #fff3e0; color: #e65100; }
            .wasm-status.ready { background: #e8f5e9; color: #2e7d32; }
            .wasm-status.fallback { background: #f5f5f5; color: #666; }
            .wasm-status::before {
                content: '';
                width: 8px;
                height: 8px;
                border-radius: 50%;
                display: inline-block;
            }
            .wasm-status.loading::before { background: #e65100; animation: pulse 1s infinite; }
            .wasm-status.ready::before { background: #4caf50; }
            .wasm-status.fallback::before { background: #9e9e9e; }
            @keyframes pulse {
                0%, 100% { opacity: 1; }
                50% { opacity: 0.5; }
            }
        `;
        document.head.appendChild(style);
    },

    async _initWasmStatus() {
        const wasmStatus = document.getElementById(DOM_IDS.WASM_STATUS);

        wasmStatus.className = 'wasm-status loading';
        wasmStatus.textContent = 'WASM 모듈 로딩 중...';

        const success = await WasmModule.initialize();

        if (success) {
            wasmStatus.className = 'wasm-status ready';
            wasmStatus.textContent = 'WASM 가속 활성화';
        } else {
            wasmStatus.className = 'wasm-status fallback';
            wasmStatus.textContent = 'JS 모드 (WASM 미지원)';
        }
    },

    _setupFileInputs() {
        FileInputManager.setup(
            DOM_IDS.ORIGIN_FILE,
            DOM_IDS.ORIGIN_FILE_NAME,
            (file) => {
                this.state.originFile = file;
                this._updateButtonState();
            }
        );

        FileInputManager.setup(
            DOM_IDS.MAPPING_FILE,
            DOM_IDS.MAPPING_FILE_NAME,
            (file) => {
                this.state.mappingFile = file;
                this._updateButtonState();
            }
        );
    },

    _updateButtonState() {
        const btn = document.getElementById(DOM_IDS.CONVERT_BTN);
        btn.disabled = !(this.state.originFile && this.state.mappingFile);
    },

    _setupConvertButton() {
        const btn = document.getElementById(DOM_IDS.CONVERT_BTN);

        btn.addEventListener('click', async () => {
            try {
                StatusManager.processing(DOM_IDS.STATUS, '변환 중...');

                const result = await ConversionExecutor.execute(
                    this.state.originFile,
                    this.state.mappingFile
                );

                StatusManager.success(
                    DOM_IDS.STATUS,
                    `변환 완료! ${result.count}건 추출 (${result.elapsed}초, ${result.mode})`
                );
            } catch (error) {
                console.error(error);
                StatusManager.error(DOM_IDS.STATUS, '오류: ' + error.message);
            }
        });
    }
};

// ========== 컨버터 내보내기 ==========
export default {
    ...CONVERTER_CONFIG,
    init: (container) => UI.initialize(container)
};
