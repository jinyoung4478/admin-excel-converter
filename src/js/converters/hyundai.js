/**
 * 현대차 간식서비스 엑셀 변환 모듈
 * WASM 가속 지원 (fallback: JS)
 */

import { ExcelCore, StatusManager, FileInputManager } from '../core.js';

// 컨버터 설정
const config = {
    id: 'hyundai',
    name: '현대차',
    description: '현대차 간식서비스 원본 엑셀을 시스템 입력용 형식으로 변환'
};

// WASM 모듈
let wasmModule = null;
let wasmReady = false;

// WASM 초기화
async function initWasm() {
    try {
        const wasm = await import('../../wasm/excel_converter_wasm.js');
        await wasm.default();
        wasmModule = wasm;
        wasmReady = true;
        console.log('WASM module loaded successfully');
        return true;
    } catch (e) {
        console.warn('WASM not available, using JS fallback:', e.message);
        return false;
    }
}

// 상태
let originFile = null;
let mappingFile = null;

// 파일을 ArrayBuffer로 읽기
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(new Uint8Array(reader.result));
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// ========== WASM 변환 ==========
async function convertWithWasm(originData, mappingData, filename) {
    // WASM으로 파싱 (JSON 결과 반환)
    const result = wasmModule.convert_excel(originData, mappingData, filename);

    if (!result.success) {
        throw new Error(result.error || '변환 실패');
    }

    return result;
}

// ========== JS Fallback ==========

// 매장명 추출
function extractStoreName(cellValue) {
    if (!cellValue || typeof cellValue !== 'string') return null;
    if (!cellValue.includes('※')) return null;

    const match = cellValue.match(/※\s*(.+?)\s*:\s*\d*/);
    return match ? match[1].trim() : null;
}

// 날짜 추출 (파일명에서)
function extractDateFromFileName(fileName) {
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

// 매핑 테이블 파싱
function parseMappingTable(workbook) {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    const mapping = {};
    data.forEach(row => {
        const originalName = row['원본 사업장명'];
        if (originalName) {
            mapping[originalName] = {
                code: row['코드'] || '',
                systemName: row['사업장명'] || ''
            };
        }
    });

    return mapping;
}

// 매장 블록 찾기
function findStoreBlocks(sheet) {
    const blocks = [];
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

    for (let row = 1; row <= range.e.r + 1; row++) {
        const bVal = ExcelCore.getCellValue(sheet, row, 2);
        const bStore = extractStoreName(bVal);
        if (bStore) {
            blocks.push({
                storeName: bStore,
                row: row,
                colNo: 2,
                colAfternoon: 3,
                colProduct: 5,
                colBox: 6
            });
        }

        const kVal = ExcelCore.getCellValue(sheet, row, 11);
        const kStore = extractStoreName(kVal);
        if (kStore) {
            blocks.push({
                storeName: kStore,
                row: row,
                colNo: 11,
                colAfternoon: 12,
                colProduct: 14,
                colBox: 15
            });
        }

        const tVal = ExcelCore.getCellValue(sheet, row, 20);
        const tStore = extractStoreName(tVal);
        if (tStore) {
            blocks.push({
                storeName: tStore,
                row: row,
                colNo: 20,
                colAfternoon: 21,
                colProduct: 23,
                colBox: 24
            });
        }
    }

    return blocks;
}

// 블록에서 상품 추출
function extractProductsFromBlock(sheet, block, maxProducts = 25) {
    const products = [];
    const startRow = block.row + 4;

    for (let row = startRow; row < startRow + maxProducts; row++) {
        const noVal = ExcelCore.getCellValue(sheet, row, block.colNo);
        if (noVal === null || isNaN(parseInt(noVal))) break;

        const productName = ExcelCore.getCellValue(sheet, row, block.colProduct);
        if (!productName) continue;

        let boxQty = ExcelCore.getCellValue(sheet, row, block.colBox);
        boxQty = parseInt(boxQty) || 0;
        if (boxQty === 0) continue;

        // 오후 진열 값 추출
        const afternoonVal = ExcelCore.getCellValue(sheet, row, block.colAfternoon);
        const afternoon = afternoonVal ? String(afternoonVal).trim() : '';

        products.push({
            storeName: block.storeName,
            productName: String(productName).trim(),
            boxQty: boxQty,
            afternoon: afternoon
        });
    }

    return products;
}

// 각 요일 시트에서 F8 셀의 Box 합계 추출
function getOriginalBoxTotal(sheet) {
    // F8 셀 (1-indexed: row=8, col=6)
    const val = ExcelCore.getCellValue(sheet, 8, 6);
    return parseInt(val) || 0;
}

// JS 변환 함수
function convertDataJS(originWorkbook, mapping, fileName) {
    const dayNames = ['월', '화', '수', '목', '금'];
    const baseDate = extractDateFromFileName(fileName);

    const dayDates = {};
    dayNames.forEach((day, idx) => {
        const date = new Date(baseDate);
        date.setDate(date.getDate() + idx);
        dayDates[day] = date;
    });

    const allData = [];
    const mappingFailures = [];

    for (const [dayName, date] of Object.entries(dayDates)) {
        if (!originWorkbook.SheetNames.includes(dayName)) continue;

        const sheet = originWorkbook.Sheets[dayName];
        const storeBlocks = findStoreBlocks(sheet);

        for (const block of storeBlocks) {
            const storeName = block.storeName;

            let code, systemName;
            if (!mapping[storeName]) {
                mappingFailures.push({ day: dayName, storeName });
                code = 'MAPPING_FAILED';
                systemName = '[매핑실패] ' + storeName;
            } else {
                code = mapping[storeName].code;
                systemName = mapping[storeName].systemName;
            }

            const products = extractProductsFromBlock(sheet, block);

            for (const product of products) {
                allData.push({
                    '일자': ExcelCore.formatDate(date),
                    '코드': code,
                    '사업장명': systemName,
                    '품목명': product.productName,
                    'Box 입수': product.boxQty,
                    '오후 진열': product.afternoon
                });
            }
        }
    }

    // 검증 데이터 (각 요일 시트의 F8 셀 기준)
    const validationData = [];

    for (const [dayName, date] of Object.entries(dayDates)) {
        if (!originWorkbook.SheetNames.includes(dayName)) continue;

        const sheet = originWorkbook.Sheets[dayName];
        const dateStr = ExcelCore.formatDate(date);
        const extractedBox = allData
            .filter(row => row['일자'] === dateStr)
            .reduce((sum, row) => sum + row['Box 입수'], 0);

        const originalBox = getOriginalBoxTotal(sheet);

        let matchResult;
        if (originalBox > 0) {
            matchResult = extractedBox === originalBox ? '일치' : `불일치 (차이: ${extractedBox - originalBox})`;
        } else {
            matchResult = '원본 데이터 없음';
        }

        validationData.push({
            '일자': dateStr,
            '요일': dayName,
            '추출 Box 합계': extractedBox,
            '원본 Box 합계': originalBox,
            '검증 결과': matchResult
        });
    }

    // 매장별 상세
    const storeDaily = {};
    allData.forEach(row => {
        const key = `${row['일자']}_${row['코드']}_${row['사업장명']}`;
        if (!storeDaily[key]) {
            storeDaily[key] = {
                '일자': row['일자'],
                '코드': row['코드'],
                '사업장명': row['사업장명'],
                'Box 합계': 0
            };
        }
        storeDaily[key]['Box 합계'] += row['Box 입수'];
    });

    return {
        data: allData,
        validation: validationData,
        storeDaily: Object.values(storeDaily),
        mappingFailures: [...new Set(mappingFailures.map(f => f.storeName))].map(s => ({ '매장명': s }))
    };
}

// 결과 엑셀 생성 (JS)
function createResultWorkbookJS(result) {
    const workbook = ExcelCore.createWorkbook();

    ExcelCore.addSheet(workbook, result.data, '데이터');
    ExcelCore.addSheet(workbook, result.validation, '검증');
    ExcelCore.addSheet(workbook, result.storeDaily, '매장별 상세');

    if (result.mappingFailures.length > 0) {
        ExcelCore.addSheet(workbook, result.mappingFailures, '매핑실패');
    }

    return workbook;
}

// ========== 메인 변환 로직 ==========
async function convert() {
    const startTime = performance.now();

    if (wasmReady) {
        // WASM 사용 (파싱은 Rust, Excel 생성은 JS)
        const [originData, mappingData] = await Promise.all([
            readFileAsArrayBuffer(originFile),
            readFileAsArrayBuffer(mappingFile)
        ]);

        const result = await convertWithWasm(originData, mappingData, originFile.name);

        // WASM 결과를 JS 형식으로 변환하여 Excel 생성
        const jsResult = {
            data: result.data.map(r => ({
                '일자': r.date,
                '코드': r.code,
                '사업장명': r.store_name,
                '품목명': r.product_name,
                'Box 입수': r.box_qty,
                '오후 진열': r.afternoon || ''
            })),
            validation: result.validation.map(r => ({
                '일자': r.date,
                '요일': r.day_name,
                '추출 Box 합계': r.extracted_box,
                '원본 Box 합계': r.original_box,
                '검증 결과': r.result
            })),
            storeDaily: result.store_daily.map(r => ({
                '일자': r.date,
                '코드': r.code,
                '사업장명': r.store_name,
                'Box 합계': r.box_sum
            })),
            mappingFailures: result.mapping_failures.map(s => ({ '매장명': s }))
        };

        const workbook = createResultWorkbookJS(jsResult);
        const outputFileName = originFile.name.replace(/\.xlsx?$/i, '_result.xlsx');
        ExcelCore.downloadExcel(workbook, outputFileName);

        const elapsed = ((performance.now() - startTime) / 1000).toFixed(2);
        return { count: result.data.length, elapsed, mode: 'WASM' };
    } else {
        // JS Fallback
        const [originWorkbook, mappingWorkbook] = await Promise.all([
            ExcelCore.readFile(originFile),
            ExcelCore.readFile(mappingFile)
        ]);

        const mapping = parseMappingTable(mappingWorkbook);
        const result = convertDataJS(originWorkbook, mapping, originFile.name);
        const workbook = createResultWorkbookJS(result);

        const outputFileName = originFile.name.replace(/\.xlsx?$/i, '_result.xlsx');
        ExcelCore.downloadExcel(workbook, outputFileName);

        const elapsed = ((performance.now() - startTime) / 1000).toFixed(2);
        return { count: result.data.length, elapsed, mode: 'JS' };
    }
}

// UI 초기화
function initUI(container) {
    container.innerHTML = `
        <div class="converter-form">
            <div class="wasm-status" id="hyundai-wasm-status"></div>

            <div class="file-input-wrapper">
                <label>1. 원본 엑셀 파일</label>
                <input type="file" id="hyundai-originFile" class="file-input" accept=".xlsx,.xls">
                <div class="file-name" id="hyundai-originFileName"></div>
            </div>

            <div class="file-input-wrapper">
                <label>2. 매핑 테이블 파일</label>
                <input type="file" id="hyundai-mappingFile" class="file-input" accept=".xlsx,.xls">
                <div class="file-name" id="hyundai-mappingFileName"></div>
            </div>

            <button class="btn" id="hyundai-convertBtn" disabled>변환하기</button>

            <div class="status" id="hyundai-status"></div>

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

    // WASM 상태 스타일
    const style = document.createElement('style');
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
        .wasm-status.loading {
            background: #fff3e0;
            color: #e65100;
        }
        .wasm-status.ready {
            background: #e8f5e9;
            color: #2e7d32;
        }
        .wasm-status.fallback {
            background: #f5f5f5;
            color: #666;
        }
        .wasm-status::before {
            content: '';
            width: 8px;
            height: 8px;
            border-radius: 50%;
            display: inline-block;
        }
        .wasm-status.loading::before {
            background: #e65100;
            animation: pulse 1s infinite;
        }
        .wasm-status.ready::before {
            background: #4caf50;
        }
        .wasm-status.fallback::before {
            background: #9e9e9e;
        }
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }
    `;
    document.head.appendChild(style);

    // WASM 초기화
    const wasmStatus = document.getElementById('hyundai-wasm-status');
    wasmStatus.className = 'wasm-status loading';
    wasmStatus.textContent = 'WASM 모듈 로딩 중...';

    initWasm().then(success => {
        if (success) {
            wasmStatus.className = 'wasm-status ready';
            wasmStatus.textContent = 'WASM 가속 활성화';
        } else {
            wasmStatus.className = 'wasm-status fallback';
            wasmStatus.textContent = 'JS 모드 (WASM 미지원)';
        }
    });

    // 파일 입력 설정
    FileInputManager.setup(
        'hyundai-originFile',
        'hyundai-originFileName',
        (file) => {
            originFile = file;
            updateButtonState();
        }
    );

    FileInputManager.setup(
        'hyundai-mappingFile',
        'hyundai-mappingFileName',
        (file) => {
            mappingFile = file;
            updateButtonState();
        }
    );

    function updateButtonState() {
        const btn = document.getElementById('hyundai-convertBtn');
        btn.disabled = !(originFile && mappingFile);
    }

    // 변환 버튼
    document.getElementById('hyundai-convertBtn').addEventListener('click', async () => {
        try {
            StatusManager.processing('hyundai-status', '변환 중...');

            const result = await convert();

            StatusManager.success(
                'hyundai-status',
                `변환 완료! ${result.count}건 추출 (${result.elapsed}초, ${result.mode})`
            );
        } catch (error) {
            console.error(error);
            StatusManager.error('hyundai-status', '오류: ' + error.message);
        }
    });
}

// 컨버터 내보내기
export default {
    ...config,
    init: initUI
};
