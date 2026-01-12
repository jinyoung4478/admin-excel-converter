// 간식서비스 엑셀 변환 - JavaScript 버전

let originFile = null;
let mappingFile = null;

// 파일 선택 이벤트
document.getElementById('originFile').addEventListener('change', function(e) {
    originFile = e.target.files[0];
    document.getElementById('originFileName').textContent = originFile ? originFile.name : '';
    this.classList.toggle('has-file', !!originFile);
    checkFilesReady();
});

document.getElementById('mappingFile').addEventListener('change', function(e) {
    mappingFile = e.target.files[0];
    document.getElementById('mappingFileName').textContent = mappingFile ? mappingFile.name : '';
    this.classList.toggle('has-file', !!mappingFile);
    checkFilesReady();
});

function checkFilesReady() {
    document.getElementById('convertBtn').disabled = !(originFile && mappingFile);
}

// 변환 버튼 클릭
document.getElementById('convertBtn').addEventListener('click', async function() {
    try {
        showStatus('processing', '<span class="spinner"></span>변환 중...');

        // 파일 읽기
        const originData = await readExcelFile(originFile);
        const mappingData = await readExcelFile(mappingFile);

        // 매핑 테이블 파싱
        const mapping = parseMappingTable(mappingData);

        // 변환 실행
        const result = convertData(originData, mapping);

        // 결과 파일 생성 및 다운로드
        const outputFileName = originFile.name.replace(/\.xlsx?$/i, '_result.xlsx');
        downloadExcel(result, outputFileName);

        showStatus('success', `변환 완료! ${result.data.length}건 추출`);
    } catch (error) {
        console.error(error);
        showStatus('error', '오류: ' + error.message);
    }
});

function showStatus(type, message) {
    const status = document.getElementById('status');
    status.className = 'status ' + type;
    status.innerHTML = message;
}

// 엑셀 파일 읽기
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                resolve(workbook);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
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

// 매장명 추출
function extractStoreName(cellValue) {
    if (!cellValue || typeof cellValue !== 'string') return null;
    if (!cellValue.includes('※')) return null;

    const match = cellValue.match(/※\s*(.+?)\s*:\s*\d*/);
    return match ? match[1].trim() : null;
}

// 날짜 추출 (파일명에서)
function extractDateFromFileName(fileName) {
    // 패턴: (MM.DD~MM.DD)
    const dateMatch = fileName.match(/\((\d+)\.(\d+)~(\d+)\.(\d+)\)/);
    const yearMatch = fileName.match(/(\d+)년\s*(\d+)월/);

    if (dateMatch && yearMatch) {
        let year = parseInt(yearMatch[1]);
        if (year < 100) year += 2000;
        const month = parseInt(dateMatch[1]) - 1; // JS는 0-indexed
        const day = parseInt(dateMatch[2]);
        return new Date(year, month, day);
    }

    return new Date(2026, 0, 12); // 기본값
}

// 셀 값 가져오기
function getCellValue(sheet, row, col) {
    const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
    const cell = sheet[cellAddress];
    return cell ? cell.v : null;
}

// 매장 블록 찾기
function findStoreBlocks(sheet) {
    const blocks = [];
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

    for (let row = 1; row <= range.e.r + 1; row++) {
        // B열 (col 2) 확인
        const bVal = getCellValue(sheet, row, 2);
        const bStore = extractStoreName(bVal);
        if (bStore) {
            blocks.push({
                storeName: bStore,
                row: row,
                colNo: 2,
                colCategory: 4,
                colProduct: 5,
                colBox: 6
            });
        }

        // K열 (col 11) 확인
        const kVal = getCellValue(sheet, row, 11);
        const kStore = extractStoreName(kVal);
        if (kStore) {
            blocks.push({
                storeName: kStore,
                row: row,
                colNo: 11,
                colCategory: 13,
                colProduct: 14,
                colBox: 15
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
        const noVal = getCellValue(sheet, row, block.colNo);

        // 숫자가 아니면 끝
        if (noVal === null || isNaN(parseInt(noVal))) break;

        const productName = getCellValue(sheet, row, block.colProduct);
        if (!productName) continue;

        let boxQty = getCellValue(sheet, row, block.colBox);
        boxQty = parseInt(boxQty) || 0;

        if (boxQty === 0) continue;

        products.push({
            storeName: block.storeName,
            productName: String(productName).trim(),
            boxQty: boxQty
        });
    }

    return products;
}

// 요일 시트에서 총계 추출
function getDaySheetTotals(workbook, dayDates) {
    const results = {};

    for (const [dayName, date] of Object.entries(dayDates)) {
        if (!workbook.SheetNames.includes(dayName)) continue;

        const sheet = workbook.Sheets[dayName];

        // 총 계 (Row 8, Col F)
        let totalBox = getCellValue(sheet, 8, 6);
        totalBox = parseInt(totalBox) || 0;

        // 개별 매장 합계 (Row 35 이후)
        let storeBoxSum = 0;
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

        for (let row = 35; row <= range.e.r + 1; row++) {
            const bVal = getCellValue(sheet, row, 2);
            const kVal = getCellValue(sheet, row, 11);

            if (bVal === '계') {
                const boxVal = getCellValue(sheet, row, 6);
                if (boxVal && parseInt(boxVal) > 0) {
                    storeBoxSum += parseInt(boxVal);
                }
            }

            if (kVal === '계') {
                const boxVal = getCellValue(sheet, row, 15);
                if (boxVal && parseInt(boxVal) > 0) {
                    storeBoxSum += parseInt(boxVal);
                }
            }
        }

        results[dayName] = {
            totalBox: totalBox,
            storeBoxSum: storeBoxSum
        };
    }

    return results;
}

// 메인 변환 함수
function convertData(originWorkbook, mapping) {
    const dayNames = ['월', '화', '수', '목', '금'];
    const baseDate = extractDateFromFileName(originFile.name);

    const dayDates = {};
    dayNames.forEach((day, idx) => {
        const date = new Date(baseDate);
        date.setDate(date.getDate() + idx);
        dayDates[day] = date;
    });

    const allData = [];
    const mappingFailures = [];

    // 각 요일 시트 처리
    for (const [dayName, date] of Object.entries(dayDates)) {
        if (!originWorkbook.SheetNames.includes(dayName)) continue;

        const sheet = originWorkbook.Sheets[dayName];
        const storeBlocks = findStoreBlocks(sheet);

        for (const block of storeBlocks) {
            const storeName = block.storeName;

            let code, systemName;
            if (!mapping[storeName]) {
                mappingFailures.push({ day: dayName, storeName: storeName });
                code = 'MAPPING_FAILED';
                systemName = '[매핑실패] ' + storeName;
            } else {
                code = mapping[storeName].code;
                systemName = mapping[storeName].systemName;
            }

            const products = extractProductsFromBlock(sheet, block);

            for (const product of products) {
                allData.push({
                    '일자': formatDate(date),
                    '코드': code,
                    '사업장명': systemName,
                    '품목명': product.productName,
                    'Box 입수': product.boxQty
                });
            }
        }
    }

    // 검증 데이터 생성
    const daySheetTotals = getDaySheetTotals(originWorkbook, dayDates);
    const validationData = [];

    for (const [dayName, date] of Object.entries(dayDates)) {
        const dateStr = formatDate(date);
        const extractedBox = allData
            .filter(row => row['일자'] === dateStr)
            .reduce((sum, row) => sum + row['Box 입수'], 0);

        const sheetData = daySheetTotals[dayName] || {};
        const originalTotal = sheetData.totalBox || 0;
        const originalStoreSum = sheetData.storeBoxSum || 0;

        let matchResult;
        if (originalStoreSum > 0) {
            matchResult = extractedBox === originalStoreSum ? '일치' : `불일치 (차이: ${extractedBox - originalStoreSum})`;
        } else {
            matchResult = '원본 데이터 없음';
        }

        validationData.push({
            '일자': dateStr,
            '요일': dayName,
            '추출 Box 합계': extractedBox,
            '원본 시트 총 계': originalTotal,
            '원본 시트 개별매장 합계': originalStoreSum,
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

function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// 엑셀 파일 다운로드
function downloadExcel(result, fileName) {
    const workbook = XLSX.utils.book_new();

    // 데이터 시트
    const dataSheet = XLSX.utils.json_to_sheet(result.data);
    XLSX.utils.book_append_sheet(workbook, dataSheet, '데이터');

    // 검증 시트
    const validationSheet = XLSX.utils.json_to_sheet(result.validation);
    XLSX.utils.book_append_sheet(workbook, validationSheet, '검증');

    // 매장별 상세 시트
    const storeSheet = XLSX.utils.json_to_sheet(result.storeDaily);
    XLSX.utils.book_append_sheet(workbook, storeSheet, '매장별 상세');

    // 매핑 실패 시트 (있을 경우)
    if (result.mappingFailures.length > 0) {
        const failSheet = XLSX.utils.json_to_sheet(result.mappingFailures);
        XLSX.utils.book_append_sheet(workbook, failSheet, '매핑실패');
    }

    // 다운로드
    XLSX.writeFile(workbook, fileName);
}
