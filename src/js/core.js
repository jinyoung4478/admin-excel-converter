/**
 * 엑셀 변환 공통 유틸리티 (ExcelJS 기반)
 */

const ExcelCore = {
    // 엑셀 파일 읽기
    async readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    const arrayBuffer = e.target.result;
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);
                    
                    // SheetJS 호환 형식으로 변환
                    const result = {
                        SheetNames: [],
                        Sheets: {}
                    };
                    
                    workbook.eachSheet((worksheet, sheetId) => {
                        const sheetName = worksheet.name;
                        result.SheetNames.push(sheetName);
                        
                        // 시트 데이터를 SheetJS 호환 형식으로 변환
                        const sheetData = {};
                        let maxRow = 0;
                        let maxCol = 0;
                        let minRow = Infinity;
                        let minCol = Infinity;
                        
                        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                            maxRow = Math.max(maxRow, rowNumber);
                            minRow = Math.min(minRow, rowNumber);
                            
                            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                                maxCol = Math.max(maxCol, colNumber);
                                minCol = Math.min(minCol, colNumber);
                                
                                const cellAddress = ExcelCore._encodeCell(rowNumber - 1, colNumber - 1);
                                sheetData[cellAddress] = {
                                    v: cell.value,
                                    t: typeof cell.value === 'number' ? 'n' : 's'
                                };
                                
                                // 수식 결과값 처리
                                if (cell.value && typeof cell.value === 'object') {
                                    if (cell.value.result !== undefined) {
                                        sheetData[cellAddress].v = cell.value.result;
                                    } else if (cell.value.text !== undefined) {
                                        sheetData[cellAddress].v = cell.value.text;
                                    }
                                }
                            });
                        });
                        
                        // 범위 설정
                        if (minRow !== Infinity && minCol !== Infinity) {
                            const startCell = ExcelCore._encodeCell(minRow - 1, minCol - 1);
                            const endCell = ExcelCore._encodeCell(maxRow - 1, maxCol - 1);
                            sheetData['!ref'] = `${startCell}:${endCell}`;
                        } else {
                            sheetData['!ref'] = 'A1:A1';
                        }
                        
                        result.Sheets[sheetName] = sheetData;
                    });
                    
                    resolve(result);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    },

    // 셀 주소 인코딩 (0-indexed row, col -> "A1" 형식)
    _encodeCell(row, col) {
        let colStr = '';
        let c = col;
        do {
            colStr = String.fromCharCode(65 + (c % 26)) + colStr;
            c = Math.floor(c / 26) - 1;
        } while (c >= 0);
        return colStr + (row + 1);
    },

    // 셀 값 가져오기 (1-indexed)
    getCellValue(sheet, row, col) {
        const cellAddress = ExcelCore._encodeCell(row - 1, col - 1);
        const cell = sheet[cellAddress];
        return cell ? cell.v : null;
    },

    // 시트를 JSON 배열로 변환 (SheetJS 호환)
    sheetToJson(sheet) {
        const ref = sheet['!ref'];
        if (!ref) return [];
        
        const range = ExcelCore.decodeRange(ref);
        const data = [];
        const headers = [];
        
        // 헤더 읽기 (첫 번째 행)
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddr = ExcelCore._encodeCell(range.s.r, col);
            const cell = sheet[cellAddr];
            headers.push(cell ? String(cell.v) : `Column${col}`);
        }
        
        // 데이터 읽기
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const rowData = {};
            let hasData = false;
            
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddr = ExcelCore._encodeCell(row, col);
                const cell = sheet[cellAddr];
                const header = headers[col - range.s.c];
                
                if (cell && cell.v !== undefined && cell.v !== null) {
                    rowData[header] = cell.v;
                    hasData = true;
                }
            }
            
            if (hasData) {
                data.push(rowData);
            }
        }
        
        return data;
    },

    // 범위 디코딩 (SheetJS 호환)
    decodeRange(ref) {
        const parts = ref.split(':');
        const start = ExcelCore._decodeCell(parts[0]);
        const end = parts[1] ? ExcelCore._decodeCell(parts[1]) : start;
        
        return {
            s: { r: start.r, c: start.c },
            e: { r: end.r, c: end.c }
        };
    },

    // 셀 주소 디코딩 ("A1" -> {r: 0, c: 0})
    _decodeCell(addr) {
        let col = 0;
        let row = 0;
        let i = 0;
        
        // 열 문자 파싱
        while (i < addr.length && /[A-Z]/i.test(addr[i])) {
            col = col * 26 + (addr.charCodeAt(i) & 0x1F);
            i++;
        }
        col--;
        
        // 행 숫자 파싱
        row = parseInt(addr.substring(i)) - 1;
        
        return { r: row, c: col };
    },

    // 날짜 포맷팅
    formatDate(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    },

    // 엑셀 파일 다운로드 (스타일 지원)
    async downloadExcel(workbookData, fileName, styleOptions = {}) {
        const workbook = new ExcelJS.Workbook();
        
        // 각 시트 추가
        for (const sheetInfo of workbookData.sheets) {
            const worksheet = workbook.addWorksheet(sheetInfo.name);
            
            if (sheetInfo.data && sheetInfo.data.length > 0) {
                // 헤더 추가 (_로 시작하는 내부 필드 제외)
                const allKeys = Object.keys(sheetInfo.data[0]);
                const headers = allKeys.filter(h => !h.startsWith('_'));
                worksheet.addRow(headers);
                
                // 헤더 스타일
                const headerRow = worksheet.getRow(1);
                headerRow.font = { bold: true };
                headerRow.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE0E0E0' }
                };
                
                // 데이터 추가
                sheetInfo.data.forEach((row, rowIndex) => {
                    const values = headers.map(h => row[h]);
                    const excelRow = worksheet.addRow(values);
                    
                    // 매핑 실패 행 스타일 적용 (_isMappingFailed 플래그 확인)
                    if (sheetInfo.name === '데이터' && row['_isMappingFailed']) {
                        excelRow.eachCell((cell) => {
                            cell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FFFFCCCC' }  // 연한 빨간색
                            };
                        });
                    }
                });
                
                // 열 너비 자동 조정
                headers.forEach((header, idx) => {
                    const column = worksheet.getColumn(idx + 1);
                    let maxLength = header.length;
                    
                    sheetInfo.data.forEach(row => {
                        const value = row[header];
                        if (value) {
                            const len = String(value).length;
                            maxLength = Math.max(maxLength, len);
                        }
                    });
                    
                    column.width = Math.min(maxLength + 2, 50);
                });
            }
        }
        
        // 파일 다운로드
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        saveAs(blob, fileName);
    },

    // 새 워크북 데이터 생성
    createWorkbook() {
        return {
            sheets: []
        };
    },

    // 시트 추가
    addSheet(workbook, data, sheetName) {
        workbook.sheets.push({
            name: sheetName,
            data: data
        });
    }
};

// 상태 표시 유틸리티
const StatusManager = {
    show(elementId, type, message) {
        const status = document.getElementById(elementId);
        if (status) {
            status.className = 'status ' + type;
            status.innerHTML = message;
        }
    },

    processing(elementId, message = '처리 중...') {
        this.show(elementId, 'processing', `<span class="spinner"></span>${message}`);
    },

    success(elementId, message) {
        this.show(elementId, 'success', message);
    },

    error(elementId, message) {
        this.show(elementId, 'error', message);
    },

    hide(elementId) {
        const status = document.getElementById(elementId);
        if (status) {
            status.className = 'status';
            status.innerHTML = '';
        }
    }
};

// 파일 입력 관리 유틸리티
const FileInputManager = {
    setup(inputId, displayId, onChange) {
        const input = document.getElementById(inputId);
        const display = document.getElementById(displayId);

        if (!input) return null;

        let file = null;

        input.addEventListener('change', function(e) {
            file = e.target.files[0];
            if (display) {
                display.textContent = file ? file.name : '';
            }
            input.classList.toggle('has-file', !!file);
            if (onChange) onChange(file);
        });

        return {
            getFile: () => file,
            clear: () => {
                file = null;
                input.value = '';
                if (display) display.textContent = '';
                input.classList.remove('has-file');
            }
        };
    }
};

export { ExcelCore, StatusManager, FileInputManager };
