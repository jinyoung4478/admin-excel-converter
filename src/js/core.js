/**
 * 엑셀 변환 공통 유틸리티 (ExcelJS 기반)
 */

// ========== 상수 ==========
const EXCEL_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const DEFAULT_CELL_REF = 'A1:A1';
const MAX_COLUMN_WIDTH = 50;
const COLUMN_WIDTH_PADDING = 2;

const HEADER_STYLE = {
    font: { bold: true },
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
    }
};

const MAPPING_FAILED_STYLE = {
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCCCC' }
    }
};

// ========== 셀 주소 유틸리티 ==========
const CellAddress = {
    /**
     * 0-indexed row, col을 "A1" 형식으로 변환
     */
    encode(row, col) {
        const colStr = this._columnToLetter(col);
        return colStr + (row + 1);
    },

    /**
     * "A1" 형식을 {r: 0, c: 0} 객체로 변환
     */
    decode(address) {
        let col = 0;
        let i = 0;

        while (i < address.length && /[A-Z]/i.test(address[i])) {
            col = col * 26 + (address.charCodeAt(i) & 0x1F);
            i++;
        }

        return {
            r: parseInt(address.substring(i)) - 1,
            c: col - 1
        };
    },

    /**
     * 범위 문자열 ("A1:B2")을 객체로 변환
     */
    decodeRange(ref) {
        const parts = ref.split(':');
        const start = this.decode(parts[0]);
        const end = parts[1] ? this.decode(parts[1]) : start;

        return {
            s: { r: start.r, c: start.c },
            e: { r: end.r, c: end.c }
        };
    },

    _columnToLetter(col) {
        let result = '';
        let c = col;
        do {
            result = String.fromCharCode(65 + (c % 26)) + result;
            c = Math.floor(c / 26) - 1;
        } while (c >= 0);
        return result;
    }
};

// ========== 엑셀 파일 읽기 ==========
const ExcelReader = {
    /**
     * 엑셀 파일을 SheetJS 호환 형식으로 읽기
     */
    async readFile(file) {
        const arrayBuffer = await this._readAsArrayBuffer(file);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        return this._convertToSheetJSFormat(workbook);
    },

    _readAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    },

    _convertToSheetJSFormat(workbook) {
        const result = {
            SheetNames: [],
            Sheets: {}
        };

        workbook.eachSheet((worksheet) => {
            const sheetName = worksheet.name;
            result.SheetNames.push(sheetName);
            result.Sheets[sheetName] = this._convertWorksheet(worksheet);
        });

        return result;
    },

    _convertWorksheet(worksheet) {
        const sheetData = {};
        let maxRow = 0, maxCol = 0;
        let minRow = Infinity, minCol = Infinity;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            maxRow = Math.max(maxRow, rowNumber);
            minRow = Math.min(minRow, rowNumber);

            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                maxCol = Math.max(maxCol, colNumber);
                minCol = Math.min(minCol, colNumber);

                const address = CellAddress.encode(rowNumber - 1, colNumber - 1);
                sheetData[address] = this._extractCellValue(cell);
            });
        });

        sheetData['!ref'] = this._buildRangeRef(minRow, minCol, maxRow, maxCol);
        return sheetData;
    },

    _extractCellValue(cell) {
        let value = cell.value;

        if (value && typeof value === 'object') {
            if (value.result !== undefined) {
                value = value.result;
            } else if (value.text !== undefined) {
                value = value.text;
            }
        }

        return {
            v: value,
            t: typeof value === 'number' ? 'n' : 's'
        };
    },

    _buildRangeRef(minRow, minCol, maxRow, maxCol) {
        if (minRow === Infinity || minCol === Infinity) {
            return DEFAULT_CELL_REF;
        }
        const start = CellAddress.encode(minRow - 1, minCol - 1);
        const end = CellAddress.encode(maxRow - 1, maxCol - 1);
        return `${start}:${end}`;
    }
};

// ========== 시트 데이터 접근 ==========
const SheetAccess = {
    /**
     * 셀 값 가져오기 (1-indexed)
     */
    getCellValue(sheet, row, col) {
        const address = CellAddress.encode(row - 1, col - 1);
        const cell = sheet[address];
        return cell ? cell.v : null;
    },

    /**
     * 시트를 JSON 배열로 변환
     */
    toJson(sheet) {
        const ref = sheet['!ref'];
        if (!ref) return [];

        const range = CellAddress.decodeRange(ref);
        const headers = this._extractHeaders(sheet, range);
        return this._extractDataRows(sheet, range, headers);
    },

    _extractHeaders(sheet, range) {
        const headers = [];
        for (let col = range.s.c; col <= range.e.c; col++) {
            const address = CellAddress.encode(range.s.r, col);
            const cell = sheet[address];
            headers.push(cell ? String(cell.v) : `Column${col}`);
        }
        return headers;
    },

    _extractDataRows(sheet, range, headers) {
        const data = [];

        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const rowData = this._extractRow(sheet, row, range, headers);
            if (rowData) {
                data.push(rowData);
            }
        }

        return data;
    },

    _extractRow(sheet, row, range, headers) {
        const rowData = {};
        let hasData = false;

        for (let col = range.s.c; col <= range.e.c; col++) {
            const address = CellAddress.encode(row, col);
            const cell = sheet[address];
            const header = headers[col - range.s.c];

            if (cell?.v !== undefined && cell.v !== null) {
                rowData[header] = cell.v;
                hasData = true;
            }
        }

        return hasData ? rowData : null;
    }
};

// ========== 엑셀 파일 생성/다운로드 ==========
const ExcelWriter = {
    /**
     * 워크북 데이터 생성
     */
    createWorkbook() {
        return { sheets: [] };
    },

    /**
     * 시트 추가
     */
    addSheet(workbook, data, sheetName) {
        workbook.sheets.push({ name: sheetName, data });
    },

    /**
     * 엑셀 파일 다운로드
     */
    async downloadExcel(workbookData, fileName) {
        const workbook = new ExcelJS.Workbook();

        for (const sheetInfo of workbookData.sheets) {
            this._addWorksheet(workbook, sheetInfo);
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: EXCEL_MIME_TYPE });
        saveAs(blob, fileName);
    },

    _addWorksheet(workbook, sheetInfo) {
        const worksheet = workbook.addWorksheet(sheetInfo.name);

        if (!sheetInfo.data?.length) return;

        const headers = this._getVisibleHeaders(sheetInfo.data[0]);
        this._addHeaderRow(worksheet, headers);
        this._addDataRows(worksheet, sheetInfo, headers);
        this._autoFitColumns(worksheet, headers, sheetInfo.data);
    },

    _getVisibleHeaders(firstRow) {
        return Object.keys(firstRow).filter(h => !h.startsWith('_'));
    },

    _addHeaderRow(worksheet, headers) {
        worksheet.addRow(headers);

        const headerRow = worksheet.getRow(1);
        headerRow.font = HEADER_STYLE.font;
        headerRow.fill = HEADER_STYLE.fill;
    },

    _addDataRows(worksheet, sheetInfo, headers) {
        const isDataSheet = sheetInfo.name === '데이터';

        sheetInfo.data.forEach((row) => {
            const values = headers.map(h => row[h]);
            const excelRow = worksheet.addRow(values);

            if (isDataSheet && row['_isMappingFailed']) {
                this._applyMappingFailedStyle(excelRow);
            }
        });
    },

    _applyMappingFailedStyle(row) {
        row.eachCell((cell) => {
            cell.fill = MAPPING_FAILED_STYLE.fill;
        });
    },

    _autoFitColumns(worksheet, headers, data) {
        headers.forEach((header, idx) => {
            const column = worksheet.getColumn(idx + 1);
            const maxLength = this._calculateMaxLength(header, data);
            column.width = Math.min(maxLength + COLUMN_WIDTH_PADDING, MAX_COLUMN_WIDTH);
        });
    },

    _calculateMaxLength(header, data) {
        let maxLength = header.length;

        data.forEach(row => {
            const value = row[header];
            if (value) {
                maxLength = Math.max(maxLength, String(value).length);
            }
        });

        return maxLength;
    }
};

// ========== 날짜 유틸리티 ==========
const DateUtils = {
    /**
     * Date 객체를 "YYYY-MM-DD" 형식으로 포맷
     */
    format(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }
};

// ========== ExcelCore (하위 호환용 통합 인터페이스) ==========
const ExcelCore = {
    // 파일 읽기
    readFile: (file) => ExcelReader.readFile(file),

    // 셀 유틸리티
    _encodeCell: (row, col) => CellAddress.encode(row, col),
    _decodeCell: (addr) => CellAddress.decode(addr),
    decodeRange: (ref) => CellAddress.decodeRange(ref),
    getCellValue: (sheet, row, col) => SheetAccess.getCellValue(sheet, row, col),
    sheetToJson: (sheet) => SheetAccess.toJson(sheet),

    // 파일 생성
    createWorkbook: () => ExcelWriter.createWorkbook(),
    addSheet: (workbook, data, name) => ExcelWriter.addSheet(workbook, data, name),
    downloadExcel: (data, fileName) => ExcelWriter.downloadExcel(data, fileName),

    // 날짜
    formatDate: (date) => DateUtils.format(date)
};

// ========== UI 유틸리티 ==========
const StatusManager = {
    show(elementId, type, message) {
        const status = document.getElementById(elementId);
        if (!status) return;

        status.className = `status ${type}`;
        status.innerHTML = message;
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
        if (!status) return;

        status.className = 'status';
        status.innerHTML = '';
    }
};

const FileInputManager = {
    setup(inputId, displayId, onChange) {
        const input = document.getElementById(inputId);
        if (!input) return null;

        const display = document.getElementById(displayId);
        let currentFile = null;

        input.addEventListener('change', (e) => {
            currentFile = e.target.files[0] || null;

            if (display) {
                display.textContent = currentFile?.name || '';
            }
            input.classList.toggle('has-file', !!currentFile);

            onChange?.(currentFile);
        });

        return {
            getFile: () => currentFile,
            clear: () => {
                currentFile = null;
                input.value = '';
                if (display) display.textContent = '';
                input.classList.remove('has-file');
            }
        };
    }
};

export { ExcelCore, StatusManager, FileInputManager, CellAddress, DateUtils };
