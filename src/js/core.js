/**
 * 엑셀 변환 공통 유틸리티
 */

const ExcelCore = {
    // 엑셀 파일 읽기
    readFile(file) {
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
    },

    // 셀 값 가져오기 (1-indexed)
    getCellValue(sheet, row, col) {
        const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
        const cell = sheet[cellAddress];
        return cell ? cell.v : null;
    },

    // 날짜 포맷팅
    formatDate(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    },

    // 엑셀 파일 다운로드
    downloadExcel(workbook, fileName) {
        XLSX.writeFile(workbook, fileName);
    },

    // 새 워크북 생성
    createWorkbook() {
        return XLSX.utils.book_new();
    },

    // 시트 추가
    addSheet(workbook, data, sheetName) {
        const sheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
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
