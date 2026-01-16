use calamine::{open_workbook_from_rs, Reader, Xlsx, Data};
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::io::Cursor;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = console)]
    fn log(s: &str);
}

macro_rules! console_log {
    ($($t:tt)*) => (log(&format!($($t)*)))
}

#[wasm_bindgen(start)]
pub fn init() {
    #[cfg(feature = "console_error_panic_hook")]
    console_error_panic_hook::set_once();
}

// 매핑 정보
#[derive(Debug, Clone)]
struct MappingEntry {
    code: String,
    system_name: String,
}

// 변환 결과 데이터 행
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataRow {
    pub date: String,
    pub code: String,
    pub store_name: String,
    pub product_name: String,
    pub box_qty: i32,
    pub afternoon: String,
}

// 검증 데이터 행
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ValidationRow {
    pub date: String,
    pub day_name: String,
    pub extracted_box: i32,
    pub original_total: i32,
    pub original_store_sum: i32,
    pub result: String,
}

// 매장별 상세 행
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StoreDailyRow {
    pub date: String,
    pub code: String,
    pub store_name: String,
    pub box_sum: i32,
}

// 변환 결과 (JS로 반환)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConversionResult {
    pub data: Vec<DataRow>,
    pub validation: Vec<ValidationRow>,
    pub store_daily: Vec<StoreDailyRow>,
    pub mapping_failures: Vec<String>,
    pub success: bool,
    pub error: Option<String>,
}

// 매장 블록 정보
#[derive(Debug, Clone)]
struct StoreBlock {
    store_name: String,
    row: u32,
    col_no: u32,
    col_afternoon: u32,
    col_product: u32,
    col_box: u32,
}

// 매장명 추출
fn extract_store_name(value: &str) -> Option<String> {
    if !value.contains('※') {
        return None;
    }

    let re = Regex::new(r"※\s*(.+?)\s*:\s*\d*").ok()?;
    re.captures(value)
        .and_then(|caps| caps.get(1))
        .map(|m| m.as_str().trim().to_string())
}

// 파일명에서 날짜 추출
fn extract_date_from_filename(filename: &str) -> (i32, u32, u32) {
    let date_re = Regex::new(r"\((\d+)\.(\d+)~(\d+)\.(\d+)\)").unwrap();
    let year_re = Regex::new(r"(\d+)년\s*(\d+)월").unwrap();

    let mut year = 2026i32;
    let mut month = 1u32;
    let mut day = 1u32;

    if let Some(year_caps) = year_re.captures(filename) {
        year = year_caps.get(1).unwrap().as_str().parse().unwrap_or(26);
        if year < 100 {
            year += 2000;
        }
    }

    if let Some(date_caps) = date_re.captures(filename) {
        month = date_caps.get(1).unwrap().as_str().parse().unwrap_or(1);
        day = date_caps.get(2).unwrap().as_str().parse().unwrap_or(1);
    }

    (year, month, day)
}

// 날짜 포맷
fn format_date(year: i32, month: u32, day: u32) -> String {
    format!("{:04}-{:02}-{:02}", year, month, day)
}

// 날짜 더하기
fn add_days(year: i32, month: u32, day: u32, days: u32) -> (i32, u32, u32) {
    let days_in_month = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let mut new_day = day + days;
    let mut new_month = month;
    let mut new_year = year;

    let is_leap = (new_year % 4 == 0 && new_year % 100 != 0) || (new_year % 400 == 0);
    let max_days = if new_month == 2 && is_leap {
        29
    } else if new_month >= 1 && new_month <= 12 {
        days_in_month[new_month as usize]
    } else {
        31
    };

    while new_day > max_days {
        new_day -= max_days;
        new_month += 1;
        if new_month > 12 {
            new_month = 1;
            new_year += 1;
        }
    }

    (new_year, new_month, new_day)
}

// 셀 값을 문자열로
fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::String(s) => s.clone(),
        Data::Float(f) => format!("{}", f),
        Data::Int(i) => format!("{}", i),
        Data::Bool(b) => format!("{}", b),
        Data::DateTime(dt) => format!("{}", dt),
        Data::DateTimeIso(s) => s.clone(),
        Data::DurationIso(s) => s.clone(),
        Data::Error(_) => String::new(),
        Data::Empty => String::new(),
    }
}

// 셀 값을 정수로
fn cell_to_int(cell: &Data) -> i32 {
    match cell {
        Data::Float(f) => *f as i32,
        Data::Int(i) => *i as i32,
        Data::String(s) => s.parse().unwrap_or(0),
        _ => 0,
    }
}

// 매핑 테이블 파싱
fn parse_mapping_table(data: &[u8]) -> Result<HashMap<String, MappingEntry>, String> {
    let cursor = Cursor::new(data);
    let mut workbook: Xlsx<_> = open_workbook_from_rs(cursor)
        .map_err(|e| format!("매핑 파일 열기 실패: {}", e))?;

    let sheet_names = workbook.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err("매핑 파일에 시트가 없습니다".to_string());
    }

    let range = workbook.worksheet_range(&sheet_names[0])
        .map_err(|e| format!("매핑 시트 읽기 실패: {}", e))?;

    let mut mapping = HashMap::new();
    let mut headers: HashMap<String, usize> = HashMap::new();

    if let Some(first_row) = range.rows().next() {
        for (idx, cell) in first_row.iter().enumerate() {
            let header = cell_to_string(cell);
            headers.insert(header, idx);
        }
    }

    let code_idx = headers.get("코드").copied();
    let orig_name_idx = headers.get("원본 사업장명").copied();
    let sys_name_idx = headers.get("사업장명").copied();

    for row in range.rows().skip(1) {
        let orig_name = orig_name_idx
            .and_then(|i| row.get(i))
            .map(cell_to_string)
            .unwrap_or_default();

        if orig_name.is_empty() {
            continue;
        }

        let code = code_idx
            .and_then(|i| row.get(i))
            .map(cell_to_string)
            .unwrap_or_default();

        let sys_name = sys_name_idx
            .and_then(|i| row.get(i))
            .map(cell_to_string)
            .unwrap_or_default();

        mapping.insert(orig_name, MappingEntry {
            code,
            system_name: sys_name,
        });
    }

    Ok(mapping)
}

// 매장 블록 찾기
fn find_store_blocks(range: &calamine::Range<Data>) -> Vec<StoreBlock> {
    let mut blocks = Vec::new();

    for (row_idx, row) in range.rows().enumerate() {
        if let Some(cell) = row.get(1) {
            let value = cell_to_string(cell);
            if let Some(store_name) = extract_store_name(&value) {
                blocks.push(StoreBlock {
                    store_name,
                    row: row_idx as u32,
                    col_no: 1,
                    col_afternoon: 2,
                    col_product: 4,
                    col_box: 5,
                });
            }
        }

        if let Some(cell) = row.get(10) {
            let value = cell_to_string(cell);
            if let Some(store_name) = extract_store_name(&value) {
                blocks.push(StoreBlock {
                    store_name,
                    row: row_idx as u32,
                    col_no: 10,
                    col_afternoon: 11,
                    col_product: 13,
                    col_box: 14,
                });
            }
        }
    }

    blocks
}

// 블록에서 상품 추출
fn extract_products_from_block(
    range: &calamine::Range<Data>,
    block: &StoreBlock,
    max_products: usize,
) -> Vec<(String, i32, String)> {
    let mut products = Vec::new();
    let start_row = block.row as usize + 4;

    for row_idx in start_row..(start_row + max_products) {
        if let Some(row) = range.rows().nth(row_idx) {
            let no_val = row.get(block.col_no as usize)
                .map(cell_to_string)
                .unwrap_or_default();

            if no_val.parse::<i32>().is_err() {
                break;
            }

            let product_name = row.get(block.col_product as usize)
                .map(cell_to_string)
                .unwrap_or_default();

            if product_name.is_empty() {
                continue;
            }

            let box_qty = row.get(block.col_box as usize)
                .map(cell_to_int)
                .unwrap_or(0);

            if box_qty == 0 {
                continue;
            }

            // 오후 진열 값 추출
            let afternoon = row.get(block.col_afternoon as usize)
                .map(cell_to_string)
                .unwrap_or_default()
                .trim()
                .to_string();

            products.push((product_name.trim().to_string(), box_qty, afternoon));
        }
    }

    products
}

// 요일별 총계 추출
fn get_day_totals(range: &calamine::Range<Data>) -> (i32, i32) {
    let total_box = range.rows().nth(7)
        .and_then(|row| row.get(5))
        .map(cell_to_int)
        .unwrap_or(0);

    let mut store_box_sum = 0;
    for (row_idx, row) in range.rows().enumerate() {
        if row_idx < 34 {
            continue;
        }

        if let Some(cell) = row.get(1) {
            if cell_to_string(cell) == "계" {
                if let Some(box_cell) = row.get(5) {
                    let val = cell_to_int(box_cell);
                    if val > 0 {
                        store_box_sum += val;
                    }
                }
            }
        }

        if let Some(cell) = row.get(10) {
            if cell_to_string(cell) == "계" {
                if let Some(box_cell) = row.get(14) {
                    let val = cell_to_int(box_cell);
                    if val > 0 {
                        store_box_sum += val;
                    }
                }
            }
        }
    }

    (total_box, store_box_sum)
}

// 메인 변환 함수 - JSON 결과 반환 (Excel 생성은 JS에서)
#[wasm_bindgen]
pub fn convert_excel(
    origin_data: &[u8],
    mapping_data: &[u8],
    filename: &str,
) -> JsValue {
    let result = convert_internal(origin_data, mapping_data, filename);
    serde_wasm_bindgen::to_value(&result).unwrap_or(JsValue::NULL)
}

fn convert_internal(
    origin_data: &[u8],
    mapping_data: &[u8],
    filename: &str,
) -> ConversionResult {
    console_log!("WASM: Starting conversion...");

    // 매핑 테이블 파싱
    let mapping = match parse_mapping_table(mapping_data) {
        Ok(m) => m,
        Err(e) => {
            return ConversionResult {
                data: vec![],
                validation: vec![],
                store_daily: vec![],
                mapping_failures: vec![],
                success: false,
                error: Some(e),
            };
        }
    };
    console_log!("WASM: Mapping loaded - {} entries", mapping.len());

    // 원본 파일 열기
    let cursor = Cursor::new(origin_data);
    let mut workbook: Xlsx<_> = match open_workbook_from_rs(cursor) {
        Ok(wb) => wb,
        Err(e) => {
            return ConversionResult {
                data: vec![],
                validation: vec![],
                store_daily: vec![],
                mapping_failures: vec![],
                success: false,
                error: Some(format!("원본 파일 열기 실패: {}", e)),
            };
        }
    };

    let sheet_names = workbook.sheet_names().to_vec();
    console_log!("WASM: Origin loaded - {} sheets", sheet_names.len());

    let day_names = ["월", "화", "수", "목", "금"];
    let (base_year, base_month, base_day) = extract_date_from_filename(filename);

    let mut all_data: Vec<DataRow> = Vec::new();
    let mut mapping_failures: Vec<String> = Vec::new();
    let mut validation: Vec<ValidationRow> = Vec::new();

    for (day_idx, day_name) in day_names.iter().enumerate() {
        if !sheet_names.contains(&day_name.to_string()) {
            continue;
        }

        let range = match workbook.worksheet_range(day_name) {
            Ok(r) => r,
            Err(_) => continue,
        };

        let (year, month, day) = add_days(base_year, base_month, base_day, day_idx as u32);
        let date_str = format_date(year, month, day);

        let blocks = find_store_blocks(&range);

        for block in &blocks {
            let (code, system_name) = if let Some(entry) = mapping.get(&block.store_name) {
                (entry.code.clone(), entry.system_name.clone())
            } else {
                if !mapping_failures.contains(&block.store_name) {
                    mapping_failures.push(block.store_name.clone());
                }
                ("MAPPING_FAILED".to_string(), format!("[매핑실패] {}", block.store_name))
            };

            let products = extract_products_from_block(&range, block, 25);

            for (product_name, box_qty, afternoon) in products {
                all_data.push(DataRow {
                    date: date_str.clone(),
                    code: code.clone(),
                    store_name: system_name.clone(),
                    product_name,
                    box_qty,
                    afternoon,
                });
            }
        }

        // 검증 데이터
        let (original_total, original_store_sum) = get_day_totals(&range);
        let extracted_box: i32 = all_data.iter()
            .filter(|r| r.date == date_str)
            .map(|r| r.box_qty)
            .sum();

        let result = if original_store_sum > 0 {
            if extracted_box == original_store_sum {
                "일치".to_string()
            } else {
                format!("불일치 (차이: {})", extracted_box - original_store_sum)
            }
        } else {
            "원본 데이터 없음".to_string()
        };

        validation.push(ValidationRow {
            date: date_str,
            day_name: day_name.to_string(),
            extracted_box,
            original_total,
            original_store_sum,
            result,
        });
    }

    console_log!("WASM: Extracted {} rows", all_data.len());

    // 매장별 상세
    let mut store_daily_map: HashMap<String, StoreDailyRow> = HashMap::new();
    for row in &all_data {
        let key = format!("{}_{}_{}", row.date, row.code, row.store_name);
        store_daily_map
            .entry(key)
            .and_modify(|e| e.box_sum += row.box_qty)
            .or_insert(StoreDailyRow {
                date: row.date.clone(),
                code: row.code.clone(),
                store_name: row.store_name.clone(),
                box_sum: row.box_qty,
            });
    }

    console_log!("WASM: Conversion complete!");

    ConversionResult {
        data: all_data,
        validation,
        store_daily: store_daily_map.into_values().collect(),
        mapping_failures,
        success: true,
        error: None,
    }
}
