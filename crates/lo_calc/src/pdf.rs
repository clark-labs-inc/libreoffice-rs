//! PDF export of a `Workbook`. Each sheet is rendered as a tab-separated text
//! page using `lo_core::pdf::write_text_pdf`.

use lo_core::{units::Length, write_text_pdf, CellAddr, CellValue, Sheet, Workbook};

use crate::{evaluate_formula, Value};

fn cell_text(sheet: &Sheet, value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number(n) => format!("{n}"),
        CellValue::Text(t) => t.clone(),
        CellValue::Bool(b) => b.to_string(),
        CellValue::Formula(formula) => match evaluate_formula(formula, sheet) {
            Ok(Value::Number(n)) => format!("{n}"),
            Ok(Value::Text(t)) => t,
            Ok(Value::Bool(b)) => b.to_string(),
            Ok(Value::Blank) => String::new(),
            _ => "#ERR".to_string(),
        },
        CellValue::Error(_) => "#ERR".to_string(),
    }
}

fn sheet_lines(sheet: &Sheet) -> Vec<String> {
    let (max_row, max_col) = sheet.max_extent();
    let mut lines = Vec::new();
    lines.push(format!("[{}]", sheet.name));
    for row in 0..=max_row {
        let mut cells = Vec::with_capacity(max_col as usize + 1);
        for col in 0..=max_col {
            let text = sheet
                .get(CellAddr::new(row, col))
                .map(|cell| cell_text(sheet, &cell.value))
                .unwrap_or_default();
            cells.push(text);
        }
        lines.push(cells.join("\t"));
    }
    lines
}

pub fn to_pdf(workbook: &Workbook) -> Vec<u8> {
    let mut lines = Vec::new();
    for sheet in &workbook.sheets {
        lines.extend(sheet_lines(sheet));
        lines.push(String::new());
    }
    write_text_pdf(&lines, Length::pt(842.0), Length::pt(595.0))
}
