//! SVG render of a `Workbook` — a single-page grid preview of the first sheet.

use lo_core::{
    svg_footer, svg_header, svg_line, svg_text, units::Length, CellAddr, CellValue, Sheet, Size,
    Workbook,
};

use crate::{evaluate_formula, Value};

const COL_WIDTH_PT: f32 = 80.0;
const ROW_HEIGHT_PT: f32 = 18.0;
const HEADER_HEIGHT_PT: f32 = 18.0;
const FONT_SIZE_PT: u32 = 11;

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

pub fn render_svg(workbook: &Workbook, size: Size) -> String {
    let mut svg = String::new();
    svg.push_str(&svg_header(size.width, size.height));

    let sheet = match workbook.sheets.first() {
        Some(s) => s,
        None => {
            svg.push_str(svg_footer());
            return svg;
        }
    };

    let (max_row, max_col) = sheet.max_extent();
    let cols = (max_col + 1).max(1) as f32;
    let rows = (max_row + 1).max(1) as f32;
    let total_w = cols * COL_WIDTH_PT;
    let total_h = HEADER_HEIGHT_PT + rows * ROW_HEIGHT_PT;

    // Header band
    svg.push_str(&svg_text(
        Length::pt(8.0),
        Length::pt(14.0),
        &sheet.name,
        FONT_SIZE_PT + 2,
        "#222222",
        "bold",
    ));

    // Grid lines
    for c in 0..=(cols as u32) {
        let x = c as f32 * COL_WIDTH_PT;
        svg.push_str(&svg_line(
            Length::pt(x),
            Length::pt(HEADER_HEIGHT_PT),
            Length::pt(x),
            Length::pt(HEADER_HEIGHT_PT + total_h),
            "#cccccc",
        ));
    }
    for r in 0..=(rows as u32) {
        let y = HEADER_HEIGHT_PT + r as f32 * ROW_HEIGHT_PT;
        svg.push_str(&svg_line(
            Length::pt(0.0),
            Length::pt(y),
            Length::pt(total_w),
            Length::pt(y),
            "#cccccc",
        ));
    }

    // Cell text
    for row in 0..=max_row {
        for col in 0..=max_col {
            let text = sheet
                .get(CellAddr::new(row, col))
                .map(|cell| cell_text(sheet, &cell.value))
                .unwrap_or_default();
            if text.is_empty() {
                continue;
            }
            let x = col as f32 * COL_WIDTH_PT + 4.0;
            let y = HEADER_HEIGHT_PT + (row as f32 + 0.7) * ROW_HEIGHT_PT;
            let weight = if sheet.has_header && row == 0 {
                "bold"
            } else {
                "normal"
            };
            svg.push_str(&svg_text(
                Length::pt(x),
                Length::pt(y),
                &text,
                FONT_SIZE_PT,
                "#000000",
                weight,
            ));
        }
    }

    svg.push_str(svg_footer());
    svg
}
