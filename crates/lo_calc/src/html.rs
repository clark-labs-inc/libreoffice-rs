//! HTML export for `Workbook`. Each sheet becomes an `<h2>` followed by a
//! `<table>` whose cells contain the rendered cell value (formulas resolved
//! using the same evaluator as the rest of `lo_calc`).

use lo_core::{escape_text, html_escape, CellValue, Sheet, Workbook};

use crate::{evaluate_formula, Value};

fn render_cell(sheet: &Sheet, value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number(n) => format!("{n}"),
        CellValue::Text(t) => html_escape(t),
        CellValue::Bool(b) => b.to_string(),
        CellValue::Formula(formula) => match evaluate_formula(formula, sheet) {
            Ok(Value::Number(n)) => format!("{n}"),
            Ok(Value::Text(t)) => html_escape(&t),
            Ok(Value::Bool(b)) => b.to_string(),
            Ok(Value::Blank) => String::new(),
            Ok(Value::Error(e)) | Err(lo_core::LoError::Eval(e)) => {
                format!("#ERR {}", html_escape(&e))
            }
            Err(e) => format!("#ERR {}", html_escape(&e.to_string())),
        },
        CellValue::Error(e) => format!("#ERR {}", html_escape(e)),
    }
}

fn render_sheet(sheet: &Sheet) -> String {
    let (max_row, max_col) = sheet.max_extent();
    let mut out = format!(
        "<h2>{}</h2>\n<table border=\"1\">\n",
        escape_text(&sheet.name)
    );
    for row in 0..=max_row {
        out.push_str("<tr>");
        for col in 0..=max_col {
            let value = sheet
                .get(lo_core::CellAddr::new(row, col))
                .map(|cell| render_cell(sheet, &cell.value))
                .unwrap_or_default();
            let tag = if sheet.has_header && row == 0 {
                "th"
            } else {
                "td"
            };
            out.push_str(&format!("<{tag}>{value}</{tag}>"));
        }
        out.push_str("</tr>\n");
    }
    out.push_str("</table>\n");
    out
}

pub fn to_html(workbook: &Workbook) -> String {
    let mut body = String::new();
    for sheet in &workbook.sheets {
        body.push_str(&render_sheet(sheet));
    }
    format!(
        "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\"/>\n<title>{}</title>\n</head>\n<body>\n{}</body>\n</html>\n",
        escape_text(&workbook.meta.title),
        body
    )
}
