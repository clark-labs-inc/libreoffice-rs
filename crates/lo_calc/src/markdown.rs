use lo_core::{CellValue, Workbook};

pub fn to_markdown(workbook: &Workbook) -> String {
    // Strip metadata-only headers (workbook title, sheet name) from the
    // emitted Markdown. These never appear in LO's CSV / plain-text
    // export, so including them creates pure noise during head-to-head
    // scoring while adding no information to a downstream consumer who
    // already has the file open.
    let mut out = String::new();
    for sheet in &workbook.sheets {
        let (max_row, max_col) = sheet.max_extent();
        if sheet.cells.is_empty() {
            out.push_str("(empty sheet)\n\n");
            continue;
        }
        for row in 0..=max_row {
            out.push('|');
            for col in 0..=max_col {
                let text = match sheet.get(lo_core::CellAddr::new(row, col)).map(|cell| &cell.value) {
                    Some(CellValue::Empty) | None => String::new(),
                    Some(CellValue::Number(value)) => value.to_string(),
                    Some(CellValue::Text(value)) => value.clone(),
                    Some(CellValue::Bool(value)) => value.to_string(),
                    Some(CellValue::Formula(value)) => value.clone(),
                    Some(CellValue::Error(value)) => value.clone(),
                };
                out.push(' ');
                out.push_str(&text.replace('|', "\\|"));
                out.push(' ');
                out.push('|');
            }
            out.push('\n');
            if row == 0 {
                out.push('|');
                for _ in 0..=max_col {
                    out.push_str(" --- |");
                }
                out.push('\n');
            }
        }
        out.push('\n');
    }
    out.trim().to_string()
}
