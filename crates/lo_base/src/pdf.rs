//! PDF export of a `DatabaseDocument`. Each table is rendered as text rows.

use lo_core::{units::Length, write_text_pdf, DatabaseDocument, DbValue};

fn db_value_to_string(value: &DbValue) -> String {
    match value {
        DbValue::Null => String::new(),
        DbValue::Integer(v) => v.to_string(),
        DbValue::Float(v) => v.to_string(),
        DbValue::Bool(v) => v.to_string(),
        DbValue::Text(v) => v.clone(),
    }
}

pub fn to_pdf(database: &DatabaseDocument) -> Vec<u8> {
    let mut lines = Vec::new();
    lines.push(database.meta.title.clone());
    lines.push(String::new());
    for table in &database.tables {
        lines.push(format!("[{}]", table.name));
        let header: Vec<String> = table.columns.iter().map(|c| c.name.clone()).collect();
        lines.push(header.join(" | "));
        for row in &table.rows {
            let cells: Vec<String> = row.iter().map(db_value_to_string).collect();
            lines.push(cells.join(" | "));
        }
        lines.push(String::new());
    }
    write_text_pdf(&lines, Length::pt(842.0), Length::pt(595.0))
}
