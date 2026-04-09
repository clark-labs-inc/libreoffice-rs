//! HTML export for `DatabaseDocument`. Each table renders as `<h2>` + table.

use lo_core::{escape_text, html_escape, DatabaseDocument, DbValue, TableData};

fn db_value_to_string(value: &DbValue) -> String {
    match value {
        DbValue::Null => String::new(),
        DbValue::Integer(v) => v.to_string(),
        DbValue::Float(v) => v.to_string(),
        DbValue::Bool(v) => v.to_string(),
        DbValue::Text(v) => v.clone(),
    }
}

fn render_table(table: &TableData) -> String {
    let mut out = format!(
        "<h2>{}</h2>\n<table border=\"1\">\n<thead><tr>",
        escape_text(&table.name)
    );
    for col in &table.columns {
        out.push_str(&format!("<th>{}</th>", html_escape(&col.name)));
    }
    out.push_str("</tr></thead>\n<tbody>\n");
    for row in &table.rows {
        out.push_str("<tr>");
        for cell in row {
            out.push_str(&format!(
                "<td>{}</td>",
                html_escape(&db_value_to_string(cell))
            ));
        }
        out.push_str("</tr>\n");
    }
    out.push_str("</tbody>\n</table>\n");
    out
}

pub fn to_html(database: &DatabaseDocument) -> String {
    let mut body = String::new();
    for table in &database.tables {
        body.push_str(&render_table(table));
    }
    format!(
        "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\"/>\n<title>{}</title>\n</head>\n<body>\n{}</body>\n</html>\n",
        escape_text(&database.meta.title),
        body
    )
}
