//! Binary importers for `DatabaseDocument`.
//!
//! Supports `csv` (single table) and `odb`.

use lo_core::{
    parse_xml_document, ColumnDef, ColumnType, DatabaseDocument, DbValue, LoError, Result,
    TableData, XmlNode,
};
use lo_zip::ZipArchive;

pub fn load_bytes(
    title: impl Into<String>,
    bytes: &[u8],
    format: &str,
    table_name: Option<&str>,
) -> Result<DatabaseDocument> {
    let title = title.into();
    match format.to_ascii_lowercase().as_str() {
        "csv" => Ok(from_csv(
            title,
            table_name.unwrap_or("data"),
            &String::from_utf8_lossy(bytes),
        )),
        "odb" => from_odb_bytes(title, bytes),
        other => Err(LoError::Unsupported(format!("base import format {other}"))),
    }
}

pub fn from_csv(title: impl Into<String>, table_name: &str, csv: &str) -> DatabaseDocument {
    let mut db = DatabaseDocument::new(title);
    let mut lines = csv.lines();
    let header = lines.next().unwrap_or("");
    let columns: Vec<ColumnDef> = header
        .split(',')
        .map(|name| ColumnDef {
            name: name.trim().to_string(),
            column_type: ColumnType::Text,
        })
        .collect();
    let mut rows: Vec<Vec<DbValue>> = Vec::new();
    for line in lines {
        let row: Vec<DbValue> = line
            .split(',')
            .map(|cell| parse_db_value(cell.trim()))
            .collect();
        rows.push(row);
    }
    let mut inferred_columns = columns.clone();
    refine_column_types(&mut inferred_columns, &rows);
    db.tables.push(TableData {
        name: table_name.to_string(),
        columns: inferred_columns,
        rows,
    });
    db
}

pub fn from_odb_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<DatabaseDocument> {
    let title = title.into();
    let zip = ZipArchive::new(bytes)?;
    let content = parse_xml_document(&zip.read_string("content.xml")?)?;
    let body = content
        .child("body")
        .ok_or_else(|| LoError::Parse("content.xml missing office:body".to_string()))?;
    let database = body
        .child("database")
        .ok_or_else(|| LoError::Parse("content.xml missing office:database".to_string()))?;

    let mut db = DatabaseDocument::new(title);

    // Two layouts are supported:
    // 1. Native ODB-style with `db:table-representations` referencing CSV
    //    files under `database/<name>.csv` (matches lo_odf::save_database_document).
    // 2. Inline `<table><row><cell/></row></table>` blocks (legacy/manual).
    let mut found_inline = false;
    if let Some(reps) = database
        .child("data-source")
        .and_then(|ds| ds.child("table-representations"))
        .or_else(|| database.child("table-representations"))
    {
        for rep in reps
            .children_named("table-representation")
            .chain(reps.children_named("db:table-representation"))
        {
            let Some(name) = rep.attr("name").or_else(|| rep.attr("db:name")) else {
                continue;
            };
            let csv_path = format!("database/{name}.csv");
            if zip.contains(&csv_path) {
                let csv = zip.read_string(&csv_path)?;
                if let Some(table) = parse_csv_table(name, &csv) {
                    db.tables.push(table);
                }
            } else {
                db.tables.push(TableData {
                    name: name.to_string(),
                    columns: Vec::new(),
                    rows: Vec::new(),
                });
            }
        }
    }

    for table_node in database.children_named("table") {
        found_inline = true;
        db.tables.push(parse_odb_table(table_node));
    }

    // Fallback: if neither produced anything, scan all `database/*.csv` files.
    if db.tables.is_empty() && !found_inline {
        for entry in zip.entries() {
            if let Some(name) = entry
                .strip_prefix("database/")
                .and_then(|name| name.strip_suffix(".csv"))
            {
                let csv = zip.read_string(entry)?;
                if let Some(table) = parse_csv_table(name, &csv) {
                    db.tables.push(table);
                }
            }
        }
    }

    Ok(db)
}

fn parse_csv_table(name: &str, csv: &str) -> Option<TableData> {
    let single = from_csv("table", name, csv);
    single.tables.into_iter().next()
}

fn parse_odb_table(node: &XmlNode) -> TableData {
    let name = node.attr("name").unwrap_or("table").to_string();
    let mut columns: Vec<ColumnDef> = Vec::new();
    let mut rows: Vec<Vec<DbValue>> = Vec::new();
    for row_node in node.children_named("row") {
        let mut row: Vec<DbValue> = Vec::new();
        for cell in row_node.children_named("cell") {
            let column_name = cell.attr("name").unwrap_or("col").to_string();
            if !columns.iter().any(|existing| existing.name == column_name) {
                columns.push(ColumnDef {
                    name: column_name,
                    column_type: ColumnType::Text,
                });
            }
            row.push(parse_db_value(cell.text_content().trim()));
        }
        rows.push(row);
    }
    refine_column_types(&mut columns, &rows);
    TableData {
        name,
        columns,
        rows,
    }
}

fn refine_column_types(columns: &mut [ColumnDef], rows: &[Vec<DbValue>]) {
    for (idx, col) in columns.iter_mut().enumerate() {
        let mut inferred: Option<ColumnType> = None;
        for row in rows {
            let Some(value) = row.get(idx) else { continue };
            if matches!(value, DbValue::Null) {
                continue;
            }
            inferred = Some(match (&inferred, value) {
                (None, DbValue::Integer(_)) => ColumnType::Integer,
                (None, DbValue::Float(_)) => ColumnType::Float,
                (None, DbValue::Bool(_)) => ColumnType::Bool,
                (None, DbValue::Text(_) | DbValue::Null) => ColumnType::Text,
                (Some(prev), value) => merge_column_type(prev, value),
            });
        }
        if let Some(t) = inferred {
            col.column_type = t;
        }
    }
}

fn parse_db_value(text: &str) -> DbValue {
    if text.is_empty() {
        return DbValue::Null;
    }
    if let Ok(value) = text.parse::<i64>() {
        return DbValue::Integer(value);
    }
    if let Ok(value) = text.parse::<f64>() {
        return DbValue::Float(value);
    }
    match text.to_ascii_lowercase().as_str() {
        "true" => DbValue::Bool(true),
        "false" => DbValue::Bool(false),
        _ => DbValue::Text(text.to_string()),
    }
}

fn merge_column_type(current: &ColumnType, value: &DbValue) -> ColumnType {
    match (current, value) {
        (_, DbValue::Null) => current.clone(),
        (ColumnType::Text, _) => ColumnType::Text,
        (ColumnType::Integer, DbValue::Integer(_)) => ColumnType::Integer,
        (ColumnType::Integer, DbValue::Float(_)) => ColumnType::Float,
        (ColumnType::Float, DbValue::Integer(_) | DbValue::Float(_)) => ColumnType::Float,
        (ColumnType::Bool, DbValue::Bool(_)) => ColumnType::Bool,
        _ => ColumnType::Text,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn csv_import_creates_table() {
        let db = from_csv("db", "people", "name,age\nAlice,30\nBob,40");
        let table = db.table("people").expect("table");
        assert_eq!(table.columns.len(), 2);
        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.columns[1].column_type, ColumnType::Integer);
    }

    #[test]
    fn odb_round_trip_imports_tables() {
        let mut db = DatabaseDocument::new("db");
        db.tables.push(TableData {
            name: "people".to_string(),
            columns: vec![
                ColumnDef {
                    name: "name".to_string(),
                    column_type: ColumnType::Text,
                },
                ColumnDef {
                    name: "age".to_string(),
                    column_type: ColumnType::Integer,
                },
            ],
            rows: vec![vec![
                DbValue::Text("Alice".to_string()),
                DbValue::Integer(30),
            ]],
        });
        let tmp = std::env::temp_dir().join("lo_base_import_test.odb");
        lo_odf::save_database_document(&tmp, &db).expect("save odb");
        let bytes = std::fs::read(&tmp).expect("read");
        let _ = std::fs::remove_file(&tmp);
        let loaded = from_odb_bytes("db", &bytes).expect("import odb");
        assert!(loaded.table("people").is_some());
    }
}
