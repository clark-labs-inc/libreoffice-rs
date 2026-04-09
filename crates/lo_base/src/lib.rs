pub mod html;
pub mod import;
pub mod pdf;
pub mod svg;

pub use html::to_html;
pub use import::{from_csv, from_odb_bytes, load_bytes};
pub use pdf::to_pdf;
pub use svg::render_svg;

use std::path::Path;

use lo_core::{
    ColumnDef, ColumnType, DatabaseDocument, DbValue, LoError, QueryDef, Result, TableData,
};

#[derive(Clone, Debug, PartialEq)]
pub struct QueryResult {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<DbValue>>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum PredicateOp {
    Eq,
    Ne,
    Lt,
    Lte,
    Gt,
    Gte,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Predicate {
    pub column: String,
    pub op: PredicateOp,
    pub value: DbValue,
}

#[derive(Clone, Debug, PartialEq)]
pub struct OrderBy {
    pub column: String,
    pub descending: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub struct SelectQuery {
    pub columns: Vec<String>,
    pub table: String,
    pub predicate: Option<Predicate>,
    pub order_by: Option<OrderBy>,
    pub limit: Option<usize>,
}

pub fn save_odb(path: impl AsRef<Path>, database: &DatabaseDocument) -> Result<()> {
    lo_odf::save_database_document(path, database)
}

/// Render the database into bytes for the requested format.
///
/// Supported (case-insensitive): `html`, `svg`, `pdf`, `odb`.
pub fn save_as(database: &DatabaseDocument, format: &str) -> Result<Vec<u8>> {
    match format.to_ascii_lowercase().as_str() {
        "html" => Ok(to_html(database).into_bytes()),
        "svg" => Ok(render_svg(database).into_bytes()),
        "pdf" => Ok(to_pdf(database)),
        "odb" => {
            let tmp = std::env::temp_dir().join(format!(
                "lo_base_{}.odb",
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_nanos())
                    .unwrap_or(0)
            ));
            lo_odf::save_database_document(&tmp, database)?;
            let bytes = std::fs::read(&tmp)?;
            let _ = std::fs::remove_file(&tmp);
            Ok(bytes)
        }
        other => Err(LoError::Unsupported(format!(
            "base format not supported: {other}"
        ))),
    }
}

pub fn database_from_csv(
    title: impl Into<String>,
    table_name: &str,
    csv: &str,
) -> Result<DatabaseDocument> {
    let mut database = DatabaseDocument::new(title);
    database.tables.push(csv_to_table(table_name, csv)?);
    Ok(database)
}

pub fn csv_to_table(table_name: &str, csv: &str) -> Result<TableData> {
    let rows = lo_calc::parse_csv(csv)?;
    if rows.is_empty() {
        return Err(LoError::InvalidInput(
            "CSV must include a header row".to_string(),
        ));
    }
    let header = &rows[0];
    let body = &rows[1..];

    let mut columns = Vec::new();
    for (index, name) in header.iter().enumerate() {
        columns.push(ColumnDef {
            name: if name.trim().is_empty() {
                format!("column_{}", index + 1)
            } else {
                name.trim().to_string()
            },
            column_type: infer_column_type(
                body.iter()
                    .filter_map(|row| row.get(index).map(String::as_str)),
            ),
        });
    }

    let mut rows_out = Vec::new();
    for row in body {
        let mut values = Vec::new();
        for (index, column) in columns.iter().enumerate() {
            let cell = row.get(index).map(String::as_str).unwrap_or("");
            values.push(parse_typed_value(cell, &column.column_type));
        }
        rows_out.push(values);
    }

    Ok(TableData {
        name: table_name.to_string(),
        columns,
        rows: rows_out,
    })
}

pub fn parse_select(sql: &str) -> Result<SelectQuery> {
    let trimmed = sql.trim().trim_end_matches(';').trim();
    let upper = trimmed.to_ascii_uppercase();
    if !upper.starts_with("SELECT ") {
        return Err(LoError::Parse("query must start with SELECT".to_string()));
    }
    let from_index = upper
        .find(" FROM ")
        .ok_or_else(|| LoError::Parse("query must contain FROM".to_string()))?;
    let columns_part = trimmed[7..from_index].trim();

    // Slice off everything after FROM and split it into table / WHERE /
    // ORDER BY / LIMIT clauses by walking the recognized keywords in order.
    let rest = &trimmed[from_index + 6..];
    let upper_rest = rest.to_ascii_uppercase();

    let where_pos = upper_rest.find(" WHERE ");
    let order_pos = upper_rest.find(" ORDER BY ");
    let limit_pos = upper_rest.find(" LIMIT ");

    let table_end = [where_pos, order_pos, limit_pos]
        .into_iter()
        .flatten()
        .min()
        .unwrap_or(rest.len());
    let table_part = rest[..table_end].trim();

    let where_part = where_pos.map(|start| {
        let begin = start + " WHERE ".len();
        let end = [order_pos, limit_pos]
            .into_iter()
            .flatten()
            .filter(|p| *p > start)
            .min()
            .unwrap_or(rest.len());
        rest[begin..end].trim()
    });

    let order_by = order_pos.map(|start| {
        let begin = start + " ORDER BY ".len();
        let end = limit_pos.filter(|p| *p > start).unwrap_or(rest.len());
        let clause = rest[begin..end].trim();
        let mut parts = clause.split_whitespace();
        let column = parts.next().unwrap_or("").to_string();
        let descending = matches!(parts.next(), Some(word) if word.eq_ignore_ascii_case("DESC"));
        OrderBy { column, descending }
    });

    let limit = limit_pos.and_then(|start| {
        let begin = start + " LIMIT ".len();
        rest[begin..].trim().parse::<usize>().ok()
    });

    let columns = if columns_part == "*" {
        vec!["*".to_string()]
    } else {
        columns_part
            .split(',')
            .map(|column| column.trim().to_string())
            .collect()
    };
    let predicate = where_part.map(parse_predicate).transpose()?;
    Ok(SelectQuery {
        columns,
        table: table_part.to_string(),
        predicate,
        order_by,
        limit,
    })
}

pub fn execute_select(database: &DatabaseDocument, sql: &str) -> Result<QueryResult> {
    let query = parse_select(sql)?;
    let table = database
        .table(&query.table)
        .ok_or_else(|| LoError::InvalidInput(format!("table not found: {}", query.table)))?;

    let selected_indices: Vec<usize> = if query.columns.len() == 1 && query.columns[0] == "*" {
        (0..table.columns.len()).collect()
    } else {
        query
            .columns
            .iter()
            .map(|column| column_index(table, column))
            .collect::<Result<Vec<_>>>()?
    };

    let mut rows: Vec<Vec<DbValue>> = Vec::new();
    for row in &table.rows {
        if predicate_matches(table, row, query.predicate.as_ref())? {
            rows.push(
                selected_indices
                    .iter()
                    .map(|&index| row[index].clone())
                    .collect(),
            );
        }
    }

    let columns: Vec<String> = selected_indices
        .iter()
        .map(|&index| table.columns[index].name.clone())
        .collect();

    if let Some(order) = &query.order_by {
        // ORDER BY may name either a selected column (preferred) or any
        // table column. We resolve it against the projected columns first
        // and fall back to the underlying table columns.
        let projected_idx = columns
            .iter()
            .position(|name| name.eq_ignore_ascii_case(&order.column));
        if let Some(idx) = projected_idx {
            sort_rows(&mut rows, idx, order.descending);
        } else if let Ok(orig_idx) = column_index(table, &order.column) {
            // Re-project rows so the ORDER BY column is comparable.
            let mut indexed: Vec<(DbValue, Vec<DbValue>)> = rows
                .into_iter()
                .zip(table.rows.iter())
                .filter_map(|(projected, original)| {
                    if predicate_matches(table, original, query.predicate.as_ref()).ok()? {
                        Some((original[orig_idx].clone(), projected))
                    } else {
                        None
                    }
                })
                .collect();
            indexed.sort_by(|a, b| compare_db_for_sort(&a.0, &b.0));
            if order.descending {
                indexed.reverse();
            }
            rows = indexed.into_iter().map(|(_, r)| r).collect();
        }
    }

    if let Some(limit) = query.limit {
        rows.truncate(limit);
    }

    Ok(QueryResult { columns, rows })
}

fn sort_rows(rows: &mut [Vec<DbValue>], index: usize, descending: bool) {
    rows.sort_by(|a, b| {
        compare_db_for_sort(
            a.get(index).unwrap_or(&DbValue::Null),
            b.get(index).unwrap_or(&DbValue::Null),
        )
    });
    if descending {
        rows.reverse();
    }
}

fn compare_db_for_sort(a: &DbValue, b: &DbValue) -> std::cmp::Ordering {
    match (db_to_f64(a), db_to_f64(b)) {
        (Ok(left), Ok(right)) => left
            .partial_cmp(&right)
            .unwrap_or(std::cmp::Ordering::Equal),
        _ => display_value(a).cmp(&display_value(b)),
    }
}

fn display_value(value: &DbValue) -> String {
    match value {
        DbValue::Null => String::new(),
        DbValue::Integer(v) => v.to_string(),
        DbValue::Float(v) => v.to_string(),
        DbValue::Bool(v) => v.to_string(),
        DbValue::Text(v) => v.clone(),
    }
}

pub fn add_query(database: &mut DatabaseDocument, name: &str, sql: &str) {
    database.queries.push(QueryDef {
        name: name.to_string(),
        sql: sql.to_string(),
    });
}

fn infer_column_type<'a>(values: impl Iterator<Item = &'a str>) -> ColumnType {
    let mut seen = false;
    let mut all_int = true;
    let mut all_float = true;
    let mut all_bool = true;

    for value in values {
        let trimmed = value.trim();
        if trimmed.is_empty() {
            continue;
        }
        seen = true;
        if trimmed.parse::<i64>().is_err() {
            all_int = false;
        }
        if trimmed.parse::<f64>().is_err() {
            all_float = false;
        }
        if !trimmed.eq_ignore_ascii_case("true") && !trimmed.eq_ignore_ascii_case("false") {
            all_bool = false;
        }
    }

    if !seen {
        ColumnType::Text
    } else if all_int {
        ColumnType::Integer
    } else if all_float {
        ColumnType::Float
    } else if all_bool {
        ColumnType::Bool
    } else {
        ColumnType::Text
    }
}

fn parse_typed_value(value: &str, column_type: &ColumnType) -> DbValue {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return DbValue::Null;
    }
    match column_type {
        ColumnType::Integer => trimmed
            .parse::<i64>()
            .map(DbValue::Integer)
            .unwrap_or_else(|_| DbValue::Text(trimmed.to_string())),
        ColumnType::Float => trimmed
            .parse::<f64>()
            .map(DbValue::Float)
            .unwrap_or_else(|_| DbValue::Text(trimmed.to_string())),
        ColumnType::Bool => {
            if trimmed.eq_ignore_ascii_case("true") {
                DbValue::Bool(true)
            } else if trimmed.eq_ignore_ascii_case("false") {
                DbValue::Bool(false)
            } else {
                DbValue::Text(trimmed.to_string())
            }
        }
        ColumnType::Text => DbValue::Text(trimmed.to_string()),
    }
}

fn parse_predicate(input: &str) -> Result<Predicate> {
    for (symbol, op) in [
        ("<=", PredicateOp::Lte),
        (">=", PredicateOp::Gte),
        ("<>", PredicateOp::Ne),
        ("=", PredicateOp::Eq),
        ("<", PredicateOp::Lt),
        (">", PredicateOp::Gt),
    ] {
        if let Some(index) = input.find(symbol) {
            let column = input[..index].trim().to_string();
            let value_text = input[index + symbol.len()..].trim();
            let value = parse_literal(value_text)?;
            return Ok(Predicate { column, op, value });
        }
    }
    Err(LoError::Parse("unsupported WHERE predicate".to_string()))
}

fn parse_literal(input: &str) -> Result<DbValue> {
    if input.starts_with('"') && input.ends_with('"') && input.len() >= 2 {
        return Ok(DbValue::Text(
            input[1..input.len() - 1].replace("\"\"", "\""),
        ));
    }
    if input.starts_with('\'') && input.ends_with('\'') && input.len() >= 2 {
        return Ok(DbValue::Text(input[1..input.len() - 1].replace("''", "'")));
    }
    if input.eq_ignore_ascii_case("true") {
        return Ok(DbValue::Bool(true));
    }
    if input.eq_ignore_ascii_case("false") {
        return Ok(DbValue::Bool(false));
    }
    if let Ok(value) = input.parse::<i64>() {
        return Ok(DbValue::Integer(value));
    }
    if let Ok(value) = input.parse::<f64>() {
        return Ok(DbValue::Float(value));
    }
    Ok(DbValue::Text(input.to_string()))
}

fn column_index(table: &TableData, name: &str) -> Result<usize> {
    table
        .columns
        .iter()
        .position(|column| column.name.eq_ignore_ascii_case(name))
        .ok_or_else(|| LoError::InvalidInput(format!("column not found: {name}")))
}

fn predicate_matches(
    table: &TableData,
    row: &[DbValue],
    predicate: Option<&Predicate>,
) -> Result<bool> {
    let Some(predicate) = predicate else {
        return Ok(true);
    };
    let index = column_index(table, &predicate.column)?;
    compare_db_values(&row[index], &predicate.value, &predicate.op)
}

fn compare_db_values(lhs: &DbValue, rhs: &DbValue, op: &PredicateOp) -> Result<bool> {
    let comparison = match (lhs, rhs) {
        (DbValue::Text(left), DbValue::Text(right)) => match left.cmp(right) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        },
        _ => {
            let left = db_to_f64(lhs)?;
            let right = db_to_f64(rhs)?;
            if left < right {
                -1
            } else if left > right {
                1
            } else {
                0
            }
        }
    };
    Ok(match op {
        PredicateOp::Eq => comparison == 0,
        PredicateOp::Ne => comparison != 0,
        PredicateOp::Lt => comparison < 0,
        PredicateOp::Lte => comparison <= 0,
        PredicateOp::Gt => comparison > 0,
        PredicateOp::Gte => comparison >= 0,
    })
}

fn db_to_f64(value: &DbValue) -> Result<f64> {
    match value {
        DbValue::Null => Ok(0.0),
        DbValue::Integer(value) => Ok(*value as f64),
        DbValue::Float(value) => Ok(*value),
        DbValue::Bool(value) => Ok(if *value { 1.0 } else { 0.0 }),
        DbValue::Text(value) => value
            .parse::<f64>()
            .map_err(|_| LoError::Eval(format!("cannot convert text to number: {value}"))),
    }
}

#[cfg(test)]
mod tests {
    use super::{csv_to_table, database_from_csv, execute_select};

    #[test]
    fn infers_schema_from_csv() {
        let table = csv_to_table("People", "name,age,active\nAlice,30,true\nBob,22,false")
            .expect("csv to table");
        assert_eq!(table.columns.len(), 3);
        assert_eq!(table.rows.len(), 2);
    }

    #[test]
    fn executes_simple_select() {
        let database = database_from_csv("Test", "People", "name,age\nAlice,30\nBob,22\nCara,41")
            .expect("database from csv");
        let result = execute_select(&database, "SELECT name FROM People WHERE age > 25")
            .expect("execute select");
        assert_eq!(result.rows.len(), 2);
    }

    #[test]
    fn select_supports_order_by_and_limit() {
        let database = database_from_csv(
            "Test",
            "People",
            "name,age\nAlice,30\nBob,22\nCara,41\nDave,18",
        )
        .expect("database from csv");
        let result = execute_select(
            &database,
            "SELECT name, age FROM People WHERE age >= 22 ORDER BY age DESC LIMIT 2",
        )
        .expect("execute select");
        assert_eq!(result.rows.len(), 2);
        assert_eq!(
            result.rows[0][0],
            lo_core::DbValue::Text("Cara".to_string())
        );
        assert_eq!(
            result.rows[1][0],
            lo_core::DbValue::Text("Alice".to_string())
        );
    }

    #[test]
    fn save_as_handles_all_formats() {
        let db = database_from_csv("Test", "People", "name,age\nAlice,30\nBob,22").unwrap();
        for fmt in ["html", "svg", "pdf", "odb"] {
            let bytes = super::save_as(&db, fmt).unwrap_or_else(|e| panic!("{fmt}: {e}"));
            assert!(!bytes.is_empty(), "{fmt} produced empty output");
        }
        assert!(super::save_as(&db, "qq").is_err());
    }
}
