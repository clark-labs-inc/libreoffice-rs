use crate::meta::Metadata;

#[derive(Clone, Debug, PartialEq)]
pub enum DbValue {
    Null,
    Integer(i64),
    Float(f64),
    Text(String),
    Bool(bool),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ColumnType {
    Integer,
    Float,
    Text,
    Bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ColumnDef {
    pub name: String,
    pub column_type: ColumnType,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct TableData {
    pub name: String,
    pub columns: Vec<ColumnDef>,
    pub rows: Vec<Vec<DbValue>>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct QueryDef {
    pub name: String,
    pub sql: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct DatabaseDocument {
    pub meta: Metadata,
    pub tables: Vec<TableData>,
    pub queries: Vec<QueryDef>,
}

impl Default for DatabaseDocument {
    fn default() -> Self {
        Self {
            meta: Metadata::default(),
            tables: Vec::new(),
            queries: Vec::new(),
        }
    }
}

impl DatabaseDocument {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            meta: Metadata::titled(title),
            ..Self::default()
        }
    }

    pub fn table(&self, name: &str) -> Option<&TableData> {
        self.tables
            .iter()
            .find(|table| table.name.eq_ignore_ascii_case(name))
    }
}
