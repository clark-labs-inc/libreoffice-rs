use std::collections::BTreeMap;

use crate::meta::Metadata;
use crate::style::CellStyle;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellAddr {
    pub row: u32,
    pub col: u32,
}

impl CellAddr {
    pub fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }

    pub fn from_a1(input: &str) -> crate::Result<Self> {
        if input.is_empty() {
            return Err(crate::LoError::Parse("empty cell address".to_string()));
        }
        let mut letters = String::new();
        let mut digits = String::new();
        for ch in input.chars() {
            if ch.is_ascii_alphabetic() {
                if !digits.is_empty() {
                    return Err(crate::LoError::Parse(format!(
                        "invalid A1 address: {input}"
                    )));
                }
                letters.push(ch.to_ascii_uppercase());
            } else if ch.is_ascii_digit() {
                digits.push(ch);
            } else {
                return Err(crate::LoError::Parse(format!(
                    "invalid A1 address: {input}"
                )));
            }
        }
        if letters.is_empty() || digits.is_empty() {
            return Err(crate::LoError::Parse(format!(
                "invalid A1 address: {input}"
            )));
        }
        let mut col = 0u32;
        for ch in letters.chars() {
            col = col * 26 + ((ch as u8 - b'A') as u32 + 1);
        }
        let row: u32 = digits
            .parse()
            .map_err(|_| crate::LoError::Parse(format!("invalid row in address: {input}")))?;
        if row == 0 || col == 0 {
            return Err(crate::LoError::Parse(format!(
                "invalid A1 address: {input}"
            )));
        }
        Ok(Self {
            row: row - 1,
            col: col - 1,
        })
    }

    pub fn to_a1(self) -> String {
        let mut col = self.col + 1;
        let mut letters = String::new();
        while col > 0 {
            let remainder = ((col - 1) % 26) as u8;
            letters.insert(0, (b'A' + remainder) as char);
            col = (col - 1) / 26;
        }
        format!("{}{}", letters, self.row + 1)
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum CellValue {
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    Formula(String),
    Error(String),
}

impl Default for CellValue {
    fn default() -> Self {
        Self::Empty
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct Cell {
    pub value: CellValue,
    pub style: Option<CellStyle>,
    pub comment: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct Range {
    pub start: CellAddr,
    pub end: CellAddr,
}

impl Range {
    pub fn new(start: CellAddr, end: CellAddr) -> Self {
        Self { start, end }
    }

    pub fn iter(self) -> impl Iterator<Item = CellAddr> {
        let row_start = self.start.row.min(self.end.row);
        let row_end = self.start.row.max(self.end.row);
        let col_start = self.start.col.min(self.end.col);
        let col_end = self.start.col.max(self.end.col);
        (row_start..=row_end)
            .flat_map(move |row| (col_start..=col_end).map(move |col| CellAddr { row, col }))
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct Sheet {
    pub name: String,
    pub cells: BTreeMap<CellAddr, Cell>,
    /// When true, row 0 is treated as a header row: ODF serialization emits
    /// it inside <table:table-header-rows> with a bold style, and calc
    /// evaluation skips it when ranges start at row 0.
    pub has_header: bool,
}

impl Sheet {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: BTreeMap::new(),
            has_header: false,
        }
    }

    pub fn set(&mut self, addr: CellAddr, value: CellValue) {
        self.cells.insert(
            addr,
            Cell {
                value,
                style: None,
                comment: None,
            },
        );
    }

    pub fn get(&self, addr: CellAddr) -> Option<&Cell> {
        self.cells.get(&addr)
    }

    pub fn max_extent(&self) -> (u32, u32) {
        let mut max_row = 0u32;
        let mut max_col = 0u32;
        for addr in self.cells.keys() {
            max_row = max_row.max(addr.row);
            max_col = max_col.max(addr.col);
        }
        (max_row, max_col)
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct NamedRange {
    pub name: String,
    pub sheet: String,
    pub range: Range,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Workbook {
    pub meta: Metadata,
    pub sheets: Vec<Sheet>,
    pub named_ranges: Vec<NamedRange>,
}

impl Default for Workbook {
    fn default() -> Self {
        Self {
            meta: Metadata::default(),
            sheets: vec![Sheet::new("Sheet1")],
            named_ranges: Vec::new(),
        }
    }
}

impl Workbook {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            meta: Metadata::titled(title),
            ..Self::default()
        }
    }

    pub fn sheet(&self, index: usize) -> Option<&Sheet> {
        self.sheets.get(index)
    }

    pub fn sheet_mut(&mut self, index: usize) -> Option<&mut Sheet> {
        self.sheets.get_mut(index)
    }

    pub fn ensure_sheet(&mut self, name: &str) -> &mut Sheet {
        if let Some(index) = self.sheets.iter().position(|sheet| sheet.name == name) {
            return &mut self.sheets[index];
        }
        self.sheets.push(Sheet::new(name));
        self.sheets.last_mut().expect("sheet was just inserted")
    }
}
