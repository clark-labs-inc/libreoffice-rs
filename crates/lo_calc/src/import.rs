//! Binary importers for `Workbook`.
//!
//! Supports `csv`, `xlsx`, `ods` format hints.

use std::collections::BTreeMap;

use lo_core::{
    parse_xml_document, CellAddr, CellValue, LoError, Result, Sheet, Workbook, XmlItem, XmlNode,
};
use lo_zip::{rels_path_for, resolve_part_target, ZipArchive};

use crate::workbook_from_csv;

pub fn load_bytes(title: impl Into<String>, bytes: &[u8], format: &str) -> Result<Workbook> {
    let title = title.into();
    match format.to_ascii_lowercase().as_str() {
        "csv" => workbook_from_csv(title.clone(), &title, &bytes_to_utf8(bytes)?),
        "xlsx" => from_xlsx_bytes(title, bytes),
        "ods" => from_ods_bytes(title, bytes),
        other => Err(LoError::Unsupported(format!("calc import format {other}"))),
    }
}

fn bytes_to_utf8(bytes: &[u8]) -> Result<String> {
    String::from_utf8(bytes.to_vec())
        .map_err(|err| LoError::Parse(format!("invalid utf-8 input: {err}")))
}

pub fn from_xlsx_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<Workbook> {
    let _ = title.into();
    let zip = ZipArchive::new(bytes)?;
    // Pull real title from docProps/core.xml; skip the placeholder caller
    // so synthesized "workbook" titles never bleed into Markdown / PDFs.
    let title: String = if zip.contains("docProps/core.xml") {
        parse_xml_document(&zip.read_string("docProps/core.xml")?)
            .ok()
            .and_then(|root| root.child("title").map(|n| n.text_content()))
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
            .unwrap_or_default()
    } else {
        String::new()
    };
    let workbook_root = parse_xml_document(&zip.read_string("xl/workbook.xml")?)?;
    let rels = parse_relationships(&zip, "xl/workbook.xml")?;
    let shared_strings = if zip.contains("xl/sharedStrings.xml") {
        parse_shared_strings(&parse_xml_document(
            &zip.read_string("xl/sharedStrings.xml")?,
        )?)
    } else {
        Vec::new()
    };

    let mut workbook = Workbook::new(title);
    workbook.sheets.clear();

    if let Some(sheets_node) = workbook_root.child("sheets") {
        for (index, sheet_node) in sheets_node.children_named("sheet").enumerate() {
            let name = sheet_node.attr("name").unwrap_or("Sheet").to_string();
            let target = sheet_node
                .attr("id")
                .or_else(|| sheet_node.attr("r:id"))
                .and_then(|id| rels.get(id))
                .cloned()
                .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", index + 1));
            if !zip.contains(&target) {
                continue;
            }
            let sheet_xml = parse_xml_document(&zip.read_string(&target)?)?;
            workbook
                .sheets
                .push(parse_xlsx_sheet(&name, &sheet_xml, &shared_strings)?);
        }
    }
    if workbook.sheets.is_empty() {
        workbook.sheets.push(Sheet::new("Sheet1"));
    }
    Ok(workbook)
}

pub fn from_ods_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<Workbook> {
    let title = title.into();
    let zip = ZipArchive::new(bytes)?;
    let content = parse_xml_document(&zip.read_string("content.xml")?)?;
    let body = content
        .child("body")
        .ok_or_else(|| LoError::Parse("content.xml missing office:body".to_string()))?;
    let spreadsheet = body
        .child("spreadsheet")
        .ok_or_else(|| LoError::Parse("content.xml missing office:spreadsheet".to_string()))?;

    let mut workbook = Workbook::new(title);
    workbook.sheets.clear();
    for table in spreadsheet.children_named("table") {
        workbook.sheets.push(parse_ods_sheet(table));
    }
    if workbook.sheets.is_empty() {
        workbook.sheets.push(Sheet::new("Sheet1"));
    }
    Ok(workbook)
}

// ---------------------------------------------------------------------------

fn parse_relationships(zip: &ZipArchive, part: &str) -> Result<BTreeMap<String, String>> {
    let rels_path = rels_path_for(part);
    if !zip.contains(&rels_path) {
        return Ok(BTreeMap::new());
    }
    let root = parse_xml_document(&zip.read_string(&rels_path)?)?;
    let mut map = BTreeMap::new();
    for rel in root.children_named("Relationship") {
        if let (Some(id), Some(target)) = (rel.attr("Id"), rel.attr("Target")) {
            map.insert(id.to_string(), resolve_part_target(part, target));
        }
    }
    Ok(map)
}

fn parse_shared_strings(root: &XmlNode) -> Vec<String> {
    let mut strings = Vec::new();
    for si in root.children_named("si") {
        strings.push(xlsx_text(si));
    }
    strings
}

fn xlsx_text(node: &XmlNode) -> String {
    let mut out = String::new();
    for item in &node.items {
        match item {
            XmlItem::Text(text) => out.push_str(text),
            XmlItem::Node(child) => match child.local_name() {
                "t" => out.push_str(&child.text_content()),
                _ => out.push_str(&xlsx_text(child)),
            },
        }
    }
    out
}

fn parse_xlsx_sheet(name: &str, root: &XmlNode, shared_strings: &[String]) -> Result<Sheet> {
    let mut sheet = Sheet::new(name.to_string());
    let Some(sheet_data) = root.child("sheetData") else {
        return Ok(sheet);
    };
    let mut next_row: u32 = 1;
    for row_node in sheet_data.children_named("row") {
        let row_number = row_node
            .attr("r")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(next_row);
        let mut next_col: u32 = 1;
        for cell in row_node.children_named("c") {
            let reference = cell.attr("r");
            let (cell_row, cell_col) = reference
                .and_then(parse_cell_ref)
                .unwrap_or((row_number, next_col));
            next_col = cell_col + 1;
            let formula = cell.child("f").map(|node| node.text_content());
            let value = if let Some(formula) = formula {
                let trimmed = formula.trim();
                if trimmed.is_empty() {
                    parse_xlsx_value(cell, shared_strings)
                } else {
                    CellValue::Formula(trimmed.to_string())
                }
            } else {
                parse_xlsx_value(cell, shared_strings)
            };
            if value != CellValue::Empty {
                sheet.set(
                    CellAddr::new(cell_row.saturating_sub(1), cell_col.saturating_sub(1)),
                    value,
                );
            }
        }
        next_row = row_number + 1;
    }
    Ok(sheet)
}

fn parse_xlsx_value(cell: &XmlNode, shared_strings: &[String]) -> CellValue {
    let cell_type = cell.attr("t").unwrap_or("");
    let value_text = cell
        .child("v")
        .map(|node| node.text_content())
        .unwrap_or_default();
    match cell_type {
        "s" => {
            let index = value_text.parse::<usize>().unwrap_or(0);
            CellValue::Text(shared_strings.get(index).cloned().unwrap_or_default())
        }
        "inlineStr" => CellValue::Text(cell.child("is").map(xlsx_text).unwrap_or_default()),
        "str" => CellValue::Text(value_text),
        "b" => CellValue::Bool(matches!(value_text.as_str(), "1" | "true" | "TRUE")),
        "e" => CellValue::Error(value_text),
        _ => {
            if value_text.is_empty() {
                CellValue::Empty
            } else if let Ok(number) = value_text.parse::<f64>() {
                CellValue::Number(number)
            } else {
                CellValue::Text(value_text)
            }
        }
    }
}

/// Parse a cell reference like "A1" -> (row, col) (1-based).
fn parse_cell_ref(reference: &str) -> Option<(u32, u32)> {
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in reference.chars() {
        if ch.is_ascii_alphabetic() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let mut col: u32 = 0;
    for ch in letters.chars() {
        col = col * 26 + ((ch as u8 - b'A') as u32 + 1);
    }
    let row = digits.parse::<u32>().ok()?;
    Some((row, col))
}

// ---------------------------------------------------------------------------
// ODS

fn parse_ods_sheet(node: &XmlNode) -> Sheet {
    let name = node.attr("name").unwrap_or("Sheet1").to_string();
    let mut sheet = Sheet::new(name);
    let mut row_index: u32 = 1;
    for row_node in node.children_named("table-row") {
        let repeat_rows = row_node
            .attr("number-rows-repeated")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(1);
        let row_values = parse_ods_row(row_node);
        for _ in 0..repeat_rows {
            for (col, value) in &row_values {
                if *value != CellValue::Empty {
                    sheet.set(
                        CellAddr::new(row_index.saturating_sub(1), col.saturating_sub(1)),
                        value.clone(),
                    );
                }
            }
            row_index += 1;
        }
    }
    sheet
}

fn parse_ods_row(row_node: &XmlNode) -> Vec<(u32, CellValue)> {
    let mut values = Vec::new();
    let mut col_index: u32 = 1;
    for cell in row_node
        .children
        .iter()
        .filter(|child| matches!(child.local_name(), "table-cell" | "covered-table-cell"))
    {
        let repeat_cols = cell
            .attr("number-columns-repeated")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(1);
        let value = if cell.local_name() == "covered-table-cell" {
            CellValue::Empty
        } else {
            parse_ods_value(cell)
        };
        for _ in 0..repeat_cols {
            values.push((col_index, value.clone()));
            col_index += 1;
        }
    }
    values
}

fn parse_ods_value(cell: &XmlNode) -> CellValue {
    if let Some(formula) = cell.attr("formula") {
        return CellValue::Formula(
            formula
                .trim_start_matches("of:=")
                .trim_start_matches('=')
                .to_string(),
        );
    }
    let value_type = cell.attr("value-type").unwrap_or("string");
    let text = cell
        .children
        .iter()
        .filter(|child| matches!(child.local_name(), "p" | "h"))
        .map(|paragraph| paragraph.text_content())
        .collect::<Vec<_>>()
        .join("\n");
    match value_type {
        "float" | "percentage" | "currency" => cell
            .attr("value")
            .and_then(|value| value.parse::<f64>().ok())
            .map(CellValue::Number)
            .unwrap_or_else(|| CellValue::Text(text)),
        "boolean" => CellValue::Bool(matches!(cell.attr("boolean-value"), Some("true" | "1"))),
        "date" | "time" => CellValue::Text(
            cell.attr("date-value")
                .or_else(|| cell.attr("time-value"))
                .unwrap_or(&text)
                .to_string(),
        ),
        "string" => {
            if text.is_empty() {
                CellValue::Empty
            } else {
                CellValue::Text(text)
            }
        }
        _ => {
            if text.is_empty() {
                CellValue::Empty
            } else {
                CellValue::Text(text)
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::to_xlsx;
    use lo_core::{CellAddr, CellValue, Workbook};

    #[test]
    fn parse_cell_ref_handles_basic() {
        assert_eq!(parse_cell_ref("A1"), Some((1, 1)));
        assert_eq!(parse_cell_ref("B2"), Some((2, 2)));
        assert_eq!(parse_cell_ref("AA10"), Some((10, 27)));
    }

    #[test]
    fn xlsx_round_trip_imports_cells() {
        let mut wb = Workbook::new("book");
        let sheet = wb.sheet_mut(0).unwrap();
        sheet.set(CellAddr::new(0, 0), CellValue::Text("a".to_string()));
        sheet.set(CellAddr::new(0, 1), CellValue::Number(42.0));
        sheet.set(CellAddr::new(1, 0), CellValue::Formula("A1".to_string()));
        let bytes = to_xlsx(&wb).expect("xlsx");
        let loaded = from_xlsx_bytes("book", &bytes).expect("import xlsx");
        let sheet = loaded.sheet(0).expect("sheet");
        assert_eq!(
            sheet.get(CellAddr::new(0, 0)).map(|c| c.value.clone()),
            Some(CellValue::Text("a".to_string()))
        );
        assert_eq!(
            sheet.get(CellAddr::new(0, 1)).map(|c| c.value.clone()),
            Some(CellValue::Number(42.0))
        );
    }

    #[test]
    fn ods_round_trip_imports_cells() {
        let mut wb = Workbook::new("book");
        let sheet = wb.sheet_mut(0).unwrap();
        sheet.set(CellAddr::new(0, 0), CellValue::Text("a".to_string()));
        let tmp = std::env::temp_dir().join("lo_calc_import_test.ods");
        lo_odf::save_spreadsheet_document(&tmp, &wb).expect("save ods");
        let bytes = std::fs::read(&tmp).expect("read");
        let _ = std::fs::remove_file(&tmp);
        let loaded = from_ods_bytes("book", &bytes).expect("import ods");
        let sheet = loaded.sheet(0).expect("sheet");
        assert_eq!(
            sheet.get(CellAddr::new(0, 0)).map(|c| c.value.clone()),
            Some(CellValue::Text("a".to_string()))
        );
    }
}
