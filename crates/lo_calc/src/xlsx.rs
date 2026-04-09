//! Minimal XLSX (SpreadsheetML) export.
//!
//! Produces a self-contained `.xlsx` file with `[Content_Types].xml`,
//! `_rels/.rels`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, and one
//! `xl/worksheets/sheetN.xml` per sheet. Cell values are written as numeric
//! or inline-string ("inlineStr") cells. Formulas are evaluated with
//! [`crate::evaluate_formula`] and the cached result is also written so that
//! consumers without a formula engine still see the right value.

use lo_core::{escape_attr, escape_text, CellAddr, CellValue, Result, Sheet, Workbook};
use lo_zip::{ooxml_package, ZipEntry};

use crate::{evaluate_formula, Value};

const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

fn col_letter(col: u32) -> String {
    let mut col = col + 1;
    let mut letters = String::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        letters.insert(0, (b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    letters
}

fn render_cell(sheet: &Sheet, addr: CellAddr, value: &CellValue) -> String {
    let r = format!("{}{}", col_letter(addr.col), addr.row + 1);
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number(n) => format!("<c r=\"{r}\"><v>{n}</v></c>"),
        CellValue::Bool(b) => format!("<c r=\"{r}\" t=\"b\"><v>{}</v></c>", if *b { 1 } else { 0 }),
        CellValue::Text(t) => format!(
            "<c r=\"{r}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{}</t></is></c>",
            escape_text(t)
        ),
        CellValue::Formula(formula) => {
            let body = formula.strip_prefix('=').unwrap_or(formula);
            match evaluate_formula(formula, sheet) {
                Ok(Value::Number(n)) => {
                    format!("<c r=\"{r}\"><f>{}</f><v>{n}</v></c>", escape_text(body))
                }
                Ok(Value::Bool(b)) => format!(
                    "<c r=\"{r}\" t=\"b\"><f>{}</f><v>{}</v></c>",
                    escape_text(body),
                    if b { 1 } else { 0 }
                ),
                Ok(Value::Text(t)) => format!(
                    "<c r=\"{r}\" t=\"str\"><f>{}</f><v>{}</v></c>",
                    escape_text(body),
                    escape_text(&t)
                ),
                Ok(Value::Blank) => format!("<c r=\"{r}\"><f>{}</f></c>", escape_text(body)),
                Ok(Value::Error(_)) | Err(_) => format!(
                    "<c r=\"{r}\" t=\"e\"><f>{}</f><v>#ERR</v></c>",
                    escape_text(body)
                ),
            }
        }
        CellValue::Error(_) => format!("<c r=\"{r}\" t=\"e\"><v>#ERR</v></c>"),
    }
}

fn render_sheet_xml(sheet: &Sheet) -> String {
    let mut rows_xml = String::new();
    let (max_row, max_col) = sheet.max_extent();
    let dimension = format!("A1:{}{}", col_letter(max_col), max_row + 1);

    // Group cells by row to honor SpreadsheetML's <row> structure.
    let mut current_row: Option<u32> = None;
    let mut row_buf = String::new();
    for (addr, cell) in &sheet.cells {
        if Some(addr.row) != current_row {
            if let Some(r) = current_row {
                rows_xml.push_str(&format!("<row r=\"{}\">{}</row>", r + 1, row_buf));
                row_buf.clear();
            }
            current_row = Some(addr.row);
        }
        row_buf.push_str(&render_cell(sheet, *addr, &cell.value));
    }
    if let Some(r) = current_row {
        rows_xml.push_str(&format!("<row r=\"{}\">{}</row>", r + 1, row_buf));
    }

    format!(
        "{XML_DECL}<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"{dimension}\"/><sheetData>{rows_xml}</sheetData></worksheet>"
    )
}

pub fn to_xlsx(workbook: &Workbook) -> Result<Vec<u8>> {
    let mut entries: Vec<ZipEntry> = Vec::new();

    // Per-sheet content
    let mut sheet_entries = Vec::new();
    for (idx, sheet) in workbook.sheets.iter().enumerate() {
        let body = render_sheet_xml(sheet);
        let path = format!("xl/worksheets/sheet{}.xml", idx + 1);
        entries.push(ZipEntry::new(path.clone(), body.into_bytes()));
        sheet_entries.push((idx + 1, sheet.name.clone(), path));
    }

    // Workbook XML referencing each sheet
    let mut sheets_tag = String::new();
    for (idx, name, _) in &sheet_entries {
        sheets_tag.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{idx}\" r:id=\"rId{idx}\"/>",
            escape_attr(name)
        ));
    }
    let workbook_xml = format!(
        "{XML_DECL}<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>{sheets_tag}</sheets></workbook>"
    );
    entries.push(ZipEntry::new("xl/workbook.xml", workbook_xml.into_bytes()));

    // Workbook relationships pointing at each sheet part
    let mut wb_rels = String::from(
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for (idx, _, path) in &sheet_entries {
        let rel_target = path.trim_start_matches("xl/");
        wb_rels.push_str(&format!(
            "<Relationship Id=\"rId{idx}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"{rel_target}\"/>"
        ));
    }
    wb_rels.push_str("</Relationships>");
    entries.push(ZipEntry::new(
        "xl/_rels/workbook.xml.rels",
        format!("{XML_DECL}{wb_rels}").into_bytes(),
    ));

    // Top-level package parts
    let mut content_types = String::from(
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
<Default Extension=\"xml\" ContentType=\"application/xml\"/>\
<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>",
    );
    for (idx, _, _) in &sheet_entries {
        content_types.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{idx}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        ));
    }
    content_types.push_str("</Types>");
    entries.insert(
        0,
        ZipEntry::new(
            "[Content_Types].xml",
            format!("{XML_DECL}{content_types}").into_bytes(),
        ),
    );

    let rels = format!(
        "{XML_DECL}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rIdWb\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\
</Relationships>"
    );
    entries.insert(1, ZipEntry::new("_rels/.rels", rels.into_bytes()));

    ooxml_package(&entries)
}
