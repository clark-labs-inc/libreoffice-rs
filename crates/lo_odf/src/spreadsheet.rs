use std::path::Path;

use lo_core::{CellAddr, CellValue, Result, Workbook};

use crate::common::{content_root_attrs, package_document, MIME_ODS};

fn odf_formula(formula: &str) -> String {
    if formula.starts_with("of:=") {
        formula.to_string()
    } else if formula.starts_with('=') {
        format!("of:{formula}")
    } else {
        format!("of:={formula}")
    }
}

fn cell_xml(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "<table:table-cell/>".to_string(),
        CellValue::Number(number) => format!(
            "<table:table-cell office:value-type=\"float\" office:value=\"{number}\"><text:p>{number}</text:p></table:table-cell>"
        ),
        CellValue::Text(text) => format!(
            "<table:table-cell office:value-type=\"string\"><text:p>{}</text:p></table:table-cell>",
            lo_core::escape_text(text)
        ),
        CellValue::Bool(value) => format!(
            "<table:table-cell table:style-name=\"ceBool\" office:value-type=\"boolean\" office:boolean-value=\"{}\"><text:p>{}</text:p></table:table-cell>",
            value,
            if *value { "TRUE" } else { "FALSE" }
        ),
        CellValue::Formula(formula) => format!(
            "<table:table-cell table:formula=\"{}\" office:value-type=\"string\"><text:p>{}</text:p></table:table-cell>",
            lo_core::escape_attr(&odf_formula(formula)),
            lo_core::escape_text(formula)
        ),
        CellValue::Error(message) => format!(
            "<table:table-cell office:value-type=\"string\"><text:p>#ERR {}</text:p></table:table-cell>",
            lo_core::escape_text(message)
        ),
    }
}

pub fn serialize_spreadsheet_document(book: &Workbook) -> String {
    let mut xml = lo_core::XmlBuilder::new();
    xml.declaration();
    xml.open("office:document-content", &content_root_attrs());
    xml.empty("office:scripts", &[]);
    // Automatic styles: a boolean display style so true/false render as
    // TRUE/FALSE instead of 0/1, a header cell style, and a reasonable
    // default column width so LibreOffice doesn't collapse everything to
    // the driver default.
    xml.open("office:automatic-styles", &[]);
    xml.raw(
        "<number:boolean-style style:name=\"NBool\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\">\
<number:boolean/></number:boolean-style>",
    );
    xml.raw(
        "<style:style style:name=\"ceBool\" style:family=\"table-cell\" style:data-style-name=\"NBool\"/>",
    );
    xml.raw(
        "<style:style style:name=\"ceHeader\" style:family=\"table-cell\">\
<style:text-properties fo:font-weight=\"bold\"/>\
<style:table-cell-properties fo:background-color=\"#eeeeee\"/>\
</style:style>",
    );
    xml.raw(
        "<style:style style:name=\"co1\" style:family=\"table-column\">\
<style:table-column-properties style:column-width=\"25mm\"/>\
</style:style>",
    );
    xml.close();
    xml.open("office:body", &[]);
    xml.open("office:spreadsheet", &[]);
    for sheet in &book.sheets {
        xml.raw(&format!(
            "<table:table table:name=\"{}\">",
            lo_core::escape_attr(&sheet.name)
        ));
        let (max_row, max_col) = sheet.max_extent();
        for _ in 0..=max_col {
            xml.raw("<table:table-column table:style-name=\"co1\"/>");
        }
        if sheet.has_header {
            xml.raw("<table:table-header-rows>");
            xml.raw("<table:table-row>");
            for col in 0..=max_col {
                let addr = CellAddr::new(0, col);
                let value = sheet
                    .get(addr)
                    .map(|c| &c.value)
                    .unwrap_or(&CellValue::Empty);
                let text = match value {
                    CellValue::Text(t) => lo_core::escape_text(t),
                    CellValue::Number(n) => n.to_string(),
                    CellValue::Bool(b) => b.to_string(),
                    _ => String::new(),
                };
                xml.raw(&format!(
                    "<table:table-cell table:style-name=\"ceHeader\" office:value-type=\"string\"><text:p>{text}</text:p></table:table-cell>"
                ));
            }
            xml.raw("</table:table-row>");
            xml.raw("</table:table-header-rows>");
        }
        let start_row = if sheet.has_header { 1 } else { 0 };
        for row in start_row..=max_row {
            xml.raw("<table:table-row>");
            for col in 0..=max_col {
                let addr = CellAddr::new(row, col);
                let cell = sheet
                    .get(addr)
                    .map(|cell| &cell.value)
                    .unwrap_or(&CellValue::Empty);
                xml.raw(&cell_xml(cell));
            }
            xml.raw("</table:table-row>");
        }
        xml.raw("</table:table>");
    }
    xml.close();
    xml.close();
    xml.close();
    xml.finish()
}

pub fn save_spreadsheet_document(path: impl AsRef<Path>, book: &Workbook) -> Result<()> {
    let content = serialize_spreadsheet_document(book);
    package_document(path, MIME_ODS, content, &book.meta, Vec::new())
}
