//! End-to-end smoke tests for the high-level convenience helpers.
//! Each test builds a tiny fixture in-memory (no external files) and
//! pushes it through the helper.

use libreoffice_pure::{
    accept_all_tracked_changes_docx_bytes, calc_convert_bytes, docx_to_html_bytes,
    docx_to_pdf_bytes, docx_to_txt_bytes, impress_convert_bytes, ods_to_csv_bytes,
    pptx_to_pdf_bytes, recalc_existing_xlsx_bytes, writer_convert_bytes, xlsx_recalc_bytes,
    xlsx_to_csv_bytes,
};
use lo_core::{CellAddr, CellValue, Workbook};

#[test]
fn docx_to_pdf_round_trip() {
    let doc = lo_writer::from_markdown(
        "demo",
        "# Title\n\nA paragraph with **bold** text.\n\n- one\n- two",
    );
    let docx = lo_writer::to_docx(&doc).expect("docx");
    let pdf = docx_to_pdf_bytes(&docx).expect("docx_to_pdf");
    assert!(pdf.starts_with(b"%PDF"), "pdf header missing");
}

#[test]
fn pptx_to_pdf_round_trip() {
    let deck = lo_impress::demo_presentation("Demo");
    let pptx = lo_impress::to_pptx(&deck).expect("pptx");
    let pdf = pptx_to_pdf_bytes(&pptx).expect("pptx_to_pdf");
    assert!(pdf.starts_with(b"%PDF"), "pdf header missing");
}

#[test]
fn xlsx_recalc_evaluates_formula() {
    // Build a small workbook with a literal and a formula referencing it.
    let mut wb = Workbook::new("recalc");
    let sheet = wb.sheet_mut(0).unwrap();
    sheet.set(CellAddr::new(0, 0), CellValue::Number(40.0));
    sheet.set(CellAddr::new(0, 1), CellValue::Number(2.0));
    sheet.set(CellAddr::new(0, 2), CellValue::Formula("A1+B1".to_string()));
    let xlsx = lo_calc::to_xlsx(&wb).expect("xlsx");
    let recalced = xlsx_recalc_bytes(&xlsx).expect("recalc");

    // The recalculated XLSX should still parse and the formula cell
    // should now carry a cached numeric value of 42 in <v>42</v>.
    let zip = lo_zip::ZipArchive::new(&recalced).expect("read recalced zip");
    assert!(zip.contains("xl/workbook.xml"));
    assert!(!zip.contains("xl/calcChain.xml"));
    let sheet_xml = zip.read_string("xl/worksheets/sheet1.xml").expect("sheet1");
    assert!(
        sheet_xml.contains("<v>42</v>") || sheet_xml.contains("<v>42.0</v>"),
        "expected formula cache: {sheet_xml}"
    );

    // Round-trip through from_xlsx should still see the formula.
    let reloaded = lo_calc::from_xlsx_bytes("rt", &recalced).expect("reload");
    let cell = reloaded
        .sheet(0)
        .and_then(|s| s.get(CellAddr::new(0, 2)))
        .expect("cell");
    assert!(matches!(cell.value, CellValue::Formula(_)));
}

#[test]
fn accept_changes_strips_revisions() {
    // Hand-build a tiny DOCX with a w:ins (insertion) and w:del (deletion)
    // and verify the helper accepts the insertion and drops the deletion.
    use lo_zip::{ooxml_package, ZipEntry};
    let content_types = r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document = r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Kept </w:t></w:r><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>inserted </w:t></w:r></w:ins><w:del w:id="2" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:delText>removed </w:delText></w:r></w:del><w:r><w:t>tail.</w:t></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_package(&[
        ZipEntry::new("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ZipEntry::new("_rels/.rels", rels.as_bytes().to_vec()),
        ZipEntry::new("word/document.xml", document.as_bytes().to_vec()),
    ])
    .expect("build docx");

    let accepted = accept_all_tracked_changes_docx_bytes(&bytes).expect("accept");
    let zip = lo_zip::ZipArchive::new(&accepted).expect("zip");
    let document_xml = zip.read_string("word/document.xml").expect("doc");
    assert!(document_xml.contains("Kept "), "kept text missing");
    assert!(document_xml.contains("inserted "), "inserted text missing");
    assert!(document_xml.contains("tail."), "tail missing");
    assert!(!document_xml.contains("removed"), "deletion not stripped");
    assert!(!document_xml.contains("<w:del"), "w:del not stripped");
    assert!(!document_xml.contains("<w:ins"), "w:ins not unwrapped");
    // Re-importing with the writer should also see the merged sentence.
    let reloaded = lo_writer::from_docx_bytes("d", &accepted).expect("reload");
    let plain = reloaded.plain_text();
    assert!(
        plain.contains("Kept inserted"),
        "merged text missing: {plain}"
    );
    assert!(
        !plain.contains("removed"),
        "deletion present in text: {plain}"
    );
}

#[test]
fn accept_changes_strips_settings_track_revisions() {
    use lo_zip::{ooxml_package, ZipEntry};
    let content_types = r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>"#;
    let rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document = r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:body></w:document>"#;
    // settings.xml turns track-changes ON via <w:trackRevisions/>.
    let settings = r#"<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:trackRevisions/><w:zoom w:percent="100"/></w:settings>"#;
    let bytes = ooxml_package(&[
        ZipEntry::new("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ZipEntry::new("_rels/.rels", rels.as_bytes().to_vec()),
        ZipEntry::new("word/document.xml", document.as_bytes().to_vec()),
        ZipEntry::new("word/settings.xml", settings.as_bytes().to_vec()),
    ])
    .expect("build docx");

    let accepted = accept_all_tracked_changes_docx_bytes(&bytes).expect("accept");
    let zip = lo_zip::ZipArchive::new(&accepted).expect("zip");
    let settings_xml = zip.read_string("word/settings.xml").expect("settings");
    assert!(
        !settings_xml.contains("trackRevisions"),
        "trackRevisions toggle was not removed from settings.xml: {settings_xml}"
    );
    // Other settings (zoom) should survive.
    assert!(
        settings_xml.contains("zoom"),
        "unrelated settings were dropped: {settings_xml}"
    );
}

#[test]
fn writer_convert_bytes_round_trips_html() {
    let doc = lo_writer::from_markdown("c", "# T\n\nhello **world**");
    let docx = lo_writer::to_docx(&doc).expect("docx");
    let html = docx_to_html_bytes(&docx).expect("docx -> html");
    let html = String::from_utf8(html).expect("utf8");
    assert!(html.contains("<h1>"), "no <h1>: {html}");
    assert!(html.contains("</h1>"), "no </h1>: {html}");
    assert!(html.contains("world"), "no world: {html}");

    let txt = docx_to_txt_bytes(&docx).expect("docx -> txt");
    let txt = String::from_utf8(txt).expect("utf8");
    assert!(txt.contains("hello"));

    // Generic dispatcher with explicit format hints.
    let svg = writer_convert_bytes(&docx, "docx", "svg").expect("docx -> svg");
    assert!(svg.starts_with(b"<svg"));
}

#[test]
fn calc_convert_bytes_extracts_csv() {
    let mut wb = Workbook::new("c");
    let sheet = wb.sheet_mut(0).unwrap();
    sheet.set(CellAddr::new(0, 0), CellValue::Text("a".to_string()));
    sheet.set(CellAddr::new(0, 1), CellValue::Text("b".to_string()));
    sheet.set(CellAddr::new(1, 0), CellValue::Number(1.0));
    sheet.set(CellAddr::new(1, 1), CellValue::Number(2.0));
    let xlsx = lo_calc::to_xlsx(&wb).expect("xlsx");
    let csv = xlsx_to_csv_bytes(&xlsx).expect("xlsx -> csv");
    let csv = String::from_utf8(csv).expect("utf8");
    assert!(csv.contains("a,b"));
    assert!(csv.contains("1,2") || csv.contains("1.0,2.0"));

    // Same thing through the generic dispatcher.
    let html = calc_convert_bytes(&xlsx, "xlsx", "html").expect("xlsx -> html");
    assert!(html.starts_with(b"<!"));
}

#[test]
fn impress_convert_bytes_renders_html() {
    let deck = lo_impress::demo_presentation("Demo");
    let pptx = lo_impress::to_pptx(&deck).expect("pptx");
    let html = impress_convert_bytes(&pptx, "pptx", "html").expect("pptx -> html");
    assert!(html.starts_with(b"<!"));
}

#[test]
fn aliases_match_canonical_helpers() {
    // recalc_existing_xlsx_bytes is just an alias for xlsx_recalc_bytes.
    let mut wb = Workbook::new("c");
    let sheet = wb.sheet_mut(0).unwrap();
    sheet.set(CellAddr::new(0, 0), CellValue::Number(1.0));
    sheet.set(CellAddr::new(0, 1), CellValue::Formula("A1+1".to_string()));
    let xlsx = lo_calc::to_xlsx(&wb).expect("xlsx");
    let alias_out = recalc_existing_xlsx_bytes(&xlsx).expect("recalc alias");
    let canonical_out = xlsx_recalc_bytes(&xlsx).expect("recalc canonical");
    // Both should produce identical bytes (the function is just a re-export).
    assert_eq!(alias_out, canonical_out);
}

#[test]
fn ods_to_csv_round_trip() {
    let mut wb = Workbook::new("c");
    let sheet = wb.sheet_mut(0).unwrap();
    sheet.set(CellAddr::new(0, 0), CellValue::Text("hello".to_string()));
    sheet.set(CellAddr::new(0, 1), CellValue::Number(123.0));
    let tmp = std::env::temp_dir().join("lo_ods_helper_test.ods");
    lo_odf::save_spreadsheet_document(&tmp, &wb).expect("save ods");
    let bytes = std::fs::read(&tmp).expect("read");
    let _ = std::fs::remove_file(&tmp);
    let csv = ods_to_csv_bytes(&bytes).expect("ods -> csv");
    let csv = String::from_utf8(csv).expect("utf8");
    assert!(csv.contains("hello"));
    assert!(csv.contains("123"));
}
