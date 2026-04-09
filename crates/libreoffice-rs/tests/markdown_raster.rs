use libreoffice_pure::{
    docx_to_jpeg_pages, docx_to_md_bytes, docx_to_png_pages, pptx_to_jpeg_pages,
    pptx_to_md_bytes, pptx_to_png_pages, xlsx_to_md_bytes,
};

#[test]
fn docx_markdown_extraction_contains_title_and_bold_text() {
    let doc = lo_writer::from_markdown("Doc", "# Title\n\nA **bold** paragraph.");
    let bytes = lo_writer::to_docx(&doc).expect("docx");
    let md = String::from_utf8(docx_to_md_bytes(&bytes).expect("md")).expect("utf8");
    assert!(md.contains("Title"));
    assert!(md.contains("bold"));
}

#[test]
fn pptx_markdown_extraction_contains_slide_text() {
    let mut builder = lo_impress::ImpressBuilder::new("Deck");
    builder.add_title_slide("Hello", "world");
    let bytes = lo_impress::to_pptx(&builder.presentation).expect("pptx");
    let md = String::from_utf8(pptx_to_md_bytes(&bytes).expect("md")).expect("utf8");
    assert!(md.contains("Slide 1"));
    assert!(md.contains("Hello"));
}

#[test]
fn xlsx_markdown_extraction_contains_cells() {
    let workbook = lo_calc::workbook_from_csv("Book", "Sheet1", "name,value\nalpha,42\n").expect("csv");
    let bytes = lo_calc::to_xlsx(&workbook).expect("xlsx");
    let md = String::from_utf8(xlsx_to_md_bytes(&bytes).expect("md")).expect("utf8");
    assert!(md.contains("name"));
    assert!(md.contains("value"));
    assert!(md.contains("alpha"));
    assert!(md.contains("42"));
}

#[test]
fn docx_png_raster_output_has_png_signature() {
    let doc = lo_writer::from_markdown("Doc", "# Title\n\nA paragraph.");
    let bytes = lo_writer::to_docx(&doc).expect("docx");
    let pages = docx_to_png_pages(&bytes, 96).expect("png pages");
    assert!(!pages.is_empty());
    assert!(pages[0].starts_with(b"\x89PNG\r\n\x1a\n"));
}

#[test]
fn pptx_png_raster_output_has_png_signature() {
    let mut builder = lo_impress::ImpressBuilder::new("Deck");
    builder.add_title_slide("Slide", "subtitle");
    let bytes = lo_impress::to_pptx(&builder.presentation).expect("pptx");
    let pages = pptx_to_png_pages(&bytes, 96).expect("png pages");
    assert_eq!(pages.len(), 1);
    assert!(pages[0].starts_with(b"\x89PNG\r\n\x1a\n"));
}

#[test]
fn jpeg_raster_output_has_jpeg_markers() {
    let doc = lo_writer::from_markdown("Doc", "# Title\n\nA paragraph.");
    let docx = lo_writer::to_docx(&doc).expect("docx");
    let doc_pages = docx_to_jpeg_pages(&docx, 96, 80).expect("jpg pages");
    assert!(!doc_pages.is_empty());
    assert!(doc_pages[0].starts_with(&[0xFF, 0xD8]));
    assert!(doc_pages[0].ends_with(&[0xFF, 0xD9]));

    let mut builder = lo_impress::ImpressBuilder::new("Deck");
    builder.add_title_slide("Slide", "subtitle");
    let pptx = lo_impress::to_pptx(&builder.presentation).expect("pptx");
    let ppt_pages = pptx_to_jpeg_pages(&pptx, 96, 80).expect("jpg pages");
    assert!(ppt_pages[0].starts_with(&[0xFF, 0xD8]));
    assert!(ppt_pages[0].ends_with(&[0xFF, 0xD9]));
}
