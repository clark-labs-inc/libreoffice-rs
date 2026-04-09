use libreoffice_pure::{
    convert_bytes, convert_bytes_auto, pdf_to_html_bytes, pdf_to_md_bytes, pdf_to_txt_bytes,
};
use lo_core::{pdf_from_objects, write_text_pdf, Length};

fn simple_tounicode_pdf() -> Vec<u8> {
    let cmap = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n5 beginbfchar\n<0048> <0048>\n<0065> <0065>\n<006C> <006C>\n<006F> <006F>\n<0020> <0020>\nendbfchar\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n";
    let contents = "BT\n/F1 12 Tf\n1 0 0 1 50 200 Tm\n<00480065006C006C006F> Tj\nET\n";
    let objects = vec![
        "<< /Type /Catalog /Pages 2 0 R >>".to_string(),
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>".to_string(),
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>".to_string(),
        format!("<< /Length {} >>\nstream\n{}endstream", contents.len(), contents),
        "<< /Type /Font /Subtype /Type0 /BaseFont /Faux /Encoding /Identity-H /DescendantFonts [] /ToUnicode 6 0 R >>".to_string(),
        format!("<< /Length {} >>\nstream\n{}endstream", cmap.len(), cmap),
    ];
    pdf_from_objects(&objects)
}

#[test]
fn pdf_to_txt_extracts_text() {
    let pdf = write_text_pdf(&["Hello from PDF".to_string()], Length::pt(300.0), Length::pt(300.0));
    let txt = String::from_utf8(pdf_to_txt_bytes(&pdf).unwrap()).unwrap();
    assert!(txt.contains("Hello from PDF"));
}

#[test]
fn convert_bytes_auto_detects_pdf() {
    let pdf = write_text_pdf(&["Auto detect".to_string()], Length::pt(200.0), Length::pt(200.0));
    let txt = String::from_utf8(convert_bytes_auto(&pdf, "txt").unwrap()).unwrap();
    assert!(txt.contains("Auto detect"));
}

#[test]
fn pdf_to_markdown_and_html_use_writer_pipeline() {
    let pdf = write_text_pdf(&["Hello markdown".to_string()], Length::pt(200.0), Length::pt(200.0));
    let md = String::from_utf8(pdf_to_md_bytes(&pdf).unwrap()).unwrap();
    let html = String::from_utf8(pdf_to_html_bytes(&pdf).unwrap()).unwrap();
    assert!(md.contains("Hello markdown"));
    assert!(html.to_ascii_lowercase().contains("hello markdown"));
}

#[test]
fn generic_convert_accepts_pdf_source_hint() {
    let pdf = write_text_pdf(&["Hinted convert".to_string()], Length::pt(200.0), Length::pt(200.0));
    let txt = String::from_utf8(convert_bytes(&pdf, "pdf", "txt").unwrap()).unwrap();
    assert!(txt.contains("Hinted convert"));
}

#[test]
fn tounicode_hex_strings_round_trip() {
    let pdf = simple_tounicode_pdf();
    let txt = String::from_utf8(pdf_to_txt_bytes(&pdf).unwrap()).unwrap();
    assert!(txt.contains("Hello"));
}
