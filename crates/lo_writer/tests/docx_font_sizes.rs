use lo_core::{Block, Inline};

#[test]
fn docx_preserves_distinct_run_font_sizes() {
    let document = lo_writer::from_docx_bytes("ignored", include_bytes!("fixtures/font-sizes.docx")).unwrap();
    let sizes: Vec<_> = document.body.iter().filter_map(|block| match block {
        Block::Paragraph(paragraph) => match &paragraph.spans[0] {
            Inline::Styled { style, .. } => Some(style.font_size_pt),
            _ => None,
        },
        _ => None,
    }).collect();
    assert_eq!(sizes, vec![36, 11, 28, 20]);
    let pdf = lo_writer::try_to_pdf(&document).unwrap();
    assert!(pdf.starts_with(b"%PDF-"));
    if let Ok(path) = std::env::var("LO_WRITER_FONT_REPRO_PDF") {
        std::fs::write(path, pdf).unwrap();
    }
}
