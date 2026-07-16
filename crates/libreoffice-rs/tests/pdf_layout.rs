use libreoffice_pure::writer_convert_bytes;

const REGRESSION_MARKDOWN: &str = include_str!("fixtures/pdf_layout_regression.md");

#[test]
fn markdown_pdf_has_no_synthetic_title_or_style_spacing_damage() {
    let pdf = writer_convert_bytes(REGRESSION_MARKDOWN.as_bytes(), "md", "pdf")
        .expect("render regression markdown");
    let pages = lo_core::extract_pages_from_pdf(&pdf).expect("extract rendered pages");
    let text = pages.join("\n");
    let normalized = text.split_whitespace().collect::<Vec<_>>().join(" ");

    assert!(!text.trim_start().starts_with("document"));
    assert!(text.trim_start().starts_with("Alex Ponomarev"));
    assert!(normalized.contains("Prepared: July 16, 2026 Scope: public-source research"));
    assert!(normalized.contains("bold evidence labels, italic source qualifications"));
    assert!(normalized.contains("Consulting brand. Positions"));
    assert!(normalized.contains("“curly quotes,”"));
    assert!(normalized.contains("José"));
    assert!(normalized.contains("Zürich"));
}

#[test]
fn short_paragraphs_and_lists_are_not_split_across_pages() {
    let pdf = writer_convert_bytes(REGRESSION_MARKDOWN.as_bytes(), "md", "pdf")
        .expect("render regression markdown");
    let pages = lo_core::extract_pages_from_pdf(&pdf).expect("extract rendered pages");

    assert_eq!(pages.len(), 2);
    assert!(!pages[1].trim_start().starts_with("rather than established facts"));
    assert!(pages[1].contains("Swisscom - Engineering Manager"));
    assert!(pages[1].contains("Skyscanner; Wargaming.net"));
}
