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
    let employment_page = pages
        .iter()
        .find(|page| page.contains("Swisscom - Engineering Manager"))
        .expect("employment list should be rendered");
    assert!(employment_page.contains("Xapo; Bitfury Holding; Contacts+; Waves Platform"));
    assert!(employment_page.contains("Skyscanner; Wargaming.net"));
}

#[test]
fn heading_is_kept_with_the_following_list() {
    let mut markdown = String::from("# Pagination fixture\n\n");
    for index in 0..34 {
        markdown.push_str(&format!("Filler paragraph {index}.\n\n"));
    }
    markdown.push_str(
        "## Context section\n\n- First context item\n- Second context item\n- Third context item\n",
    );

    let pdf = writer_convert_bytes(markdown.as_bytes(), "md", "pdf")
        .expect("render heading pagination fixture");
    let pages = lo_core::extract_pages_from_pdf(&pdf).expect("extract rendered pages");
    let context_page = pages
        .iter()
        .find(|page| page.contains("Context section"))
        .expect("context heading should be rendered");

    assert!(context_page.contains("First context item"));
    assert!(context_page.contains("Third context item"));
}

#[test]
fn multi_page_tables_repeat_the_header_and_preserve_every_row() {
    let mut markdown = String::from(
        "# Multi-page table\n\n| Company | Amount / Round | Description | Source |\n|---|---|---|---|\n",
    );
    for index in 0..40 {
        markdown.push_str(&format!(
            "| Company {index} | ${index}M Series A | Row {index} has enough descriptive text to wrap consistently. | Source {index} |\n"
        ));
    }

    let pdf = writer_convert_bytes(markdown.as_bytes(), "md", "pdf")
        .expect("render multi-page table fixture");
    let pages = lo_core::extract_pages_from_pdf(&pdf).expect("extract rendered pages");
    assert!(pages.len() >= 2, "expected a multi-page table");
    for page in &pages {
        assert!(page.contains("Company"), "table header missing from page");
        assert!(page.contains("Amount / Round"), "table header missing from page");
    }
    let text = pages.join("\n");
    for index in 0..40 {
        assert!(text.contains(&format!("Company {index}")), "row {index} missing");
        assert!(text.contains(&format!("Source {index}")), "row {index} source missing");
    }
}
