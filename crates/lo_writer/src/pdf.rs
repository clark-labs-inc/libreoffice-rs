//! PDF export for `TextDocument` using the `layout` engine.
//!
//! `to_pdf` walks the document body and produces a multi-page,
//! block-aware PDF (headings/lists/tables/page-breaks). For callers
//! that just want a single-page plaintext dump there is still
//! `to_pdf_with_size`, which falls through to the same layout
//! engine but on a custom page size.

use lo_core::{units::Length, TextDocument};

use crate::layout::render_document_pdf;

/// Default page is A4 in points (595×842pt). Multi-page output is
/// produced as needed by the layout engine.
pub fn to_pdf(document: &TextDocument) -> Vec<u8> {
    render_document_pdf(document)
}

/// `width`/`height` are accepted for API compatibility but currently
/// ignored — the layout engine always renders at A4. We keep the
/// signature so downstream callers don't have to change.
pub fn to_pdf_with_size(document: &TextDocument, _width: Length, _height: Length) -> Vec<u8> {
    render_document_pdf(document)
}
