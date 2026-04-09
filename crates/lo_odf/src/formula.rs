use std::path::Path;

use lo_core::{FormulaDocument, Result};
use lo_math::to_mathml_string;

use crate::common::{package_document, MIME_ODF};

pub fn serialize_formula_document(document: &FormulaDocument) -> String {
    // ODF formula documents use math:math as the top-level element of
    // content.xml, not office:document-content. LibreOffice refuses to load a
    // formula package whose content.xml has an office root.
    let mut out = String::new();
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    let body = to_mathml_string(&document.root);
    // to_mathml_string emits a <math:math>…</math:math> element; inject the
    // namespace declaration on the root tag if it isn't already present.
    if let Some(rest) = body.strip_prefix("<math:math>") {
        out.push_str("<math:math xmlns:math=\"http://www.w3.org/1998/Math/MathML\">");
        out.push_str(rest);
    } else if body.starts_with("<math:math ") {
        out.push_str(&body);
    } else {
        out.push_str("<math:math xmlns:math=\"http://www.w3.org/1998/Math/MathML\">");
        out.push_str(&body);
        out.push_str("</math:math>");
    }
    out
}

pub fn save_formula_document(path: impl AsRef<Path>, document: &FormulaDocument) -> Result<()> {
    let content = serialize_formula_document(document);
    package_document(path, MIME_ODF, content, &document.meta, Vec::new())
}
