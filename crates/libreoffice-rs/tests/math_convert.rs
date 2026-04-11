use libreoffice_pure::math_convert_bytes;
use lo_zip::ZipArchive;

/// `convert --to odf` on a math input was hitting
/// `lo_math::save_as` which does not know about ODF packaging. Route it
/// through `lo_odf::save_formula_document_bytes` so the generic convert
/// router produces a valid ODF formula archive that real LibreOffice can
/// still open as a Math document.
#[test]
fn latex_to_odf_produces_valid_math_archive() {
    let latex = br"\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}";
    let bytes = math_convert_bytes(latex, "latex", "odf").expect("latex -> odf");
    let zip = ZipArchive::new(&bytes).expect("valid zip");
    let mimetype = zip.read_string("mimetype").expect("mimetype");
    assert_eq!(
        mimetype.trim(),
        "application/vnd.oasis.opendocument.formula"
    );
    let content = zip.read_string("content.xml").expect("content.xml");
    assert!(content.contains("math:math"));
    assert!(content.contains("math:mfrac") || content.contains("math:mi"));
}

#[test]
fn mathml_to_odf_roundtrip() {
    let mathml = br#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi><mo>=</mo><mn>42</mn></math>"#;
    let bytes = math_convert_bytes(mathml, "mathml", "odf").expect("mathml -> odf");
    let zip = ZipArchive::new(&bytes).expect("valid zip");
    let content = zip.read_string("content.xml").expect("content.xml");
    assert!(content.contains("math:math"));
}

#[test]
fn odf_to_odf_roundtrip() {
    let latex = br"x^2 + y^2 = r^2";
    let initial = math_convert_bytes(latex, "latex", "odf").expect("latex -> odf");
    let round = math_convert_bytes(&initial, "odf", "odf").expect("odf -> odf");
    let zip = ZipArchive::new(&round).expect("valid zip");
    assert_eq!(
        zip.read_string("mimetype").unwrap().trim(),
        "application/vnd.oasis.opendocument.formula"
    );
}
