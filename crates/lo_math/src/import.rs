//! Math importers: read MathML, plain text, and ODF formula packages.

use lo_core::{parse_xml_document, FormulaDocument, FormulaNode, LoError, Result, XmlNode};
use lo_zip::ZipArchive;

use crate::flatten;

pub fn load_bytes(title: impl Into<String>, bytes: &[u8], format: &str) -> Result<FormulaDocument> {
    let title = title.into();
    let root = match format.to_ascii_lowercase().as_str() {
        "txt" | "text" | "latex" => crate::parse_latex(&bytes_to_utf8(bytes)?)?,
        "mathml" | "mml" | "xml" => parse_mathml_bytes(bytes)?,
        "odf" | "odfmath" | "odf-formula" => parse_odf_formula(bytes)?,
        other => return Err(LoError::Unsupported(format!("math import format {other}"))),
    };
    Ok(FormulaDocument::new(title, root))
}

/// Convenience: pull a single source string out of any supported format.
pub fn load_source(bytes: &[u8], format: &str) -> Result<String> {
    let doc = load_bytes("formula", bytes, format)?;
    Ok(flatten(&doc.root))
}

fn bytes_to_utf8(bytes: &[u8]) -> Result<String> {
    String::from_utf8(bytes.to_vec())
        .map_err(|err| LoError::Parse(format!("invalid utf-8 input: {err}")))
}

fn parse_mathml_bytes(bytes: &[u8]) -> Result<FormulaNode> {
    let text = bytes_to_utf8(bytes)?;
    let root = parse_xml_document(&text)?;
    parse_mathml_root(&root)
}

fn parse_odf_formula(bytes: &[u8]) -> Result<FormulaNode> {
    let zip = ZipArchive::new(bytes)?;
    let content = parse_xml_document(&zip.read_string("content.xml")?)?;
    // ODF math content.xml may either be `<office:document><office:body><office:math>`
    // (lo_odf style) or wrap a MathML document directly.
    let math_node = if content.local_name() == "math" {
        content
    } else if let Some(body) = content.child("body") {
        if let Some(math) = body.child("math") {
            math.clone()
        } else {
            return Err(LoError::Parse(
                "content.xml missing office:math/math".to_string(),
            ));
        }
    } else {
        return Err(LoError::Parse("unexpected math content.xml".to_string()));
    };
    let inner = math_node
        .children
        .iter()
        .find(|child| !matches!(child.local_name(), "annotation" | "annotation-xml"))
        .cloned()
        .unwrap_or(math_node);
    parse_mathml_root(&inner)
}

fn parse_mathml_root(root: &XmlNode) -> Result<FormulaNode> {
    match root.local_name() {
        "math" => parse_mathml_sequence(root),
        _ => parse_mathml_node(root),
    }
}

fn parse_mathml_sequence(node: &XmlNode) -> Result<FormulaNode> {
    let mut items = Vec::new();
    for child in &node.children {
        if matches!(child.local_name(), "annotation" | "annotation-xml") {
            continue;
        }
        items.push(parse_mathml_node(child)?);
    }
    if items.is_empty() {
        let text = node.text_content();
        if text.trim().is_empty() {
            Ok(FormulaNode::Group(Vec::new()))
        } else {
            Ok(FormulaNode::Identifier(text.trim().to_string()))
        }
    } else if items.len() == 1 {
        Ok(items.remove(0))
    } else {
        Ok(FormulaNode::Group(items))
    }
}

fn parse_mathml_node(node: &XmlNode) -> Result<FormulaNode> {
    Ok(match node.local_name() {
        "math" | "mrow" | "semantics" => parse_mathml_sequence(node)?,
        "mi" => FormulaNode::Identifier(node.text_content().trim().to_string()),
        "mn" => FormulaNode::Number(node.text_content().trim().to_string()),
        "mo" => FormulaNode::Symbol(node.text_content().trim().to_string()),
        "mtext" => FormulaNode::Identifier(node.text_content().trim().to_string()),
        "mfrac" => {
            let a = node
                .children
                .first()
                .ok_or_else(|| LoError::Parse("mfrac missing numerator".to_string()))?;
            let b = node
                .children
                .get(1)
                .ok_or_else(|| LoError::Parse("mfrac missing denominator".to_string()))?;
            FormulaNode::Fraction {
                numerator: Box::new(parse_mathml_node(a)?),
                denominator: Box::new(parse_mathml_node(b)?),
            }
        }
        "msqrt" => {
            // No dedicated Sqrt variant in current model; render as group `sqrt(<inner>)`.
            let inner = parse_mathml_sequence(node)?;
            FormulaNode::Group(vec![
                FormulaNode::Identifier("sqrt".to_string()),
                FormulaNode::Group(vec![inner]),
            ])
        }
        "msup" => {
            let base = node
                .children
                .first()
                .ok_or_else(|| LoError::Parse("msup missing base".to_string()))?;
            let exp = node
                .children
                .get(1)
                .ok_or_else(|| LoError::Parse("msup missing exponent".to_string()))?;
            FormulaNode::Superscript {
                base: Box::new(parse_mathml_node(base)?),
                exponent: Box::new(parse_mathml_node(exp)?),
            }
        }
        "msub" => {
            let base = node
                .children
                .first()
                .ok_or_else(|| LoError::Parse("msub missing base".to_string()))?;
            let sub = node
                .children
                .get(1)
                .ok_or_else(|| LoError::Parse("msub missing subscript".to_string()))?;
            FormulaNode::Subscript {
                base: Box::new(parse_mathml_node(base)?),
                subscript: Box::new(parse_mathml_node(sub)?),
            }
        }
        "msubsup" => {
            let base = node
                .children
                .first()
                .ok_or_else(|| LoError::Parse("msubsup missing base".to_string()))?;
            let sub = node
                .children
                .get(1)
                .ok_or_else(|| LoError::Parse("msubsup missing subscript".to_string()))?;
            let sup = node
                .children
                .get(2)
                .ok_or_else(|| LoError::Parse("msubsup missing superscript".to_string()))?;
            FormulaNode::Subscript {
                base: Box::new(FormulaNode::Superscript {
                    base: Box::new(parse_mathml_node(base)?),
                    exponent: Box::new(parse_mathml_node(sup)?),
                }),
                subscript: Box::new(parse_mathml_node(sub)?),
            }
        }
        "mfenced" => FormulaNode::Group(vec![parse_mathml_sequence(node)?]),
        "none" => FormulaNode::Group(Vec::new()),
        _ => {
            if node.children.is_empty() {
                let text = node.text_content();
                if text.chars().all(|ch| ch.is_ascii_digit() || ch == '.') {
                    FormulaNode::Number(text)
                } else {
                    FormulaNode::Identifier(text)
                }
            } else {
                parse_mathml_sequence(node)?
            }
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{from_latex, to_mathml_string};

    #[test]
    fn mathml_import_round_trip() {
        let doc = from_latex("demo", "\\frac{x^2}{y}").expect("latex");
        let mathml = to_mathml_string(&doc.root);
        let source = load_source(mathml.as_bytes(), "mathml").expect("import mathml");
        assert!(source.contains('/') || source.contains("frac"));
    }

    #[test]
    fn odf_import_round_trip() {
        let doc = from_latex("demo", "x^2").expect("latex");
        let tmp = std::env::temp_dir().join("lo_math_import_test.odf");
        lo_odf::save_formula_document(&tmp, &doc).expect("save odf");
        let bytes = std::fs::read(&tmp).expect("read");
        let _ = std::fs::remove_file(&tmp);
        let source = load_source(&bytes, "odf").expect("import odf");
        assert!(source.contains('^') || source.contains('x'));
    }
}
