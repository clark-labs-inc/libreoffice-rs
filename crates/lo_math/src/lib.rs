use lo_core::{
    escape_text, geometry::Point, svg_footer, svg_header, svg_rect, svg_text, units::Length,
    write_text_pdf, FormulaDocument, FormulaNode, LoError, Rect, Result, Size,
};

pub mod import;
pub use import::{load_bytes, load_source};

pub fn from_latex(title: impl Into<String>, input: &str) -> Result<FormulaDocument> {
    Ok(FormulaDocument::new(title, parse_latex(input)?))
}

pub fn parse_latex(input: &str) -> Result<FormulaNode> {
    let chars: Vec<char> = input.chars().collect();
    let mut parser = Parser { chars, index: 0 };
    let node = parser.parse_sequence(None)?;
    parser.skip_ws();
    if parser.index != parser.chars.len() {
        return Err(LoError::Parse(
            "unexpected trailing characters in formula".to_string(),
        ));
    }
    Ok(node)
}

/// Flatten a formula node back into a TeX-ish single-line string. Used for
/// PDF and SVG previews where we have no real math layout engine.
pub fn flatten(node: &FormulaNode) -> String {
    match node {
        FormulaNode::Number(value)
        | FormulaNode::Identifier(value)
        | FormulaNode::Symbol(value) => value.clone(),
        FormulaNode::Operator { op, lhs, rhs } => {
            format!("{} {} {}", flatten(lhs), op, flatten(rhs))
        }
        FormulaNode::Fraction {
            numerator,
            denominator,
        } => {
            format!("({})/({})", flatten(numerator), flatten(denominator))
        }
        FormulaNode::Superscript { base, exponent } => {
            format!("{}^{}", flatten(base), flatten(exponent))
        }
        FormulaNode::Subscript { base, subscript } => {
            format!("{}_{}", flatten(base), flatten(subscript))
        }
        FormulaNode::Group(nodes) => {
            let mut parts = Vec::with_capacity(nodes.len());
            for n in nodes {
                parts.push(flatten(n));
            }
            parts.join(" ")
        }
    }
}

/// Render a formula as a single-page SVG. The formula is drawn as one line of
/// text inside a bordered box. Width/height are in points.
pub fn render_svg(node: &FormulaNode, size: Size) -> String {
    let mut svg = String::new();
    svg.push_str(&svg_header(size.width, size.height));
    svg.push_str(&svg_rect(
        Rect {
            origin: Point::new(Length::pt(0.5), Length::pt(0.5)),
            size: Size::new(
                Length::pt(size.width.as_pt() - 1.0),
                Length::pt(size.height.as_pt() - 1.0),
            ),
        },
        "#888888",
        Some("#ffffff"),
    ));
    svg.push_str(&svg_text(
        Length::pt(20.0),
        Length::pt(size.height.as_pt() / 2.0 + 10.0),
        &flatten(node),
        24,
        "#000000",
        "normal",
    ));
    svg.push_str(svg_footer());
    svg
}

/// Render a formula as a single-page text PDF.
pub fn to_pdf(node: &FormulaNode) -> Vec<u8> {
    write_text_pdf(&[flatten(node)], Length::pt(595.0), Length::pt(200.0))
}

/// Dispatch a formula to bytes for the requested format.
///
/// Supported (case-insensitive): `mathml`/`mml`, `svg`, `pdf`.
///
/// ODF formula files (`.odf`) are produced by `lo_odf::save_formula_document`,
/// which lives in the `lo_odf` crate so that all ODF packaging stays in one
/// place; that path is the one validated against real LibreOffice.
pub fn save_as(document: &FormulaDocument, format: &str) -> Result<Vec<u8>> {
    match format.to_ascii_lowercase().as_str() {
        "mathml" | "mml" => Ok(to_mathml_string(&document.root).into_bytes()),
        "svg" => {
            let size = Size::new(Length::pt(800.0), Length::pt(160.0));
            Ok(render_svg(&document.root, size).into_bytes())
        }
        "pdf" => Ok(to_pdf(&document.root)),
        other => Err(LoError::Unsupported(format!(
            "math format not supported: {other}"
        ))),
    }
}

pub fn to_mathml_string(node: &FormulaNode) -> String {
    let mut out = String::from("<math:math>");
    push_mathml(node, &mut out);
    out.push_str("</math:math>");
    out
}

fn push_mathml(node: &FormulaNode, out: &mut String) {
    match node {
        FormulaNode::Number(value) => {
            out.push_str("<math:mn>");
            out.push_str(&escape_text(value));
            out.push_str("</math:mn>");
        }
        FormulaNode::Identifier(value) => {
            out.push_str("<math:mi>");
            out.push_str(&escape_text(value));
            out.push_str("</math:mi>");
        }
        FormulaNode::Symbol(value) => {
            out.push_str("<math:mo>");
            out.push_str(&escape_text(value));
            out.push_str("</math:mo>");
        }
        FormulaNode::Operator { op, lhs, rhs } => {
            out.push_str("<math:mrow>");
            push_mathml(lhs, out);
            out.push_str("<math:mo>");
            out.push_str(&escape_text(op));
            out.push_str("</math:mo>");
            push_mathml(rhs, out);
            out.push_str("</math:mrow>");
        }
        FormulaNode::Fraction {
            numerator,
            denominator,
        } => {
            out.push_str("<math:mfrac>");
            push_mathml(numerator, out);
            push_mathml(denominator, out);
            out.push_str("</math:mfrac>");
        }
        FormulaNode::Superscript { base, exponent } => {
            out.push_str("<math:msup>");
            push_mathml(base, out);
            push_mathml(exponent, out);
            out.push_str("</math:msup>");
        }
        FormulaNode::Subscript { base, subscript } => {
            out.push_str("<math:msub>");
            push_mathml(base, out);
            push_mathml(subscript, out);
            out.push_str("</math:msub>");
        }
        FormulaNode::Group(nodes) => {
            out.push_str("<math:mrow>");
            for node in nodes {
                push_mathml(node, out);
            }
            out.push_str("</math:mrow>");
        }
    }
}

struct Parser {
    chars: Vec<char>,
    index: usize,
}

impl Parser {
    fn skip_ws(&mut self) {
        while self.index < self.chars.len() && self.chars[self.index].is_whitespace() {
            self.index += 1;
        }
    }

    fn parse_sequence(&mut self, terminator: Option<char>) -> Result<FormulaNode> {
        let mut nodes = Vec::new();
        let mut found_terminator = terminator.is_none();
        loop {
            self.skip_ws();
            if self.index >= self.chars.len() {
                break;
            }
            if let Some(end) = terminator {
                if self.chars[self.index] == end {
                    self.index += 1;
                    found_terminator = true;
                    break;
                }
            }
            let mut node = self.parse_atom()?;
            loop {
                self.skip_ws();
                match self.peek() {
                    Some('^') => {
                        self.index += 1;
                        let exponent = self.parse_script_atom()?;
                        node = FormulaNode::Superscript {
                            base: Box::new(node),
                            exponent: Box::new(exponent),
                        };
                    }
                    Some('_') => {
                        self.index += 1;
                        let subscript = self.parse_script_atom()?;
                        node = FormulaNode::Subscript {
                            base: Box::new(node),
                            subscript: Box::new(subscript),
                        };
                    }
                    _ => break,
                }
            }
            nodes.push(node);
        }

        if let Some(term) = terminator {
            if !found_terminator {
                return Err(LoError::Parse(format!(
                    "unterminated group ending with {term}"
                )));
            }
        }

        if nodes.is_empty() {
            Ok(FormulaNode::Group(Vec::new()))
        } else if nodes.len() == 1 {
            Ok(nodes.remove(0))
        } else {
            Ok(FormulaNode::Group(nodes))
        }
    }

    fn parse_script_atom(&mut self) -> Result<FormulaNode> {
        self.skip_ws();
        if self.peek() == Some('{') {
            self.index += 1;
            self.parse_sequence(Some('}'))
        } else {
            self.parse_atom()
        }
    }

    fn parse_atom(&mut self) -> Result<FormulaNode> {
        self.skip_ws();
        let ch = self
            .peek()
            .ok_or_else(|| LoError::Parse("unexpected end of formula".to_string()))?;

        if ch == '{' {
            self.index += 1;
            return self.parse_sequence(Some('}'));
        }

        if ch == '\\' {
            return self.parse_command();
        }

        if ch.is_ascii_digit() {
            return Ok(FormulaNode::Number(
                self.consume_while(|c| c.is_ascii_digit() || c == '.'),
            ));
        }

        if ch.is_ascii_alphabetic() {
            return Ok(FormulaNode::Identifier(
                self.consume_while(|c| c.is_ascii_alphanumeric()),
            ));
        }

        self.index += 1;
        Ok(FormulaNode::Symbol(ch.to_string()))
    }

    fn parse_command(&mut self) -> Result<FormulaNode> {
        self.expect('\\')?;
        let command = self.consume_while(|c| c.is_ascii_alphabetic());
        match command.as_str() {
            "frac" => {
                self.skip_ws();
                self.expect('{')?;
                let numerator = self.parse_sequence(Some('}'))?;
                self.skip_ws();
                self.expect('{')?;
                let denominator = self.parse_sequence(Some('}'))?;
                Ok(FormulaNode::Fraction {
                    numerator: Box::new(numerator),
                    denominator: Box::new(denominator),
                })
            }
            "cdot" | "times" => Ok(FormulaNode::Symbol("·".to_string())),
            "pm" => Ok(FormulaNode::Symbol("±".to_string())),
            "alpha" | "beta" | "gamma" | "delta" | "theta" | "lambda" | "mu" | "pi" | "sigma"
            | "omega" => Ok(FormulaNode::Identifier(command)),
            _ => Ok(FormulaNode::Identifier(command)),
        }
    }

    fn expect(&mut self, expected: char) -> Result<()> {
        match self.peek() {
            Some(ch) if ch == expected => {
                self.index += 1;
                Ok(())
            }
            _ => Err(LoError::Parse(format!("expected '{expected}'"))),
        }
    }

    fn consume_while(&mut self, predicate: impl Fn(char) -> bool) -> String {
        let start = self.index;
        while self.index < self.chars.len() && predicate(self.chars[self.index]) {
            self.index += 1;
        }
        self.chars[start..self.index].iter().collect()
    }

    fn peek(&self) -> Option<char> {
        self.chars.get(self.index).copied()
    }
}

#[cfg(test)]
mod tests {
    use super::{from_latex, parse_latex, to_mathml_string};

    #[test]
    fn parses_fraction() {
        let node = parse_latex(r"\frac{a+b}{c}").expect("parse latex");
        let dump = format!("{node:?}");
        assert!(dump.contains("Fraction"));
    }

    #[test]
    fn mathml_contains_tags() {
        let doc = from_latex("Quadratic", r"x^2 + y_1").expect("from latex");
        let mathml = to_mathml_string(&doc.root);
        assert!(mathml.contains("math:msup"));
        assert!(mathml.contains("math:msub"));
    }

    #[test]
    fn save_as_handles_all_formats() {
        let doc = from_latex("Q", r"\frac{a+b}{c}").expect("parse");
        for fmt in ["mathml", "svg", "pdf"] {
            let bytes = super::save_as(&doc, fmt).unwrap_or_else(|e| panic!("{fmt}: {e}"));
            assert!(!bytes.is_empty(), "{fmt} produced empty output");
        }
        assert!(super::save_as(&doc, "qq").is_err());
    }
}
