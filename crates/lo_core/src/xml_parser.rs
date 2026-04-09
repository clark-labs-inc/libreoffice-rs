//! Minimal pure-Rust XML parser used by importers (DOCX, XLSX, ODF, …).
//!
//! It is intentionally lenient: namespaces are kept on element/attribute
//! names but accessors expose `local_name()` for convenience.

use std::collections::BTreeMap;

use crate::{LoError, Result};

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XmlItem {
    Text(String),
    Node(XmlNode),
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct XmlNode {
    pub name: String,
    pub attributes: BTreeMap<String, String>,
    pub children: Vec<XmlNode>,
    pub items: Vec<XmlItem>,
    pub text: String,
}

impl XmlNode {
    pub fn local_name(&self) -> &str {
        local_name(&self.name)
    }

    pub fn attr(&self, name: &str) -> Option<&str> {
        self.attributes.get(name).map(String::as_str).or_else(|| {
            self.attributes
                .iter()
                .find(|(key, _)| key.as_str() == name || local_name(key.as_str()) == name)
                .map(|(_, value)| value.as_str())
        })
    }

    pub fn child(&self, name: &str) -> Option<&XmlNode> {
        self.children
            .iter()
            .find(|child| child.local_name() == name || child.name == name)
    }

    pub fn children_named<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a XmlNode> + 'a {
        self.children
            .iter()
            .filter(move |child| child.local_name() == name || child.name == name)
    }

    pub fn descendants_named<'a>(&'a self, name: &'a str, out: &mut Vec<&'a XmlNode>) {
        for child in &self.children {
            if child.local_name() == name || child.name == name {
                out.push(child);
            }
            child.descendants_named(name, out);
        }
    }

    pub fn text_content(&self) -> String {
        let mut out = String::new();
        collect_text(self, &mut out);
        out
    }
}

fn collect_text(node: &XmlNode, out: &mut String) {
    if !node.text.is_empty() {
        out.push_str(&node.text);
    }
    for child in &node.children {
        collect_text(child, out);
    }
}

pub fn local_name(name: &str) -> &str {
    name.rsplit_once(':')
        .map(|(_, local)| local)
        .unwrap_or(name)
}

/// Serialize an `XmlNode` tree back into a self-closed/open XML string.
/// Includes the XML declaration. Useful when an importer parses, mutates,
/// and re-emits the document (e.g. recalculating an XLSX in place).
pub fn serialize_xml_document(root: &XmlNode) -> String {
    let mut out = String::new();
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    serialize_xml_node(root, &mut out);
    out
}

pub fn serialize_xml_node(node: &XmlNode, out: &mut String) {
    out.push('<');
    out.push_str(&node.name);
    for (key, value) in &node.attributes {
        out.push(' ');
        out.push_str(key);
        out.push_str("=\"");
        out.push_str(&xml_attr_escape(value));
        out.push('"');
    }
    if node.items.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for item in &node.items {
        match item {
            XmlItem::Text(text) => out.push_str(&xml_text_escape(text)),
            XmlItem::Node(child) => serialize_xml_node(child, out),
        }
    }
    out.push_str("</");
    out.push_str(&node.name);
    out.push('>');
}

fn xml_text_escape(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
    out
}

fn xml_attr_escape(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

pub fn parse_xml_document(xml: &str) -> Result<XmlNode> {
    let bytes = xml.as_bytes();
    let mut stack: Vec<XmlNode> = Vec::new();
    let mut root: Option<XmlNode> = None;
    let mut index = 0usize;

    while index < bytes.len() {
        if bytes[index] == b'<' {
            if bytes[index..].starts_with(b"<!--") {
                let end = find_bytes(bytes, index + 4, b"-->")?;
                index = end + 3;
                continue;
            }
            if bytes[index..].starts_with(b"<![CDATA[") {
                let end = find_bytes(bytes, index + 9, b"]]>")?;
                let text = String::from_utf8(bytes[index + 9..end].to_vec())
                    .map_err(|err| LoError::Parse(format!("invalid cdata utf-8: {err}")))?;
                if let Some(current) = stack.last_mut() {
                    current.text.push_str(&text);
                    current.items.push(XmlItem::Text(text));
                }
                index = end + 3;
                continue;
            }
            if bytes[index..].starts_with(b"<?") {
                let end = find_bytes(bytes, index + 2, b"?>")?;
                index = end + 2;
                continue;
            }
            if bytes[index..].starts_with(b"<!") {
                let end = find_byte(bytes, index + 2, b'>')?;
                index = end + 1;
                continue;
            }
            if bytes[index..].starts_with(b"</") {
                let end = find_byte(bytes, index + 2, b'>')?;
                let name = String::from_utf8(bytes[index + 2..end].to_vec())
                    .map_err(|err| LoError::Parse(format!("invalid closing tag name: {err}")))?;
                let node = stack.pop().ok_or_else(|| {
                    LoError::Parse("xml closing tag without opening tag".to_string())
                })?;
                if local_name(name.trim()) != node.local_name() {
                    return Err(LoError::Parse(format!(
                        "xml closing tag mismatch: expected {}, found {}",
                        node.name,
                        name.trim()
                    )));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node.clone());
                    parent.items.push(XmlItem::Node(node));
                } else if root.is_none() {
                    root = Some(node);
                } else {
                    return Err(LoError::Parse("multiple xml roots".to_string()));
                }
                index = end + 1;
                continue;
            }

            let end = find_tag_end(bytes, index + 1)?;
            let raw = String::from_utf8(bytes[index + 1..end].to_vec())
                .map_err(|err| LoError::Parse(format!("invalid tag utf-8: {err}")))?;
            let self_closing = raw.trim_end().ends_with('/');
            let raw = if self_closing {
                raw.trim_end().trim_end_matches('/').trim_end().to_string()
            } else {
                raw
            };
            let (name, attributes) = parse_start_tag(&raw)?;
            let node = XmlNode {
                name,
                attributes,
                children: Vec::new(),
                items: Vec::new(),
                text: String::new(),
            };
            if self_closing {
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node.clone());
                    parent.items.push(XmlItem::Node(node));
                } else if root.is_none() {
                    root = Some(node);
                } else {
                    return Err(LoError::Parse("multiple xml roots".to_string()));
                }
            } else {
                stack.push(node);
            }
            index = end + 1;
        } else {
            let next = find_byte_optional(bytes, index, b'<').unwrap_or(bytes.len());
            let raw_text = String::from_utf8(bytes[index..next].to_vec())
                .map_err(|err| LoError::Parse(format!("invalid text utf-8: {err}")))?;
            let decoded = decode_entities(&raw_text);
            if let Some(current) = stack.last_mut() {
                current.text.push_str(&decoded);
                if !decoded.is_empty() {
                    current.items.push(XmlItem::Text(decoded));
                }
            }
            index = next;
        }
    }

    while let Some(node) = stack.pop() {
        if let Some(parent) = stack.last_mut() {
            parent.children.push(node.clone());
            parent.items.push(XmlItem::Node(node));
        } else if root.is_none() {
            root = Some(node);
        } else {
            return Err(LoError::Parse("multiple xml roots".to_string()));
        }
    }

    root.ok_or_else(|| LoError::Parse("empty xml document".to_string()))
}

fn find_bytes(haystack: &[u8], start: usize, needle: &[u8]) -> Result<usize> {
    haystack[start..]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|offset| start + offset)
        .ok_or_else(|| LoError::Parse("unterminated xml construct".to_string()))
}

fn find_byte(bytes: &[u8], start: usize, byte: u8) -> Result<usize> {
    find_byte_optional(bytes, start, byte)
        .ok_or_else(|| LoError::Parse("unterminated xml tag".to_string()))
}

fn find_byte_optional(bytes: &[u8], start: usize, byte: u8) -> Option<usize> {
    bytes[start..]
        .iter()
        .position(|&value| value == byte)
        .map(|offset| start + offset)
}

fn find_tag_end(bytes: &[u8], start: usize) -> Result<usize> {
    let mut quote: Option<u8> = None;
    for index in start..bytes.len() {
        let byte = bytes[index];
        match quote {
            Some(current) if byte == current => quote = None,
            Some(_) => {}
            None if byte == b'\'' || byte == b'"' => quote = Some(byte),
            None if byte == b'>' => return Ok(index),
            None => {}
        }
    }
    Err(LoError::Parse("unterminated xml start tag".to_string()))
}

fn parse_start_tag(raw: &str) -> Result<(String, BTreeMap<String, String>)> {
    let mut chars = raw.chars().peekable();
    let mut name = String::new();
    while let Some(&ch) = chars.peek() {
        if ch.is_whitespace() {
            break;
        }
        name.push(ch);
        chars.next();
    }
    if name.is_empty() {
        return Err(LoError::Parse("empty xml tag name".to_string()));
    }
    while matches!(chars.peek(), Some(ch) if ch.is_whitespace()) {
        chars.next();
    }
    let mut attrs = BTreeMap::new();
    while chars.peek().is_some() {
        let mut key = String::new();
        while let Some(&ch) = chars.peek() {
            if ch.is_whitespace() || ch == '=' {
                break;
            }
            key.push(ch);
            chars.next();
        }
        while matches!(chars.peek(), Some(ch) if ch.is_whitespace()) {
            chars.next();
        }
        if chars.next() != Some('=') {
            return Err(LoError::Parse(format!("malformed xml attribute {key}")));
        }
        while matches!(chars.peek(), Some(ch) if ch.is_whitespace()) {
            chars.next();
        }
        let quote = chars
            .next()
            .ok_or_else(|| LoError::Parse("unexpected end of xml attribute".to_string()))?;
        if quote != '\'' && quote != '"' {
            return Err(LoError::Parse("xml attribute must be quoted".to_string()));
        }
        let mut value = String::new();
        for ch in chars.by_ref() {
            if ch == quote {
                break;
            }
            value.push(ch);
        }
        attrs.insert(key, decode_entities(&value));
        while matches!(chars.peek(), Some(ch) if ch.is_whitespace()) {
            chars.next();
        }
    }
    Ok((name, attrs))
}

pub fn decode_entities(text: &str) -> String {
    // Walk the input by char (not by byte) so multi-byte UTF-8 sequences
    // such as `æ` (U+00E6 → 0xC3 0xA6) round-trip correctly. The previous
    // implementation cast each byte to a `char`, which mojibake-d every
    // non-ASCII character into Latin-1.
    let mut out = String::with_capacity(text.len());
    let mut iter = text.char_indices().peekable();
    while let Some((idx, ch)) = iter.next() {
        if ch != '&' {
            out.push(ch);
            continue;
        }
        // Find the matching ';' (still within the original `text`).
        if let Some(rel_end) = text[idx + 1..].find(';') {
            let end = idx + 1 + rel_end;
            let entity = &text[idx + 1..end];
            let mut consumed = false;
            match entity {
                "amp" => {
                    out.push('&');
                    consumed = true;
                }
                "lt" => {
                    out.push('<');
                    consumed = true;
                }
                "gt" => {
                    out.push('>');
                    consumed = true;
                }
                "quot" => {
                    out.push('"');
                    consumed = true;
                }
                "apos" => {
                    out.push('\'');
                    consumed = true;
                }
                _ if entity.starts_with("#x") || entity.starts_with("#X") => {
                    if let Ok(value) = u32::from_str_radix(&entity[2..], 16) {
                        if let Some(ch) = char::from_u32(value) {
                            out.push(ch);
                            consumed = true;
                        }
                    }
                }
                _ if entity.starts_with('#') => {
                    if let Ok(value) = entity[1..].parse::<u32>() {
                        if let Some(ch) = char::from_u32(value) {
                            out.push(ch);
                            consumed = true;
                        }
                    }
                }
                _ => {
                    out.push('&');
                    out.push_str(entity);
                    out.push(';');
                    consumed = true;
                }
            }
            if consumed {
                // Advance the char iterator past the entity body + closing ';'.
                while let Some(&(next_idx, _)) = iter.peek() {
                    if next_idx > end {
                        break;
                    }
                    iter.next();
                }
                continue;
            }
        }
        out.push('&');
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parser_handles_simple_tree() {
        let root = parse_xml_document("<root><a x=\"1\">hi</a><b/></root>").unwrap();
        assert_eq!(root.local_name(), "root");
        assert_eq!(root.child("a").unwrap().text_content(), "hi");
        assert_eq!(root.child("a").unwrap().attr("x"), Some("1"));
    }

    #[test]
    fn decode_entities_handles_named_and_numeric() {
        assert_eq!(decode_entities("&amp;&lt;&#65;&#x42;"), "&<AB");
    }
}
