//! Shared pure-Rust PDF writer and reader.
//!
//! The original workspace only exposed tiny writer helpers for PDF export.
//! This module keeps those entry points but also adds a native parser,
//! stream decoder, page walker, and text extractor so the rest of the
//! workspace can treat PDF as an input format.

use std::cmp::Ordering;
use std::collections::BTreeMap;

use crate::units::Length;
use crate::{LoError, Result};

// ---------------------------------------------------------------------------
// Writer surface kept for backward compatibility.
// ---------------------------------------------------------------------------

fn pdf_escape(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('(', "\\(")
        .replace(')', "\\)")
}

/// Build a self-contained single-page text PDF.
///
/// `page_width` and `page_height` are interpreted as PDF user-space points.
pub fn write_text_pdf(lines: &[String], page_width: Length, page_height: Length) -> Vec<u8> {
    let width_pt = page_width.as_pt();
    let height_pt = page_height.as_pt();
    let mut content = String::new();
    content.push_str("BT\n/F1 12 Tf\n14 TL\n50 ");
    content.push_str(&format!("{:.2}", height_pt - 50.0));
    content.push_str(" Td\n");
    for (index, line) in lines.iter().enumerate() {
        if index > 0 {
            content.push_str("T*\n");
        }
        content.push('(');
        content.push_str(&pdf_escape(line));
        content.push_str(") Tj\n");
    }
    content.push_str("ET\n");
    let objects = vec![
        "<< /Type /Catalog /Pages 2 0 R >>".to_string(),
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>".to_string(),
        format!(
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width_pt:.2} {height_pt:.2}] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>"
        ),
        format!("<< /Length {} >>\nstream\n{}endstream", content.len(), content),
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>".to_string(),
    ];
    pdf_from_objects(&objects)
}

/// Serialize a list of already-rendered PDF object bodies into a complete
/// PDF byte stream with header, xref table, and trailer.
pub fn pdf_from_objects(objects: &[String]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"%PDF-1.4\n%");
    out.extend_from_slice(&[0xE2, 0xE3, 0xCF, 0xD3]);
    out.push(b'\n');
    let mut offsets = Vec::with_capacity(objects.len() + 1);
    offsets.push(0usize);
    for (index, object) in objects.iter().enumerate() {
        offsets.push(out.len());
        out.extend_from_slice(format!("{} 0 obj\n{}\nendobj\n", index + 1, object).as_bytes());
    }
    let xref_pos = out.len();
    out.extend_from_slice(format!("xref\n0 {}\n", objects.len() + 1).as_bytes());
    out.extend_from_slice(b"0000000000 65535 f \n");
    for offset in offsets.iter().skip(1) {
        out.extend_from_slice(format!("{:010} 00000 n \n", offset).as_bytes());
    }
    out.extend_from_slice(
        format!(
            "trailer\n<< /Size {} /Root 1 0 R >>\nstartxref\n{}\n%%EOF\n",
            objects.len() + 1,
            xref_pos
        )
        .as_bytes(),
    );
    out
}

/// Indirect object identifier.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct PdfObjectId {
    pub object: u32,
    pub generation: u16,
}

impl PdfObjectId {
    pub const fn new(object: u32, generation: u16) -> Self {
        Self { object, generation }
    }
}

/// Typed PDF object tree used by the reader and the higher-level writer.
#[derive(Clone, Debug, PartialEq)]
pub enum PdfValue {
    Null,
    Bool(bool),
    Number(f64),
    Name(String),
    String(Vec<u8>),
    Array(Vec<PdfValue>),
    Dict(BTreeMap<String, PdfValue>),
    Stream(PdfStream),
    Ref(PdfObjectId),
}

impl PdfValue {
    fn as_dict(&self) -> Option<&BTreeMap<String, PdfValue>> {
        match self {
            Self::Dict(dict) => Some(dict),
            Self::Stream(stream) => Some(&stream.dict),
            _ => None,
        }
    }

    fn as_name(&self) -> Option<&str> {
        match self {
            Self::Name(name) => Some(name.as_str()),
            _ => None,
        }
    }

    fn as_number(&self) -> Option<f64> {
        match self {
            Self::Number(value) => Some(*value),
            _ => None,
        }
    }

    fn as_ref(&self) -> Option<PdfObjectId> {
        match self {
            Self::Ref(id) => Some(*id),
            _ => None,
        }
    }
}

/// A decoded PDF stream object.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PdfStream {
    pub dict: BTreeMap<String, PdfValue>,
    pub data: Vec<u8>,
}

/// Builder for richer programmatic PDF output.
#[derive(Clone, Debug, Default)]
pub struct PdfBuilder {
    objects: Vec<Vec<u8>>,
    root: Option<PdfObjectId>,
    info: Option<PdfObjectId>,
}

impl PdfBuilder {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_object(&mut self, value: PdfValue) -> PdfObjectId {
        let id = PdfObjectId::new((self.objects.len() + 1) as u32, 0);
        self.objects.push(serialize_pdf_value(&value));
        id
    }

    pub fn set_root(&mut self, id: PdfObjectId) {
        self.root = Some(id);
    }

    pub fn set_info(&mut self, id: PdfObjectId) {
        self.info = Some(id);
    }

    pub fn finish(self) -> Result<Vec<u8>> {
        let root = self
            .root
            .ok_or_else(|| LoError::InvalidInput("pdf builder missing catalog root".to_string()))?;
        let mut out = Vec::new();
        out.extend_from_slice(b"%PDF-1.4\n%");
        out.extend_from_slice(&[0xE2, 0xE3, 0xCF, 0xD3]);
        out.push(b'\n');
        let mut offsets = Vec::with_capacity(self.objects.len() + 1);
        offsets.push(0usize);
        for (index, object) in self.objects.iter().enumerate() {
            offsets.push(out.len());
            out.extend_from_slice(format!("{} 0 obj\n", index + 1).as_bytes());
            out.extend_from_slice(object);
            if !out.ends_with(b"\n") {
                out.push(b'\n');
            }
            out.extend_from_slice(b"endobj\n");
        }
        let xref_pos = out.len();
        out.extend_from_slice(format!("xref\n0 {}\n", self.objects.len() + 1).as_bytes());
        out.extend_from_slice(b"0000000000 65535 f \n");
        for offset in offsets.iter().skip(1) {
            out.extend_from_slice(format!("{:010} 00000 n \n", offset).as_bytes());
        }
        let mut trailer = BTreeMap::new();
        trailer.insert("Size".to_string(), PdfValue::Number((self.objects.len() + 1) as f64));
        trailer.insert("Root".to_string(), PdfValue::Ref(root));
        if let Some(info) = self.info {
            trailer.insert("Info".to_string(), PdfValue::Ref(info));
        }
        out.extend_from_slice(b"trailer\n");
        out.extend_from_slice(&serialize_pdf_value(&PdfValue::Dict(trailer)));
        out.extend_from_slice(format!("\nstartxref\n{}\n%%EOF\n", xref_pos).as_bytes());
        Ok(out)
    }
}

fn serialize_pdf_value(value: &PdfValue) -> Vec<u8> {
    match value {
        PdfValue::Null => b"null".to_vec(),
        PdfValue::Bool(v) => {
            if *v {
                b"true".to_vec()
            } else {
                b"false".to_vec()
            }
        }
        PdfValue::Number(number) => {
            if number.fract() == 0.0 {
                format!("{:.0}", number).into_bytes()
            } else {
                format!("{number}").into_bytes()
            }
        }
        PdfValue::Name(name) => format!("/{}", escape_pdf_name(name)).into_bytes(),
        PdfValue::String(bytes) => serialize_pdf_string(bytes),
        PdfValue::Array(items) => {
            let mut out = Vec::new();
            out.push(b'[');
            for (index, item) in items.iter().enumerate() {
                if index > 0 {
                    out.push(b' ');
                }
                out.extend_from_slice(&serialize_pdf_value(item));
            }
            out.push(b']');
            out
        }
        PdfValue::Dict(dict) => serialize_pdf_dict(dict),
        PdfValue::Stream(stream) => {
            let mut dict = stream.dict.clone();
            dict.insert(
                "Length".to_string(),
                PdfValue::Number(stream.data.len() as f64),
            );
            let mut out = serialize_pdf_dict(&dict);
            out.extend_from_slice(b"\nstream\n");
            out.extend_from_slice(&stream.data);
            if !stream.data.ends_with(b"\n") {
                out.push(b'\n');
            }
            out.extend_from_slice(b"endstream");
            out
        }
        PdfValue::Ref(id) => format!("{} {} R", id.object, id.generation).into_bytes(),
    }
}

fn serialize_pdf_dict(dict: &BTreeMap<String, PdfValue>) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"<<");
    for (key, value) in dict {
        out.push(b' ');
        out.extend_from_slice(format!("/{} ", escape_pdf_name(key)).as_bytes());
        out.extend_from_slice(&serialize_pdf_value(value));
    }
    out.extend_from_slice(b" >>");
    out
}

fn escape_pdf_name(name: &str) -> String {
    let mut out = String::with_capacity(name.len());
    for byte in name.as_bytes() {
        match *byte {
            b'#' | b'/' | b'%' | b'(' | b')' | b'<' | b'>' | b'[' | b']' | b'{' | b'}'
            | b' ' | b'\t' | b'\r' | b'\n' => out.push_str(&format!("#{byte:02X}")),
            _ => out.push(*byte as char),
        }
    }
    out
}

fn serialize_pdf_string(bytes: &[u8]) -> Vec<u8> {
    let printable = bytes.iter().all(|byte| matches!(*byte, 0x20..=0x7E) && *byte != b'(' && *byte != b')' && *byte != b'\\');
    if printable {
        let mut out = Vec::new();
        out.push(b'(');
        for &byte in bytes {
            match byte {
                b'(' | b')' | b'\\' => {
                    out.push(b'\\');
                    out.push(byte);
                }
                _ => out.push(byte),
            }
        }
        out.push(b')');
        out
    } else {
        let mut out = String::from("<");
        for byte in bytes {
            out.push_str(&format!("{byte:02X}"));
        }
        out.push('>');
        out.into_bytes()
    }
}

// ---------------------------------------------------------------------------
// Reader / parser surface.
// ---------------------------------------------------------------------------

/// A positioned span extracted from a PDF page.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PdfTextSpan {
    pub x: f32,
    pub y: f32,
    pub end_x: f32,
    pub font_size: f32,
    pub text: String,
}

/// Extracted text content for one page.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PdfTextPage {
    pub width: f32,
    pub height: f32,
    pub spans: Vec<PdfTextSpan>,
}

impl PdfTextPage {
    /// Reconstruct a plain-text approximation of the page by grouping spans by
    /// baseline and sorting each line left-to-right.
    pub fn plain_text(&self) -> String {
        if self.spans.is_empty() {
            return String::new();
        }
        let mut spans = self.spans.clone();
        spans.sort_by(|a, b| {
            match b.y.partial_cmp(&a.y).unwrap_or(Ordering::Equal) {
                Ordering::Equal => a.x.partial_cmp(&b.x).unwrap_or(Ordering::Equal),
                other => other,
            }
        });
        let mut lines: Vec<Vec<PdfTextSpan>> = Vec::new();
        for span in spans {
            let tolerance = span.font_size.max(8.0) * 0.35;
            if let Some(line) = lines.last_mut() {
                let baseline = line.first().map(|first| first.y).unwrap_or(span.y);
                if (baseline - span.y).abs() <= tolerance {
                    line.push(span);
                    continue;
                }
            }
            lines.push(vec![span]);
        }
        let mut out = String::new();
        for (line_index, line) in lines.iter_mut().enumerate() {
            if line_index > 0 {
                out.push('\n');
            }
            line.sort_by(|a, b| a.x.partial_cmp(&b.x).unwrap_or(Ordering::Equal));
            let mut prev_end: Option<f32> = None;
            let mut prev_font: f32 = 10.0;
            for span in line {
                if let Some(end_x) = prev_end {
                    let gap = span.x - end_x;
                    if gap > prev_font.max(span.font_size) * 0.18 {
                        if !out.ends_with(' ') && !out.ends_with('\n') {
                            out.push(' ');
                        }
                    }
                }
                out.push_str(span.text.trim_matches('\0'));
                prev_end = Some(span.end_x.max(span.x));
                prev_font = span.font_size.max(1.0);
            }
        }
        out
    }
}

/// Parsed PDF document with extracted page text and metadata.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ParsedPdf {
    pub pages: Vec<PdfTextPage>,
    pub metadata: BTreeMap<String, String>,
}

impl ParsedPdf {
    /// Flatten the document into plain text with form-feed page separators.
    pub fn extract_text(&self) -> String {
        let mut out = String::new();
        for (index, page) in self.pages.iter().enumerate() {
            if index > 0 {
                out.push('\u{000C}');
                out.push('\n');
            }
            out.push_str(&page.plain_text());
        }
        out
    }

    /// Return page-by-page text chunks.
    pub fn page_texts(&self) -> Vec<String> {
        self.pages.iter().map(PdfTextPage::plain_text).collect()
    }
}

/// Parse a byte stream into a structured PDF text model.
pub fn parse_pdf(bytes: &[u8]) -> Result<ParsedPdf> {
    let file = PdfFile::parse(bytes)?;
    let metadata = file.extract_metadata();
    let page_nodes = file.collect_pages()?;
    let mut pages = Vec::with_capacity(page_nodes.len());
    for page in page_nodes {
        pages.push(file.extract_page_text(&page)?);
    }
    Ok(ParsedPdf { pages, metadata })
}

/// Extract the flattened plain text from a PDF byte stream.
pub fn extract_text_from_pdf(bytes: &[u8]) -> Result<String> {
    Ok(parse_pdf(bytes)?.extract_text())
}

/// Extract one text chunk per page from a PDF byte stream.
pub fn extract_pages_from_pdf(bytes: &[u8]) -> Result<Vec<String>> {
    Ok(parse_pdf(bytes)?.page_texts())
}

#[derive(Clone, Debug, Default)]
struct PdfFile {
    objects: BTreeMap<PdfObjectId, PdfValue>,
    trailer: BTreeMap<String, PdfValue>,
}

impl PdfFile {
    fn parse(bytes: &[u8]) -> Result<Self> {
        let mut objects = BTreeMap::new();
        let mut pos = 0usize;
        while pos < bytes.len() {
            match try_parse_indirect_object(bytes, pos, &objects)? {
                Some((id, value, end)) => {
                    objects.insert(id, value);
                    pos = end;
                }
                None => pos += 1,
            }
        }
        let trailer = parse_trailer(bytes).unwrap_or_default();
        let mut file = Self { objects, trailer };
        file.expand_object_streams()?;
        Ok(file)
    }

    fn resolve(&self, value: &PdfValue) -> Option<PdfValue> {
        let mut current = value.clone();
        for _ in 0..32 {
            match current {
                PdfValue::Ref(id) => {
                    current = self.objects.get(&id)?.clone();
                }
                _ => return Some(current),
            }
        }
        None
    }

    fn resolve_dict(&self, value: &PdfValue) -> Option<BTreeMap<String, PdfValue>> {
        match self.resolve(value)? {
            PdfValue::Dict(dict) => Some(dict),
            PdfValue::Stream(stream) => Some(stream.dict),
            _ => None,
        }
    }

    fn resolve_stream(&self, value: &PdfValue) -> Option<PdfStream> {
        match self.resolve(value)? {
            PdfValue::Stream(stream) => Some(stream),
            _ => None,
        }
    }

    fn resolve_number(&self, value: &PdfValue) -> Option<f64> {
        self.resolve(value)?.as_number()
    }

    fn extract_metadata(&self) -> BTreeMap<String, String> {
        let mut meta = BTreeMap::new();
        if let Some(info) = self.trailer.get("Info") {
            if let Some(dict) = self.resolve_dict(info) {
                for key in ["Title", "Author", "Subject", "Keywords", "Creator", "Producer"] {
                    if let Some(value) = dict.get(key) {
                        if let Some(text) = decode_pdf_text_object(&self.resolve(value).unwrap_or_else(|| value.clone())) {
                            if !text.trim().is_empty() {
                                meta.insert(key.to_string(), text);
                            }
                        }
                    }
                }
            }
        }
        meta
    }

    fn catalog_id(&self) -> Option<PdfObjectId> {
        if let Some(root) = self.trailer.get("Root").and_then(PdfValue::as_ref) {
            return Some(root);
        }
        self.objects.iter().find_map(|(id, value)| {
            let dict = value.as_dict()?;
            if dict.get("Type").and_then(PdfValue::as_name) == Some("Catalog") {
                Some(*id)
            } else {
                None
            }
        })
    }

    fn collect_pages(&self) -> Result<Vec<PageNode>> {
        if let Some(catalog_id) = self.catalog_id() {
            if let Some(catalog) = self.objects.get(&catalog_id).and_then(PdfValue::as_dict) {
                if let Some(pages_ref) = catalog.get("Pages").and_then(PdfValue::as_ref) {
                    let mut out = Vec::new();
                    self.walk_pages(pages_ref, &InheritedPageState::default(), &mut out)?;
                    if !out.is_empty() {
                        return Ok(out);
                    }
                }
            }
        }
        let mut fallback = Vec::new();
        for (id, value) in &self.objects {
            let Some(dict) = value.as_dict() else { continue };
            if dict.get("Type").and_then(PdfValue::as_name) == Some("Page") {
                let media_box = dict
                    .get("MediaBox")
                    .and_then(|value| parse_media_box(self, value))
                    .unwrap_or((0.0, 0.0, 612.0, 792.0));
                let resources = dict
                    .get("Resources")
                    .and_then(|value| self.resolve_dict(value))
                    .unwrap_or_default();
                fallback.push(PageNode {
                    id: *id,
                    dict: dict.clone(),
                    resources,
                    media_box,
                });
            }
        }
        fallback.sort_by_key(|page| page.id);
        Ok(fallback)
    }

    fn walk_pages(
        &self,
        node_id: PdfObjectId,
        inherited: &InheritedPageState,
        out: &mut Vec<PageNode>,
    ) -> Result<()> {
        let dict = self
            .objects
            .get(&node_id)
            .and_then(PdfValue::as_dict)
            .cloned()
            .ok_or_else(|| LoError::Parse(format!("missing page tree node {} {}", node_id.object, node_id.generation)))?;
        let ty = dict.get("Type").and_then(PdfValue::as_name).unwrap_or("");
        let mut next = inherited.clone();
        if let Some(resources) = dict.get("Resources").and_then(|value| self.resolve_dict(value)) {
            next.resources = Some(resources);
        }
        if let Some(media_box) = dict.get("MediaBox").and_then(|value| parse_media_box(self, value)) {
            next.media_box = Some(media_box);
        }
        if ty == "Pages" || dict.contains_key("Kids") {
            if let Some(PdfValue::Array(kids)) = dict.get("Kids") {
                for kid in kids {
                    if let Some(id) = kid.as_ref() {
                        self.walk_pages(id, &next, out)?;
                    }
                }
            }
            return Ok(());
        }
        let media_box = next.media_box.unwrap_or((0.0, 0.0, 612.0, 792.0));
        let resources = next.resources.clone().unwrap_or_default();
        out.push(PageNode {
            id: node_id,
            dict,
            resources,
            media_box,
        });
        Ok(())
    }

    fn extract_page_text(&self, page: &PageNode) -> Result<PdfTextPage> {
        let resources = self.build_resources(&page.resources)?;
        let mut spans = Vec::new();
        for stream in self.page_content_streams(&page.dict)? {
            let data = self.decode_stream(&stream)?;
            self.extract_content_stream(&data, &resources, Matrix::identity(), &mut spans)?;
        }
        spans.retain(|span| !span.text.trim().is_empty());
        let (_, _, x1, y1) = page.media_box;
        Ok(PdfTextPage {
            width: x1,
            height: y1,
            spans,
        })
    }

    fn page_content_streams(&self, page: &BTreeMap<String, PdfValue>) -> Result<Vec<PdfStream>> {
        let Some(contents) = page.get("Contents") else {
            return Ok(Vec::new());
        };
        let mut out = Vec::new();
        match self.resolve(contents) {
            Some(PdfValue::Stream(stream)) => out.push(stream),
            Some(PdfValue::Array(items)) => {
                for item in items {
                    if let Some(stream) = self.resolve_stream(&item) {
                        out.push(stream);
                    }
                }
            }
            Some(other) => {
                return Err(LoError::Parse(format!(
                    "page Contents is not a stream/array: {other:?}"
                )))
            }
            None => {}
        }
        Ok(out)
    }

    fn decode_stream(&self, stream: &PdfStream) -> Result<Vec<u8>> {
        let filters = collect_filter_names(stream.dict.get("Filter"));
        let mut out = stream.data.clone();
        for filter in filters {
            out = match filter.as_str() {
                "FlateDecode" | "Fl" => decode_flate_stream(&out)?,
                "ASCIIHexDecode" | "AHx" => decode_ascii_hex(&out)?,
                "ASCII85Decode" | "A85" => decode_ascii85(&out)?,
                "RunLengthDecode" | "RL" => decode_run_length(&out)?,
                "LZWDecode" | "LZW" => decode_lzw(&out)?,
                other => {
                    return Err(LoError::Unsupported(format!(
                        "pdf filter not supported: {other}"
                    )))
                }
            };
        }
        Ok(out)
    }

    fn expand_object_streams(&mut self) -> Result<()> {
        let object_streams: Vec<(PdfObjectId, PdfStream)> = self
            .objects
            .iter()
            .filter_map(|(id, value)| match value {
                PdfValue::Stream(stream)
                    if stream
                        .dict
                        .get("Type")
                        .and_then(PdfValue::as_name)
                        == Some("ObjStm") => Some((*id, stream.clone())),
                _ => None,
            })
            .collect();
        for (_id, stream) in object_streams {
            let decoded = self.decode_stream(&stream)?;
            let n = stream
                .dict
                .get("N")
                .and_then(PdfValue::as_number)
                .unwrap_or(0.0) as usize;
            let first = stream
                .dict
                .get("First")
                .and_then(PdfValue::as_number)
                .unwrap_or(0.0) as usize;
            if first > decoded.len() {
                continue;
            }
            let header = std::str::from_utf8(&decoded[..first]).unwrap_or("");
            let mut header_numbers = header
                .split_whitespace()
                .filter_map(|part| part.parse::<usize>().ok());
            let mut entries = Vec::new();
            for _ in 0..n {
                let Some(obj_num) = header_numbers.next() else { break };
                let Some(offset) = header_numbers.next() else { break };
                entries.push((obj_num as u32, offset));
            }
            for (obj_num, offset) in entries {
                let start = first + offset;
                if start >= decoded.len() {
                    continue;
                }
                let mut parser = Parser::new(&decoded[start..]);
                if let Ok(value) = parser.parse_value() {
                    self.objects
                        .entry(PdfObjectId::new(obj_num, 0))
                        .or_insert(value);
                }
            }
        }
        Ok(())
    }

    fn build_resources(&self, dict: &BTreeMap<String, PdfValue>) -> Result<PdfResources> {
        let mut resources = PdfResources::default();
        if let Some(fonts) = dict.get("Font").and_then(|value| self.resolve_dict(value)) {
            for (name, value) in fonts {
                let Some(font_value) = self.resolve(&value) else { continue };
                resources
                    .fonts
                    .insert(name.clone(), FontDecoder::from_pdf_value(self, &font_value)?);
            }
        }
        if let Some(xobjects) = dict.get("XObject").and_then(|value| self.resolve_dict(value)) {
            for (name, value) in xobjects {
                let Some(stream) = self.resolve_stream(&value) else { continue };
                let subtype = stream
                    .dict
                    .get("Subtype")
                    .and_then(PdfValue::as_name)
                    .unwrap_or("")
                    .to_string();
                if subtype != "Form" {
                    continue;
                }
                let matrix = stream
                    .dict
                    .get("Matrix")
                    .and_then(|value| parse_matrix(self, value))
                    .unwrap_or_else(Matrix::identity);
                let form_resources = stream
                    .dict
                    .get("Resources")
                    .and_then(|value| self.resolve_dict(value))
                    .unwrap_or_default();
                resources.xobjects.insert(
                    name.clone(),
                    PdfXObject {
                        data: self.decode_stream(&stream)?,
                        resources: form_resources,
                        matrix,
                    },
                );
            }
        }
        Ok(resources)
    }

    fn extract_content_stream(
        &self,
        data: &[u8],
        resources: &PdfResources,
        initial_ctm: Matrix,
        spans: &mut Vec<PdfTextSpan>,
    ) -> Result<()> {
        let mut parser = ContentParser::new(data);
        let mut operands: Vec<ContentToken> = Vec::new();
        let mut graphics_stack: Vec<Matrix> = vec![initial_ctm];
        let mut text = TextState::default();
        while let Some(token) = parser.next_token()? {
            match token {
                ContentToken::Operator(op) => {
                    match op.as_str() {
                        "q" => graphics_stack.push(*graphics_stack.last().unwrap_or(&Matrix::identity())),
                        "Q" => {
                            if graphics_stack.len() > 1 {
                                graphics_stack.pop();
                            }
                        }
                        "cm" => {
                            if let Some(matrix) = take_six_numbers(&mut operands) {
                                let current = *graphics_stack.last().unwrap_or(&Matrix::identity());
                                if let Some(last) = graphics_stack.last_mut() {
                                    *last = current.multiply(&Matrix::new(matrix[0], matrix[1], matrix[2], matrix[3], matrix[4], matrix[5]));
                                }
                            }
                        }
                        "BT" => text = TextState::default(),
                        "ET" => {}
                        "Tf" => {
                            if operands.len() >= 2 {
                                let size = operands.pop().and_then(|t| t.as_number()).unwrap_or(12.0);
                                let name = operands.pop().and_then(|t| t.into_name()).unwrap_or_default();
                                text.font = name;
                                text.font_size = size.max(1.0);
                            }
                        }
                        "Tm" => {
                            if let Some(values) = take_six_numbers(&mut operands) {
                                let matrix = Matrix::new(values[0], values[1], values[2], values[3], values[4], values[5]);
                                text.text_matrix = matrix;
                                text.line_matrix = matrix;
                            }
                        }
                        "Td" => {
                            if let Some([tx, ty]) = take_two_numbers(&mut operands) {
                                text.translate(tx, ty);
                            }
                        }
                        "TD" => {
                            if let Some([tx, ty]) = take_two_numbers(&mut operands) {
                                text.leading = -ty;
                                text.translate(tx, ty);
                            }
                        }
                        "T*" => text.translate(0.0, -text.leading),
                        "Tc" => {
                            if let Some(value) = operands.pop().and_then(|t| t.as_number()) {
                                text.char_spacing = value;
                            }
                        }
                        "Tw" => {
                            if let Some(value) = operands.pop().and_then(|t| t.as_number()) {
                                text.word_spacing = value;
                            }
                        }
                        "TL" => {
                            if let Some(value) = operands.pop().and_then(|t| t.as_number()) {
                                text.leading = value;
                            }
                        }
                        "Tz" => {
                            if let Some(value) = operands.pop().and_then(|t| t.as_number()) {
                                text.horizontal_scaling = value / 100.0;
                            }
                        }
                        "Ts" => {
                            if let Some(value) = operands.pop().and_then(|t| t.as_number()) {
                                text.rise = value;
                            }
                        }
                        "Tj" => {
                            if let Some(bytes) = operands.pop().and_then(|t| t.into_bytes()) {
                                show_pdf_text(resources, &graphics_stack, &mut text, &bytes, spans);
                            }
                        }
                        "TJ" => {
                            if let Some(items) = operands.pop().and_then(|t| t.into_array()) {
                                for item in items {
                                    match item {
                                        ContentToken::String(bytes) => {
                                            show_pdf_text(resources, &graphics_stack, &mut text, &bytes, spans);
                                        }
                                        ContentToken::Number(adjust) => {
                                            let shift = -(adjust / 1000.0) * text.font_size * text.horizontal_scaling;
                                            text.text_matrix = text.text_matrix.translate(shift, 0.0);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                        }
                        "'" => {
                            text.translate(0.0, -text.leading);
                            if let Some(bytes) = operands.pop().and_then(|t| t.into_bytes()) {
                                show_pdf_text(resources, &graphics_stack, &mut text, &bytes, spans);
                            }
                        }
                        "\"" => {
                            if operands.len() >= 3 {
                                let string = operands.pop().and_then(|t| t.into_bytes()).unwrap_or_default();
                                text.char_spacing = operands.pop().and_then(|t| t.as_number()).unwrap_or(text.char_spacing);
                                text.word_spacing = operands.pop().and_then(|t| t.as_number()).unwrap_or(text.word_spacing);
                                text.translate(0.0, -text.leading);
                                show_pdf_text(resources, &graphics_stack, &mut text, &string, spans);
                            }
                        }
                        "Do" => {
                            if let Some(name) = operands.pop().and_then(|t| t.into_name()) {
                                if let Some(xobj) = resources.xobjects.get(&name) {
                                    let nested = self.build_resources(&xobj.resources)?;
                                    let ctm = graphics_stack
                                        .last()
                                        .copied()
                                        .unwrap_or_else(Matrix::identity)
                                        .multiply(&xobj.matrix);
                                    self.extract_content_stream(&xobj.data, &nested, ctm, spans)?;
                                }
                            }
                        }
                        "BI" => parser.skip_inline_image(),
                        _ => {}
                    }
                    operands.clear();
                }
                other => operands.push(other),
            }
        }
        Ok(())
    }
}

#[derive(Clone, Debug)]
struct PageNode {
    id: PdfObjectId,
    dict: BTreeMap<String, PdfValue>,
    resources: BTreeMap<String, PdfValue>,
    media_box: (f32, f32, f32, f32),
}

#[derive(Clone, Debug, Default)]
struct InheritedPageState {
    resources: Option<BTreeMap<String, PdfValue>>,
    media_box: Option<(f32, f32, f32, f32)>,
}

fn parse_media_box(file: &PdfFile, value: &PdfValue) -> Option<(f32, f32, f32, f32)> {
    let PdfValue::Array(items) = file.resolve(value)? else {
        return None;
    };
    if items.len() != 4 {
        return None;
    }
    Some((
        file.resolve_number(&items[0])? as f32,
        file.resolve_number(&items[1])? as f32,
        file.resolve_number(&items[2])? as f32,
        file.resolve_number(&items[3])? as f32,
    ))
}

fn parse_matrix(file: &PdfFile, value: &PdfValue) -> Option<Matrix> {
    let PdfValue::Array(items) = file.resolve(value)? else {
        return None;
    };
    if items.len() != 6 {
        return None;
    }
    Some(Matrix::new(
        file.resolve_number(&items[0])? as f32,
        file.resolve_number(&items[1])? as f32,
        file.resolve_number(&items[2])? as f32,
        file.resolve_number(&items[3])? as f32,
        file.resolve_number(&items[4])? as f32,
        file.resolve_number(&items[5])? as f32,
    ))
}

fn decode_pdf_text_object(value: &PdfValue) -> Option<String> {
    match value {
        PdfValue::String(bytes) => decode_pdf_string_bytes(bytes),
        PdfValue::Name(name) => Some(name.clone()),
        PdfValue::Number(n) => Some(n.to_string()),
        _ => None,
    }
}

fn try_parse_indirect_object(
    bytes: &[u8],
    start: usize,
    objects: &BTreeMap<PdfObjectId, PdfValue>,
) -> Result<Option<(PdfObjectId, PdfValue, usize)>> {
    if start > 0 && !is_pdf_space(bytes[start - 1]) {
        return Ok(None);
    }
    let mut header = Parser::new_at(bytes, start);
    let Some(object) = header.parse_unsigned_number().ok() else {
        return Ok(None);
    };
    header.skip_ws_and_comments();
    let Some(generation) = header.parse_unsigned_number().ok() else {
        return Ok(None);
    };
    header.skip_ws_and_comments();
    if !header.consume_keyword(b"obj") {
        return Ok(None);
    }
    header.skip_ws_and_comments();
    let mut value = header.parse_value()?;
    header.skip_ws_and_comments();
    if matches!(value, PdfValue::Dict(_)) && header.consume_keyword(b"stream") {
        let data_start = skip_stream_eol(bytes, header.pos);
        let length = match &value {
            PdfValue::Dict(dict) => dict
                .get("Length")
                .and_then(|value| resolve_length_hint(value, objects)),
            _ => None,
        };
        let (data, after_stream) = read_stream_bytes(bytes, data_start, length)?;
        let dict = match value {
            PdfValue::Dict(dict) => dict,
            _ => unreachable!(),
        };
        value = PdfValue::Stream(PdfStream { dict, data });
        header.pos = after_stream;
        header.skip_ws_and_comments();
    }
    let endobj = find_token(bytes, header.pos, b"endobj")
        .ok_or_else(|| LoError::Parse("pdf indirect object missing endobj".to_string()))?;
    Ok(Some((
        PdfObjectId::new(object, generation as u16),
        value,
        endobj + "endobj".len(),
    )))
}

fn resolve_length_hint(
    value: &PdfValue,
    objects: &BTreeMap<PdfObjectId, PdfValue>,
) -> Option<usize> {
    match value {
        PdfValue::Number(number) => Some((*number).max(0.0) as usize),
        PdfValue::Ref(id) => match objects.get(id) {
            Some(PdfValue::Number(number)) => Some((*number).max(0.0) as usize),
            _ => None,
        },
        _ => None,
    }
}

fn skip_stream_eol(bytes: &[u8], mut pos: usize) -> usize {
    if bytes.get(pos) == Some(&b'\r') {
        pos += 1;
    }
    if bytes.get(pos) == Some(&b'\n') {
        pos += 1;
    }
    pos
}

fn read_stream_bytes(bytes: &[u8], data_start: usize, length: Option<usize>) -> Result<(Vec<u8>, usize)> {
    if let Some(length) = length {
        let end = data_start.saturating_add(length);
        if end > bytes.len() {
            return Err(LoError::Parse("pdf stream length is out of bounds".to_string()));
        }
        let mut cursor = end;
        if bytes.get(cursor) == Some(&b'\r') {
            cursor += 1;
        }
        if bytes.get(cursor) == Some(&b'\n') {
            cursor += 1;
        }
        let endstream = find_token(bytes, cursor, b"endstream")
            .ok_or_else(|| LoError::Parse("pdf stream missing endstream".to_string()))?;
        return Ok((bytes[data_start..end].to_vec(), endstream + "endstream".len()));
    }
    let endstream = find_token(bytes, data_start, b"endstream")
        .ok_or_else(|| LoError::Parse("pdf stream missing endstream".to_string()))?;
    let mut data_end = endstream;
    while data_end > data_start && matches!(bytes[data_end - 1], b'\r' | b'\n') {
        data_end -= 1;
    }
    Ok((bytes[data_start..data_end].to_vec(), endstream + "endstream".len()))
}

fn parse_trailer(bytes: &[u8]) -> Option<BTreeMap<String, PdfValue>> {
    let index = rfind_bytes(bytes, b"trailer")?;
    let mut parser = Parser::new_at(bytes, index + "trailer".len());
    parser.skip_ws_and_comments();
    let PdfValue::Dict(dict) = parser.parse_value().ok()? else {
        return None;
    };
    Some(dict)
}

fn find_token(bytes: &[u8], start: usize, token: &[u8]) -> Option<usize> {
    let mut index = start;
    while index + token.len() <= bytes.len() {
        if &bytes[index..index + token.len()] == token {
            let prev_ok = index == 0 || is_pdf_space(bytes[index - 1]) || is_pdf_delim(bytes[index - 1]);
            let next_ok = index + token.len() == bytes.len()
                || is_pdf_space(bytes[index + token.len()])
                || is_pdf_delim(bytes[index + token.len()]);
            if prev_ok && next_ok {
                return Some(index);
            }
        }
        index += 1;
    }
    None
}

fn rfind_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack.windows(needle.len()).rposition(|window| window == needle)
}

#[derive(Clone, Debug, Default)]
struct PdfResources {
    fonts: BTreeMap<String, FontDecoder>,
    xobjects: BTreeMap<String, PdfXObject>,
}

#[derive(Clone, Debug)]
struct PdfXObject {
    data: Vec<u8>,
    resources: BTreeMap<String, PdfValue>,
    matrix: Matrix,
}

#[derive(Clone, Debug)]
struct FontDecoder {
    cmap: Option<ToUnicodeMap>,
    encoding: SimpleEncoding,
}

impl Default for FontDecoder {
    fn default() -> Self {
        Self {
            cmap: None,
            encoding: SimpleEncoding::WinAnsi,
        }
    }
}

impl FontDecoder {
    fn from_pdf_value(file: &PdfFile, value: &PdfValue) -> Result<Self> {
        let Some(dict) = value.as_dict() else {
            return Ok(Self::default());
        };
        let encoding = parse_simple_encoding(file, dict);
        let cmap = if let Some(to_unicode) = dict.get("ToUnicode") {
            if let Some(stream) = file.resolve_stream(to_unicode) {
                Some(ToUnicodeMap::parse(&file.decode_stream(&stream)?)?)
            } else {
                None
            }
        } else {
            None
        };
        Ok(Self { cmap, encoding })
    }

    fn decode(&self, bytes: &[u8]) -> String {
        if let Some(cmap) = &self.cmap {
            let text = cmap.decode(bytes);
            if !text.is_empty() {
                return text;
            }
        }
        decode_simple_encoding(bytes, self.encoding)
    }
}

#[derive(Clone, Copy, Debug)]
enum SimpleEncoding {
    Ascii,
    WinAnsi,
}

fn parse_simple_encoding(file: &PdfFile, dict: &BTreeMap<String, PdfValue>) -> SimpleEncoding {
    if let Some(encoding) = dict.get("Encoding") {
        match file.resolve(encoding) {
            Some(PdfValue::Name(name)) => {
                let lower = name.to_ascii_lowercase();
                if lower.contains("winansi") || lower.contains("standard") {
                    SimpleEncoding::WinAnsi
                } else {
                    SimpleEncoding::Ascii
                }
            }
            Some(PdfValue::Dict(inner)) => inner
                .get("BaseEncoding")
                .and_then(PdfValue::as_name)
                .map(|name| {
                    let lower = name.to_ascii_lowercase();
                    if lower.contains("winansi") || lower.contains("standard") {
                        SimpleEncoding::WinAnsi
                    } else {
                        SimpleEncoding::Ascii
                    }
                })
                .unwrap_or(SimpleEncoding::Ascii),
            _ => SimpleEncoding::Ascii,
        }
    } else {
        SimpleEncoding::WinAnsi
    }
}

#[derive(Clone, Debug, Default)]
struct ToUnicodeMap {
    code_ranges: Vec<(usize, u32, u32)>,
    direct: BTreeMap<u32, String>,
}

impl ToUnicodeMap {
    fn parse(data: &[u8]) -> Result<Self> {
        let text = String::from_utf8_lossy(data);
        let mut out = Self::default();
        let lines: Vec<&str> = text.lines().collect();
        let mut index = 0usize;
        while index < lines.len() {
            let line = lines[index].trim();
            if let Some(count) = line.strip_suffix("begincodespacerange").and_then(|prefix| prefix.trim().parse::<usize>().ok()) {
                for offset in 1..=count {
                    if let Some((start, end)) = parse_two_hex_strings(lines.get(index + offset).copied().unwrap_or("")) {
                        out.code_ranges.push((start.len(), hex_bytes_to_u32(&start), hex_bytes_to_u32(&end)));
                    }
                }
                index += count + 1;
                continue;
            }
            if let Some(count) = line.strip_suffix("beginbfchar").and_then(|prefix| prefix.trim().parse::<usize>().ok()) {
                for offset in 1..=count {
                    if let Some((src, dst)) = parse_two_hex_strings(lines.get(index + offset).copied().unwrap_or("")) {
                        out.direct
                            .insert(hex_bytes_to_u32(&src), decode_utf16be_fallback(&dst));
                    }
                }
                index += count + 1;
                continue;
            }
            if let Some(count) = line.strip_suffix("beginbfrange").and_then(|prefix| prefix.trim().parse::<usize>().ok()) {
                for offset in 1..=count {
                    let entry = lines.get(index + offset).copied().unwrap_or("").trim();
                    if let Some((start, end, dst)) = parse_bfrange_entry(entry) {
                        let start_code = hex_bytes_to_u32(&start);
                        let end_code = hex_bytes_to_u32(&end);
                        match dst {
                            BfRangeDest::Sequential(base) => {
                                let mut current = base.clone();
                                for code in start_code..=end_code {
                                    out.direct.insert(code, decode_utf16be_fallback(&current));
                                    increment_utf16be_bytes(&mut current);
                                }
                            }
                            BfRangeDest::Explicit(items) => {
                                for (offset, item) in items.into_iter().enumerate() {
                                    out.direct.insert(start_code + offset as u32, decode_utf16be_fallback(&item));
                                }
                            }
                        }
                    }
                }
                index += count + 1;
                continue;
            }
            index += 1;
        }
        if out.code_ranges.is_empty() {
            out.code_ranges.push((1, 0x00, 0xFF));
            out.code_ranges.push((2, 0x0000, 0xFFFF));
        }
        out.code_ranges.sort_by_key(|entry| entry.0);
        Ok(out)
    }

    fn decode(&self, bytes: &[u8]) -> String {
        let mut out = String::new();
        let mut pos = 0usize;
        let max_len = self.code_ranges.iter().map(|entry| entry.0).max().unwrap_or(1);
        while pos < bytes.len() {
            let mut matched = false;
            for len in (1..=max_len).rev() {
                if pos + len > bytes.len() {
                    continue;
                }
                let code = hex_bytes_to_u32(&bytes[pos..pos + len]);
                if !self.code_ranges.iter().any(|(range_len, start, end)| *range_len == len && code >= *start && code <= *end) {
                    continue;
                }
                if let Some(mapped) = self.direct.get(&code) {
                    out.push_str(mapped);
                    pos += len;
                    matched = true;
                    break;
                }
            }
            if !matched {
                out.push_str(&decode_simple_encoding(&bytes[pos..pos + 1], SimpleEncoding::WinAnsi));
                pos += 1;
            }
        }
        out
    }
}

enum BfRangeDest {
    Sequential(Vec<u8>),
    Explicit(Vec<Vec<u8>>),
}

fn parse_bfrange_entry(line: &str) -> Option<(Vec<u8>, Vec<u8>, BfRangeDest)> {
    let trimmed = line.trim();
    let (start, end) = parse_two_hex_strings(trimmed)?;

    let first_end = trimmed.find('>')?;
    let rest_after_first = &trimmed[first_end + 1..];
    let second_start = rest_after_first.find('<')?;
    let rest_after_second_start = &rest_after_first[second_start..];
    let second_end = rest_after_second_start.find('>')?;
    let tail = rest_after_second_start[second_end + 1..].trim();

    if tail.starts_with('[') {
        let mut items = Vec::new();
        let mut cursor = tail;
        while let Some(begin) = cursor.find('<') {
            let rest = &cursor[begin..];
            let end_idx = rest.find('>')?;
            items.push(parse_hex_bytes(&rest[..=end_idx])?);
            cursor = &rest[end_idx + 1..];
        }
        Some((start, end, BfRangeDest::Explicit(items)))
    } else {
        let dest_start = tail.find('<')?;
        let dest_rest = &tail[dest_start..];
        let dest_end = dest_rest.find('>')?;
        let dest = parse_hex_bytes(&dest_rest[..=dest_end])?;
        Some((start, end, BfRangeDest::Sequential(dest)))
    }
}

fn parse_two_hex_strings(line: &str) -> Option<(Vec<u8>, Vec<u8>)> {
    let mut cursor = line;
    let first_start = cursor.find('<')?;
    cursor = &cursor[first_start..];
    let first_end = cursor.find('>')?;
    let first = parse_hex_bytes(&cursor[..=first_end])?;
    cursor = &cursor[first_end + 1..];
    let second_start = cursor.find('<')?;
    cursor = &cursor[second_start..];
    let second_end = cursor.find('>')?;
    let second = parse_hex_bytes(&cursor[..=second_end])?;
    Some((first, second))
}

fn parse_hex_bytes(token: &str) -> Option<Vec<u8>> {
    let mut trimmed = token.trim();
    if !trimmed.starts_with('<') || !trimmed.contains('>') {
        return None;
    }
    trimmed = trimmed.trim_start_matches('<');
    trimmed = trimmed.trim_end_matches('>');
    let mut cleaned = String::new();
    for ch in trimmed.chars() {
        if !ch.is_ascii_whitespace() {
            cleaned.push(ch);
        }
    }
    if cleaned.len() % 2 == 1 {
        cleaned.push('0');
    }
    let mut out = Vec::new();
    let bytes = cleaned.as_bytes();
    let mut index = 0usize;
    while index + 1 < bytes.len() {
        let hi = (bytes[index] as char).to_digit(16)? as u8;
        let lo = (bytes[index + 1] as char).to_digit(16)? as u8;
        out.push((hi << 4) | lo);
        index += 2;
    }
    Some(out)
}

fn increment_utf16be_bytes(bytes: &mut Vec<u8>) {
    if bytes.len() >= 2 {
        for chunk_index in (0..bytes.len()).step_by(2).rev() {
            let hi = bytes[chunk_index] as u16;
            let lo = bytes[chunk_index + 1] as u16;
            let value = (hi << 8) | lo;
            if value < 0xFFFF {
                let next = value + 1;
                bytes[chunk_index] = (next >> 8) as u8;
                bytes[chunk_index + 1] = next as u8;
                return;
            }
            bytes[chunk_index] = 0;
            bytes[chunk_index + 1] = 0;
        }
    }
}

fn decode_utf16be_fallback(bytes: &[u8]) -> String {
    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        return decode_utf16be_fallback(&bytes[2..]);
    }
    if bytes.len() % 2 == 0 && !bytes.is_empty() {
        let mut units = Vec::with_capacity(bytes.len() / 2);
        for chunk in bytes.chunks_exact(2) {
            units.push(u16::from_be_bytes([chunk[0], chunk[1]]));
        }
        if let Ok(text) = String::from_utf16(&units) {
            return text;
        }
    }
    bytes.iter().map(|&b| char::from(b)).collect()
}

fn hex_bytes_to_u32(bytes: &[u8]) -> u32 {
    let mut out = 0u32;
    for &byte in bytes {
        out = (out << 8) | byte as u32;
    }
    out
}

fn decode_pdf_string_bytes(bytes: &[u8]) -> Option<String> {
    if bytes.is_empty() {
        return Some(String::new());
    }
    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        return Some(decode_utf16be_fallback(bytes));
    }
    if bytes.len() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE {
        let mut units = Vec::new();
        for chunk in bytes[2..].chunks_exact(2) {
            units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        if let Ok(text) = String::from_utf16(&units) {
            return Some(text);
        }
    }
    Some(decode_simple_encoding(bytes, SimpleEncoding::WinAnsi))
}

fn decode_simple_encoding(bytes: &[u8], encoding: SimpleEncoding) -> String {
    match encoding {
        SimpleEncoding::Ascii => bytes.iter().map(|&byte| char::from(byte)).collect(),
        SimpleEncoding::WinAnsi => bytes.iter().map(|&byte| decode_win_ansi(byte)).collect(),
    }
}

fn decode_win_ansi(byte: u8) -> char {
    match byte {
        0x80 => '€',
        0x82 => '‚',
        0x83 => 'ƒ',
        0x84 => '„',
        0x85 => '…',
        0x86 => '†',
        0x87 => '‡',
        0x88 => 'ˆ',
        0x89 => '‰',
        0x8A => 'Š',
        0x8B => '‹',
        0x8C => 'Œ',
        0x8E => 'Ž',
        0x91 => '‘',
        0x92 => '’',
        0x93 => '“',
        0x94 => '”',
        0x95 => '•',
        0x96 => '–',
        0x97 => '—',
        0x98 => '˜',
        0x99 => '™',
        0x9A => 'š',
        0x9B => '›',
        0x9C => 'œ',
        0x9E => 'ž',
        0x9F => 'Ÿ',
        _ => byte as char,
    }
}

#[derive(Clone, Copy, Debug)]
struct Matrix {
    a: f32,
    b: f32,
    c: f32,
    d: f32,
    e: f32,
    f: f32,
}

impl Matrix {
    fn identity() -> Self {
        Self::new(1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    }

    fn new(a: f32, b: f32, c: f32, d: f32, e: f32, f: f32) -> Self {
        Self { a, b, c, d, e, f }
    }

    fn multiply(&self, other: &Self) -> Self {
        Self {
            a: self.a * other.a + self.c * other.b,
            b: self.b * other.a + self.d * other.b,
            c: self.a * other.c + self.c * other.d,
            d: self.b * other.c + self.d * other.d,
            e: self.a * other.e + self.c * other.f + self.e,
            f: self.b * other.e + self.d * other.f + self.f,
        }
    }

    fn translate(&self, tx: f32, ty: f32) -> Self {
        self.multiply(&Self::new(1.0, 0.0, 0.0, 1.0, tx, ty))
    }

    fn transform_point(&self, x: f32, y: f32) -> (f32, f32) {
        (
            self.a * x + self.c * y + self.e,
            self.b * x + self.d * y + self.f,
        )
    }
}

#[derive(Clone, Debug)]
struct TextState {
    font: String,
    font_size: f32,
    leading: f32,
    char_spacing: f32,
    word_spacing: f32,
    horizontal_scaling: f32,
    rise: f32,
    text_matrix: Matrix,
    line_matrix: Matrix,
}

impl Default for TextState {
    fn default() -> Self {
        Self {
            font: String::new(),
            font_size: 12.0,
            leading: 12.0,
            char_spacing: 0.0,
            word_spacing: 0.0,
            horizontal_scaling: 1.0,
            rise: 0.0,
            text_matrix: Matrix::identity(),
            line_matrix: Matrix::identity(),
        }
    }
}

impl TextState {
    fn translate(&mut self, tx: f32, ty: f32) {
        self.line_matrix = self.line_matrix.translate(tx, ty);
        self.text_matrix = self.line_matrix;
    }
}

fn show_pdf_text(
    resources: &PdfResources,
    graphics_stack: &[Matrix],
    text: &mut TextState,
    bytes: &[u8],
    spans: &mut Vec<PdfTextSpan>,
) {
    let font = resources.fonts.get(&text.font).cloned().unwrap_or_default();
    let decoded = font.decode(bytes);
    let display = decoded.replace('\u{0}', "");
    if display.trim().is_empty() {
        return;
    }
    let ctm = graphics_stack.last().copied().unwrap_or_else(Matrix::identity);
    let text_ctm = ctm.multiply(&text.text_matrix);
    let (x, y) = text_ctm.transform_point(0.0, text.rise);
    let advance = estimate_text_advance(&display, text.font_size, text.char_spacing, text.word_spacing, text.horizontal_scaling);
    spans.push(PdfTextSpan {
        x,
        y,
        end_x: x + advance,
        font_size: text.font_size,
        text: display,
    });
    text.text_matrix = text.text_matrix.translate(advance, 0.0);
}

fn estimate_text_advance(
    text: &str,
    font_size: f32,
    char_spacing: f32,
    word_spacing: f32,
    horizontal_scaling: f32,
) -> f32 {
    let mut width = 0.0;
    for ch in text.chars() {
        width += if ch.is_ascii_punctuation() {
            font_size * 0.35
        } else if ch.is_whitespace() {
            font_size * 0.33 + word_spacing
        } else if ch.is_ascii_digit() {
            font_size * 0.55
        } else if ch.is_ascii_uppercase() {
            font_size * 0.62
        } else {
            font_size * 0.52
        };
        width += char_spacing;
    }
    width * horizontal_scaling.max(0.01)
}

#[derive(Clone, Debug, PartialEq)]
enum ContentToken {
    Number(f32),
    Name(String),
    String(Vec<u8>),
    Array(Vec<ContentToken>),
    Operator(String),
}

impl ContentToken {
    fn as_number(self) -> Option<f32> {
        match self {
            Self::Number(value) => Some(value),
            _ => None,
        }
    }

    fn into_name(self) -> Option<String> {
        match self {
            Self::Name(name) => Some(name),
            _ => None,
        }
    }

    fn into_bytes(self) -> Option<Vec<u8>> {
        match self {
            Self::String(bytes) => Some(bytes),
            _ => None,
        }
    }

    fn into_array(self) -> Option<Vec<ContentToken>> {
        match self {
            Self::Array(items) => Some(items),
            _ => None,
        }
    }
}

fn take_six_numbers(operands: &mut Vec<ContentToken>) -> Option<[f32; 6]> {
    if operands.len() < 6 {
        return None;
    }
    let mut values = [0.0; 6];
    for index in (0..6).rev() {
        values[index] = operands.pop()?.as_number()?;
    }
    Some(values)
}

fn take_two_numbers(operands: &mut Vec<ContentToken>) -> Option<[f32; 2]> {
    if operands.len() < 2 {
        return None;
    }
    let second = operands.pop()?.as_number()?;
    let first = operands.pop()?.as_number()?;
    Some([first, second])
}

struct ContentParser<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> ContentParser<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }

    fn next_token(&mut self) -> Result<Option<ContentToken>> {
        self.skip_ws_and_comments();
        let Some(byte) = self.peek() else {
            return Ok(None);
        };
        let token = match byte {
            b'/' => ContentToken::Name(self.parse_name()?),
            b'(' => ContentToken::String(self.parse_literal_string()?),
            b'<' if self.peek_n(1) != Some(b'<') => ContentToken::String(self.parse_hex_string()?),
            b'[' => ContentToken::Array(self.parse_array()?),
            b'+' | b'-' | b'.' | b'0'..=b'9' => ContentToken::Number(self.parse_number()? as f32),
            _ => ContentToken::Operator(self.parse_operator()),
        };
        Ok(Some(token))
    }

    fn skip_inline_image(&mut self) {
        if let Some(index) = self.data[self.pos..].windows(2).position(|window| window == b"EI") {
            self.pos += index + 2;
        } else {
            self.pos = self.data.len();
        }
    }

    fn parse_array(&mut self) -> Result<Vec<ContentToken>> {
        self.expect(b'[')?;
        let mut items = Vec::new();
        loop {
            self.skip_ws_and_comments();
            if self.peek() == Some(b']') {
                self.pos += 1;
                break;
            }
            let Some(token) = self.next_token()? else {
                return Err(LoError::Parse("unterminated pdf content array".to_string()));
            };
            items.push(token);
        }
        Ok(items)
    }

    fn parse_name(&mut self) -> Result<String> {
        self.expect(b'/')?;
        let start = self.pos;
        while let Some(byte) = self.peek() {
            if is_pdf_space(byte) || is_pdf_delim(byte) {
                break;
            }
            self.pos += 1;
        }
        Ok(String::from_utf8_lossy(&self.data[start..self.pos]).to_string())
    }

    fn parse_number(&mut self) -> Result<f64> {
        let start = self.pos;
        if matches!(self.peek(), Some(b'+') | Some(b'-')) {
            self.pos += 1;
        }
        while let Some(byte) = self.peek() {
            if byte.is_ascii_digit() || byte == b'.' {
                self.pos += 1;
            } else {
                break;
            }
        }
        let text = std::str::from_utf8(&self.data[start..self.pos])
            .map_err(|err| LoError::Parse(format!("invalid pdf content number: {err}")))?;
        text.parse::<f64>()
            .map_err(|err| LoError::Parse(format!("invalid pdf content number: {err}")))
    }

    fn parse_operator(&mut self) -> String {
        let start = self.pos;
        while let Some(byte) = self.peek() {
            if is_pdf_space(byte) || matches!(byte, b'[' | b']' | b'<' | b'>' | b'(' | b')' | b'/' ) {
                break;
            }
            self.pos += 1;
        }
        String::from_utf8_lossy(&self.data[start..self.pos]).to_string()
    }

    fn parse_hex_string(&mut self) -> Result<Vec<u8>> {
        self.expect(b'<')?;
        let start = self.pos;
        while let Some(byte) = self.peek() {
            if byte == b'>' {
                break;
            }
            self.pos += 1;
        }
        let end = self.pos;
        self.expect(b'>')?;
        parse_hex_bytes(std::str::from_utf8(&self.data[start - 1..=end]).unwrap_or(""))
            .ok_or_else(|| LoError::Parse("invalid pdf content hex string".to_string()))
    }

    fn parse_literal_string(&mut self) -> Result<Vec<u8>> {
        self.expect(b'(')?;
        let mut depth = 1usize;
        let mut out = Vec::new();
        while let Some(byte) = self.next_byte() {
            match byte {
                b'(' => {
                    depth += 1;
                    out.push(byte);
                }
                b')' => {
                    depth -= 1;
                    if depth == 0 {
                        return Ok(out);
                    }
                    out.push(byte);
                }
                b'\\' => {
                    let Some(escaped) = self.next_byte() else { break };
                    match escaped {
                        b'n' => out.push(b'\n'),
                        b'r' => out.push(b'\r'),
                        b't' => out.push(b'\t'),
                        b'b' => out.push(0x08),
                        b'f' => out.push(0x0C),
                        b'(' | b')' | b'\\' => out.push(escaped),
                        b'\r' => {
                            if self.peek() == Some(b'\n') {
                                self.pos += 1;
                            }
                        }
                        b'\n' => {}
                        b'0'..=b'7' => {
                            let mut value = (escaped - b'0') as u16;
                            for _ in 0..2 {
                                match self.peek() {
                                    Some(next @ b'0'..=b'7') => {
                                        self.pos += 1;
                                        value = (value << 3) | (next - b'0') as u16;
                                    }
                                    _ => break,
                                }
                            }
                            out.push(value as u8);
                        }
                        other => out.push(other),
                    }
                }
                _ => out.push(byte),
            }
        }
        Err(LoError::Parse("unterminated pdf content string".to_string()))
    }

    fn skip_ws_and_comments(&mut self) {
        loop {
            while let Some(byte) = self.peek() {
                if is_pdf_space(byte) {
                    self.pos += 1;
                } else {
                    break;
                }
            }
            if self.peek() == Some(b'%') {
                while let Some(byte) = self.next_byte() {
                    if byte == b'\n' || byte == b'\r' {
                        break;
                    }
                }
                continue;
            }
            break;
        }
    }

    fn expect(&mut self, byte: u8) -> Result<()> {
        match self.next_byte() {
            Some(value) if value == byte => Ok(()),
            other => Err(LoError::Parse(format!(
                "expected byte {byte:?}, got {other:?}"
            ))),
        }
    }

    fn peek(&self) -> Option<u8> {
        self.data.get(self.pos).copied()
    }

    fn peek_n(&self, offset: usize) -> Option<u8> {
        self.data.get(self.pos + offset).copied()
    }

    fn next_byte(&mut self) -> Option<u8> {
        let byte = self.peek()?;
        self.pos += 1;
        Some(byte)
    }
}

struct Parser<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> Parser<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }

    fn new_at(data: &'a [u8], pos: usize) -> Self {
        Self { data, pos }
    }

    fn parse_value(&mut self) -> Result<PdfValue> {
        self.skip_ws_and_comments();
        match self.peek() {
            Some(b'/') => Ok(PdfValue::Name(self.parse_name()?)),
            Some(b'(') => Ok(PdfValue::String(self.parse_literal_string()?)),
            Some(b'<') if self.peek_n(1) == Some(b'<') => Ok(PdfValue::Dict(self.parse_dict()?)),
            Some(b'<') => Ok(PdfValue::String(self.parse_hex_string()?)),
            Some(b'[') => Ok(PdfValue::Array(self.parse_array()?)),
            Some(b't') if self.consume_keyword(b"true") => Ok(PdfValue::Bool(true)),
            Some(b'f') if self.consume_keyword(b"false") => Ok(PdfValue::Bool(false)),
            Some(b'n') if self.consume_keyword(b"null") => Ok(PdfValue::Null),
            Some(b'+') | Some(b'-') | Some(b'.') | Some(b'0'..=b'9') => self.parse_number_or_ref(),
            other => Err(LoError::Parse(format!("unexpected pdf token at {:?}", other))),
        }
    }

    fn parse_unsigned_number(&mut self) -> Result<u32> {
        let start = self.pos;
        while let Some(byte) = self.peek() {
            if byte.is_ascii_digit() {
                self.pos += 1;
            } else {
                break;
            }
        }
        if self.pos == start {
            return Err(LoError::Parse("expected unsigned integer".to_string()));
        }
        std::str::from_utf8(&self.data[start..self.pos])
            .map_err(|err| LoError::Parse(format!("invalid pdf integer: {err}")))?
            .parse::<u32>()
            .map_err(|err| LoError::Parse(format!("invalid pdf integer: {err}")))
    }

    fn parse_number_or_ref(&mut self) -> Result<PdfValue> {
        let first_start = self.pos;
        let first = self.parse_number_token()?;
        let after_first = self.pos;
        if let Some(first_int) = first.integer {
            self.skip_ws_and_comments();
            let save = self.pos;
            if let Ok(second) = self.parse_number_token() {
                if let Some(second_int) = second.integer {
                    self.skip_ws_and_comments();
                    if self.consume_keyword(b"R") {
                        return Ok(PdfValue::Ref(PdfObjectId::new(first_int as u32, second_int as u16)));
                    }
                }
            }
            self.pos = save;
        }
        self.pos = after_first;
        let text = std::str::from_utf8(&self.data[first_start..after_first])
            .map_err(|err| LoError::Parse(format!("invalid pdf number: {err}")))?;
        let number = text
            .parse::<f64>()
            .map_err(|err| LoError::Parse(format!("invalid pdf number: {err}")))?;
        Ok(PdfValue::Number(number))
    }

    fn parse_number_token(&mut self) -> Result<NumberToken> {
        let start = self.pos;
        if matches!(self.peek(), Some(b'+') | Some(b'-')) {
            self.pos += 1;
        }
        let mut has_dot = false;
        while let Some(byte) = self.peek() {
            if byte.is_ascii_digit() {
                self.pos += 1;
            } else if byte == b'.' && !has_dot {
                has_dot = true;
                self.pos += 1;
            } else {
                break;
            }
        }
        if self.pos == start {
            return Err(LoError::Parse("expected pdf number".to_string()));
        }
        let text = std::str::from_utf8(&self.data[start..self.pos])
            .map_err(|err| LoError::Parse(format!("invalid pdf number: {err}")))?
            .to_string();
        let integer = if !text.contains('.') {
            text.parse::<i64>().ok()
        } else {
            None
        };
        Ok(NumberToken { integer })
    }

    fn parse_name(&mut self) -> Result<String> {
        self.expect(b'/')?;
        let mut out = String::new();
        while let Some(byte) = self.peek() {
            if is_pdf_space(byte) || is_pdf_delim(byte) {
                break;
            }
            if byte == b'#' {
                self.pos += 1;
                let a = self.next_byte().ok_or_else(|| LoError::Parse("truncated pdf name hex escape".to_string()))?;
                let b = self.next_byte().ok_or_else(|| LoError::Parse("truncated pdf name hex escape".to_string()))?;
                let hi = (a as char)
                    .to_digit(16)
                    .ok_or_else(|| LoError::Parse("invalid pdf name hex escape".to_string()))? as u8;
                let lo = (b as char)
                    .to_digit(16)
                    .ok_or_else(|| LoError::Parse("invalid pdf name hex escape".to_string()))? as u8;
                out.push(((hi << 4) | lo) as char);
                continue;
            }
            self.pos += 1;
            out.push(byte as char);
        }
        Ok(out)
    }

    fn parse_array(&mut self) -> Result<Vec<PdfValue>> {
        self.expect(b'[')?;
        let mut items = Vec::new();
        loop {
            self.skip_ws_and_comments();
            if self.peek() == Some(b']') {
                self.pos += 1;
                break;
            }
            items.push(self.parse_value()?);
        }
        Ok(items)
    }

    fn parse_dict(&mut self) -> Result<BTreeMap<String, PdfValue>> {
        self.expect(b'<')?;
        self.expect(b'<')?;
        let mut dict = BTreeMap::new();
        loop {
            self.skip_ws_and_comments();
            if self.peek() == Some(b'>') && self.peek_n(1) == Some(b'>') {
                self.pos += 2;
                break;
            }
            let key = self.parse_name()?;
            let value = self.parse_value()?;
            dict.insert(key, value);
        }
        Ok(dict)
    }

    fn parse_literal_string(&mut self) -> Result<Vec<u8>> {
        let mut content = ContentParser::new(&self.data[self.pos..]);
        let value = content.parse_literal_string()?;
        self.pos += content.pos;
        Ok(value)
    }

    fn parse_hex_string(&mut self) -> Result<Vec<u8>> {
        let mut content = ContentParser::new(&self.data[self.pos..]);
        let value = content.parse_hex_string()?;
        self.pos += content.pos;
        Ok(value)
    }

    fn skip_ws_and_comments(&mut self) {
        loop {
            while let Some(byte) = self.peek() {
                if is_pdf_space(byte) {
                    self.pos += 1;
                } else {
                    break;
                }
            }
            if self.peek() == Some(b'%') {
                while let Some(byte) = self.next_byte() {
                    if byte == b'\n' || byte == b'\r' {
                        break;
                    }
                }
                continue;
            }
            break;
        }
    }

    fn consume_keyword(&mut self, keyword: &[u8]) -> bool {
        if self.data.get(self.pos..self.pos + keyword.len()) == Some(keyword) {
            self.pos += keyword.len();
            true
        } else {
            false
        }
    }

    fn expect(&mut self, byte: u8) -> Result<()> {
        match self.next_byte() {
            Some(value) if value == byte => Ok(()),
            other => Err(LoError::Parse(format!(
                "expected byte {byte:?}, got {other:?}"
            ))),
        }
    }

    fn peek(&self) -> Option<u8> {
        self.data.get(self.pos).copied()
    }

    fn peek_n(&self, offset: usize) -> Option<u8> {
        self.data.get(self.pos + offset).copied()
    }

    fn next_byte(&mut self) -> Option<u8> {
        let byte = self.peek()?;
        self.pos += 1;
        Some(byte)
    }
}

struct NumberToken {
    integer: Option<i64>,
}

fn is_pdf_space(byte: u8) -> bool {
    matches!(byte, b'\0' | b'\t' | b'\n' | 0x0C | b'\r' | b' ')
}

fn is_pdf_delim(byte: u8) -> bool {
    matches!(byte, b'(' | b')' | b'<' | b'>' | b'[' | b']' | b'{' | b'}' | b'/' | b'%')
}

fn collect_filter_names(value: Option<&PdfValue>) -> Vec<String> {
    match value {
        Some(PdfValue::Name(name)) => vec![name.clone()],
        Some(PdfValue::Array(items)) => items
            .iter()
            .filter_map(|item| item.as_name().map(str::to_string))
            .collect(),
        _ => Vec::new(),
    }
}

fn decode_ascii_hex(data: &[u8]) -> Result<Vec<u8>> {
    let mut cleaned = Vec::new();
    for &byte in data {
        if byte == b'>' {
            break;
        }
        if !is_pdf_space(byte) {
            cleaned.push(byte);
        }
    }
    if cleaned.len() % 2 == 1 {
        cleaned.push(b'0');
    }
    let mut out = Vec::with_capacity(cleaned.len() / 2);
    let mut index = 0usize;
    while index + 1 < cleaned.len() {
        let hi = (cleaned[index] as char)
            .to_digit(16)
            .ok_or_else(|| LoError::Parse("invalid ASCIIHex digit".to_string()))? as u8;
        let lo = (cleaned[index + 1] as char)
            .to_digit(16)
            .ok_or_else(|| LoError::Parse("invalid ASCIIHex digit".to_string()))? as u8;
        out.push((hi << 4) | lo);
        index += 2;
    }
    Ok(out)
}

fn decode_ascii85(data: &[u8]) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    let mut group = [0u32; 5];
    let mut count = 0usize;
    let mut index = 0usize;
    while index < data.len() {
        let byte = data[index];
        index += 1;
        match byte {
            b'~' => break,
            b'z' if count == 0 => {
                out.extend_from_slice(&[0, 0, 0, 0]);
            }
            _ if is_pdf_space(byte) => continue,
            b'!'..=b'u' => {
                group[count] = (byte - b'!') as u32;
                count += 1;
                if count == 5 {
                    let mut value = 0u32;
                    for item in group {
                        value = value.saturating_mul(85).saturating_add(item);
                    }
                    out.extend_from_slice(&value.to_be_bytes());
                    count = 0;
                }
            }
            other => {
                return Err(LoError::Parse(format!(
                    "invalid ASCII85 byte: {other}"
                )))
            }
        }
    }
    if count > 0 {
        for slot in group.iter_mut().skip(count) {
            *slot = 84;
        }
        let mut value = 0u32;
        for item in group {
            value = value.saturating_mul(85).saturating_add(item);
        }
        let bytes = value.to_be_bytes();
        out.extend_from_slice(&bytes[..count - 1]);
    }
    Ok(out)
}

fn decode_run_length(data: &[u8]) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    let mut index = 0usize;
    while index < data.len() {
        let len = data[index];
        index += 1;
        match len {
            128 => break,
            0..=127 => {
                let count = len as usize + 1;
                if index + count > data.len() {
                    return Err(LoError::Parse("truncated RunLength stream".to_string()));
                }
                out.extend_from_slice(&data[index..index + count]);
                index += count;
            }
            129..=255 => {
                let count = 257usize - len as usize;
                let value = *data
                    .get(index)
                    .ok_or_else(|| LoError::Parse("truncated RunLength repeat".to_string()))?;
                index += 1;
                for _ in 0..count {
                    out.push(value);
                }
            }
        }
    }
    Ok(out)
}

fn decode_flate_stream(data: &[u8]) -> Result<Vec<u8>> {
    if data.len() >= 2 {
        let cmf = data[0];
        let flg = data[1];
        if (cmf & 0x0F) == 8 && ((cmf as u16) << 8 | flg as u16) % 31 == 0 {
            if data.len() < 6 {
                return Err(LoError::Parse("truncated zlib stream".to_string()));
            }
            let start = 2 + if flg & 0x20 != 0 { 4 } else { 0 };
            if start > data.len().saturating_sub(4) {
                return Err(LoError::Parse("invalid zlib stream".to_string()));
            }
            return inflate_raw_deflate(&data[start..data.len() - 4], 0);
        }
    }
    inflate_raw_deflate(data, 0)
}

fn decode_lzw(data: &[u8]) -> Result<Vec<u8>> {
    let mut reader = MsbBitReader::new(data);
    let mut dict: Vec<Vec<u8>> = vec![Vec::new(); 258];
    for value in 0u16..=255 {
        dict[value as usize] = vec![value as u8];
    }
    let clear = 256u16;
    let eod = 257u16;
    let mut code_size = 9u8;
    let mut next_code = 258u16;
    let mut prev: Option<Vec<u8>> = None;
    let mut out = Vec::new();
    while let Some(code) = reader.read_bits(code_size)? {
        let code = code as u16;
        if code == clear {
            dict.truncate(258);
            if dict.len() < 258 {
                dict.resize(258, Vec::new());
            }
            for value in 0u16..=255 {
                if dict[value as usize].is_empty() {
                    dict[value as usize] = vec![value as u8];
                }
            }
            code_size = 9;
            next_code = 258;
            prev = None;
            continue;
        }
        if code == eod {
            break;
        }
        let entry = if let Some(existing) = dict.get(code as usize).filter(|entry| !entry.is_empty()) {
            existing.clone()
        } else if code == next_code {
            let mut generated = prev.clone().ok_or_else(|| {
                LoError::Parse("invalid LZW back-reference".to_string())
            })?;
            let first = *generated.first().ok_or_else(|| {
                LoError::Parse("empty LZW previous entry".to_string())
            })?;
            generated.push(first);
            generated
        } else {
            return Err(LoError::Parse(format!("invalid LZW code {code}")));
        };
        out.extend_from_slice(&entry);
        if let Some(previous) = prev.take() {
            let mut new_entry = previous;
            if let Some(&first) = entry.first() {
                new_entry.push(first);
                if dict.len() == next_code as usize {
                    dict.push(new_entry);
                } else if dict.len() > next_code as usize {
                    dict[next_code as usize] = new_entry;
                } else {
                    while dict.len() < next_code as usize {
                        dict.push(Vec::new());
                    }
                    dict.push(new_entry);
                }
                next_code += 1;
                if next_code == (1u16 << code_size) && code_size < 12 {
                    code_size += 1;
                }
            }
        }
        prev = Some(entry);
    }
    Ok(out)
}

struct MsbBitReader<'a> {
    data: &'a [u8],
    bit_pos: usize,
}

impl<'a> MsbBitReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, bit_pos: 0 }
    }

    fn read_bits(&mut self, count: u8) -> Result<Option<u32>> {
        if count == 0 {
            return Ok(Some(0));
        }
        if self.bit_pos + count as usize > self.data.len() * 8 {
            return Ok(None);
        }
        let mut out = 0u32;
        for _ in 0..count {
            let byte_index = self.bit_pos / 8;
            let bit_in_byte = 7 - (self.bit_pos % 8);
            let byte = self.data[byte_index];
            out = (out << 1) | ((byte >> bit_in_byte) & 1) as u32;
            self.bit_pos += 1;
        }
        Ok(Some(out))
    }
}

// Raw deflate decoder adapted from the workspace ZIP reader so PDF parsing can
// stay dependency-free and shared.
struct BitReader<'a> {
    data: &'a [u8],
    bit_pos: usize,
}

impl<'a> BitReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, bit_pos: 0 }
    }

    fn read_bits(&mut self, count: u8) -> Result<u32> {
        if count == 0 {
            return Ok(0);
        }
        let mut out = 0u32;
        for bit_index in 0..count {
            let byte_index = self.bit_pos / 8;
            let bit_in_byte = self.bit_pos % 8;
            let byte = *self
                .data
                .get(byte_index)
                .ok_or_else(|| LoError::Parse("unexpected end of deflate stream".to_string()))?;
            let bit = (byte >> bit_in_byte) & 1;
            out |= (bit as u32) << bit_index;
            self.bit_pos += 1;
        }
        Ok(out)
    }

    fn align_byte(&mut self) {
        self.bit_pos = (self.bit_pos + 7) & !7;
    }

    fn read_aligned_bytes(&mut self, len: usize) -> Result<&'a [u8]> {
        self.align_byte();
        let start = self.bit_pos / 8;
        let end = start + len;
        let slice = self.data.get(start..end).ok_or_else(|| {
            LoError::Parse("unexpected end of aligned deflate block".to_string())
        })?;
        self.bit_pos += len * 8;
        Ok(slice)
    }
}

#[derive(Clone, Debug)]
struct Huffman {
    max_len: u8,
    table: BTreeMap<(u8, u16), u16>,
}

impl Huffman {
    fn from_code_lengths(lengths: &[u8]) -> Result<Self> {
        let max_len = *lengths.iter().max().unwrap_or(&0);
        if max_len == 0 {
            return Err(LoError::Parse("empty huffman table".to_string()));
        }
        let mut bl_count = vec![0u16; max_len as usize + 1];
        for &len in lengths {
            if len > 0 {
                bl_count[len as usize] += 1;
            }
        }
        let mut next_code = vec![0u16; max_len as usize + 1];
        let mut code = 0u16;
        for bits in 1..=max_len as usize {
            code = (code + bl_count[bits - 1]) << 1;
            next_code[bits] = code;
        }
        let mut table = BTreeMap::new();
        for (symbol, &len) in lengths.iter().enumerate() {
            if len == 0 {
                continue;
            }
            let canonical = next_code[len as usize];
            next_code[len as usize] += 1;
            let reversed = reverse_bits(canonical, len);
            table.insert((len, reversed), symbol as u16);
        }
        Ok(Self { max_len, table })
    }

    fn decode_symbol(&self, bits: &mut BitReader<'_>) -> Result<u16> {
        let mut code = 0u16;
        for len in 1..=self.max_len {
            let bit = bits.read_bits(1)? as u16;
            code |= bit << (len - 1);
            if let Some(symbol) = self.table.get(&(len, code)) {
                return Ok(*symbol);
            }
        }
        Err(LoError::Parse("invalid huffman code".to_string()))
    }
}

fn reverse_bits(mut code: u16, len: u8) -> u16 {
    let mut out = 0u16;
    for _ in 0..len {
        out = (out << 1) | (code & 1);
        code >>= 1;
    }
    out
}

fn inflate_raw_deflate(data: &[u8], expected_len: usize) -> Result<Vec<u8>> {
    let mut reader = BitReader::new(data);
    let mut out = Vec::with_capacity(expected_len.max(256));
    loop {
        let is_final = reader.read_bits(1)? != 0;
        let block_type = reader.read_bits(2)? as u8;
        match block_type {
            0 => read_stored_block(&mut reader, &mut out)?,
            1 => {
                let litlen = fixed_literal_huffman()?;
                let dist = fixed_distance_huffman()?;
                read_huffman_block(&mut reader, &litlen, &dist, &mut out)?;
            }
            2 => {
                let (litlen, dist) = read_dynamic_huffman_tables(&mut reader)?;
                read_huffman_block(&mut reader, &litlen, &dist, &mut out)?;
            }
            3 => return Err(LoError::Parse("reserved deflate block type".to_string())),
            _ => unreachable!(),
        }
        if is_final {
            break;
        }
    }
    Ok(out)
}

fn read_stored_block(reader: &mut BitReader<'_>, out: &mut Vec<u8>) -> Result<()> {
    reader.align_byte();
    let header = reader.read_aligned_bytes(4)?;
    let len = u16::from_le_bytes([header[0], header[1]]);
    let nlen = u16::from_le_bytes([header[2], header[3]]);
    if len != !nlen {
        return Err(LoError::Parse(
            "stored deflate block length checksum mismatch".to_string(),
        ));
    }
    let bytes = reader.read_aligned_bytes(len as usize)?;
    out.extend_from_slice(bytes);
    Ok(())
}

fn read_dynamic_huffman_tables(reader: &mut BitReader<'_>) -> Result<(Huffman, Huffman)> {
    let hlit = reader.read_bits(5)? as usize + 257;
    let hdist = reader.read_bits(5)? as usize + 1;
    let hclen = reader.read_bits(4)? as usize + 4;

    let order = [16usize, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
    let mut code_lengths = vec![0u8; 19];
    for i in 0..hclen {
        code_lengths[order[i]] = reader.read_bits(3)? as u8;
    }
    let code_length_huffman = Huffman::from_code_lengths(&code_lengths)?;

    let total = hlit + hdist;
    let mut lengths = Vec::with_capacity(total);
    while lengths.len() < total {
        match code_length_huffman.decode_symbol(reader)? {
            symbol @ 0..=15 => lengths.push(symbol as u8),
            16 => {
                let repeat = reader.read_bits(2)? as usize + 3;
                let previous = *lengths.last().ok_or_else(|| {
                    LoError::Parse("repeat code without previous code length".to_string())
                })?;
                for _ in 0..repeat {
                    lengths.push(previous);
                }
            }
            17 => {
                let repeat = reader.read_bits(3)? as usize + 3;
                for _ in 0..repeat {
                    lengths.push(0u8);
                }
            }
            18 => {
                let repeat = reader.read_bits(7)? as usize + 11;
                for _ in 0..repeat {
                    lengths.push(0u8);
                }
            }
            other => {
                return Err(LoError::Parse(format!(
                    "invalid code length symbol {other}"
                )))
            }
        }
    }

    let litlen = Huffman::from_code_lengths(&lengths[..hlit])?;
    let dist_lengths = &lengths[hlit..hlit + hdist];
    let dist = if dist_lengths.iter().all(|&len| len == 0) {
        Huffman::from_code_lengths(&[1])?
    } else {
        Huffman::from_code_lengths(dist_lengths)?
    };
    Ok((litlen, dist))
}

fn read_huffman_block(
    reader: &mut BitReader<'_>,
    litlen: &Huffman,
    dist: &Huffman,
    out: &mut Vec<u8>,
) -> Result<()> {
    loop {
        let symbol = litlen.decode_symbol(reader)?;
        match symbol {
            0..=255 => out.push(symbol as u8),
            256 => return Ok(()),
            257..=285 => {
                let length = decode_length(reader, symbol)?;
                let distance_symbol = dist.decode_symbol(reader)?;
                let distance = decode_distance(reader, distance_symbol)?;
                if distance == 0 || distance > out.len() {
                    return Err(LoError::Parse(
                        "invalid deflate back-reference distance".to_string(),
                    ));
                }
                let start = out.len() - distance;
                for i in 0..length {
                    let byte = out[start + (i % distance)];
                    out.push(byte);
                }
            }
            other => {
                return Err(LoError::Parse(format!(
                    "invalid deflate literal/length symbol {other}"
                )))
            }
        }
    }
}

fn fixed_literal_huffman() -> Result<Huffman> {
    let mut lengths = vec![0u8; 288];
    for symbol in 0..=143 {
        lengths[symbol] = 8;
    }
    for symbol in 144..=255 {
        lengths[symbol] = 9;
    }
    for symbol in 256..=279 {
        lengths[symbol] = 7;
    }
    for symbol in 280..=287 {
        lengths[symbol] = 8;
    }
    Huffman::from_code_lengths(&lengths)
}

fn fixed_distance_huffman() -> Result<Huffman> {
    Huffman::from_code_lengths(&[5u8; 32])
}

fn decode_length(reader: &mut BitReader<'_>, symbol: u16) -> Result<usize> {
    const BASES: [usize; 29] = [
        3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99,
        115, 131, 163, 195, 227, 258,
    ];
    const EXTRA: [u8; 29] = [
        0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
    ];
    if !(257..=285).contains(&symbol) {
        return Err(LoError::Parse(format!(
            "invalid deflate length symbol {symbol}"
        )));
    }
    if symbol == 285 {
        return Ok(258);
    }
    let index = (symbol - 257) as usize;
    let extra_bits = EXTRA[index];
    let extra = if extra_bits == 0 {
        0
    } else {
        reader.read_bits(extra_bits)? as usize
    };
    Ok(BASES[index] + extra)
}

fn decode_distance(reader: &mut BitReader<'_>, symbol: u16) -> Result<usize> {
    const BASES: [usize; 30] = [
        1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025,
        1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
    ];
    const EXTRA: [u8; 30] = [
        0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12,
        12, 13, 13,
    ];
    let index = symbol as usize;
    if index >= BASES.len() {
        return Err(LoError::Parse(format!(
            "invalid deflate distance symbol {symbol}"
        )));
    }
    let extra_bits = EXTRA[index];
    let extra = if extra_bits == 0 {
        0
    } else {
        reader.read_bits(extra_bits)? as usize
    };
    Ok(BASES[index] + extra)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_pdf_header_and_eof() {
        let pdf = write_text_pdf(
            &["Hello, world".to_string(), "Second line".to_string()],
            Length::pt(595.0),
            Length::pt(842.0),
        );
        assert!(pdf.starts_with(b"%PDF-1.4"));
        assert!(pdf.ends_with(b"%%EOF\n"));
        assert!(pdf.windows(4).any(|w| w == b"xref"));
    }

    #[test]
    fn escapes_pdf_text_special_chars() {
        let pdf = write_text_pdf(&["a (b) \\c".to_string()], Length::pt(100.0), Length::pt(100.0));
        let s = String::from_utf8_lossy(&pdf);
        assert!(s.contains("a \\(b\\) \\\\c"));
    }

    #[test]
    fn parses_simple_text_pdf() {
        let pdf = write_text_pdf(&["Hello PDF".to_string()], Length::pt(200.0), Length::pt(200.0));
        let extracted = extract_text_from_pdf(&pdf).unwrap();
        assert!(extracted.contains("Hello PDF"));
    }

    #[test]
    fn decodes_ascii_hex() {
        assert_eq!(decode_ascii_hex(b"48656C6C6F>").unwrap(), b"Hello");
    }

    #[test]
    fn decodes_run_length() {
        let data = [2u8, b'a', b'b', b'c', 253, b'Z', 128];
        assert_eq!(decode_run_length(&data).unwrap(), b"abcZZZZ");
    }
}
