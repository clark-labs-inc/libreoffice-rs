//! High-level pure-Rust convenience helpers that mirror Clark's current
//! `soffice --headless` usage without relying on LibreOffice itself.
//!
//! The Clark-focused surface in this crate is:
//! - visual `DOCX -> PDF`
//! - visual `PPTX -> PDF`
//! - `DOC -> DOCX`
//! - `XLSX` recalc with cached `<v>` patching
//! - tracked-change acceptance for `DOCX`
//! - generic `convert_bytes` / `convert_bytes_auto`
//! - JSON recalc reports compatible with Clark's existing `recalc.py`
//! - direct DOCX/PPTX page rasterization to PNG/JPEG
//! - Markdown extraction for DOCX/PPTX/XLSX
//! - PDF -> TXT/MD/HTML via the native PDF reader

mod xlsx_eval;

use std::collections::{BTreeMap, BTreeSet};
use std::path::Path;

use lo_core::{
    parse_xml_document, serialize_xml_document, CellAddr, LoError, Result, Workbook,
    XmlItem, XmlNode,
};
use lo_zip::{normalize_zip_path, rels_path_for, resolve_part_target, ZipArchive};
use xlsx_eval::{translate_shared_formula, EvalValue, WorkbookEvaluator};

/// Convert a DOCX byte stream into a PDF using Writer's native Rust
/// layout/rendering path.
pub fn docx_to_pdf_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    let doc = lo_writer::from_docx_bytes("document", bytes)?;
    lo_writer::save_as(&doc, "pdf")
}

/// Convert a legacy binary `.doc` file (Word 97-2003) into a DOCX byte
/// stream by extracting the piece-table text and re-emitting it.
pub fn doc_to_docx_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    let doc = lo_writer::from_doc_bytes("document", bytes)?;
    lo_writer::save_as(&doc, "docx")
}

/// Convert a PPTX byte stream into a PDF using Impress's native Rust
/// renderer.
pub fn pptx_to_pdf_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    let deck = lo_impress::from_pptx_bytes("presentation", bytes)?;
    lo_impress::save_as(&deck, "pdf")
}

// ---------------------------------------------------------------------------
// Generic format conversion and sniffing
// ---------------------------------------------------------------------------

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Family {
    Writer,
    Calc,
    Impress,
    Draw,
    Math,
    Base,
}

fn canonical_format_hint(format: &str) -> String {
    let trimmed = format.trim();
    let trimmed = trimmed.strip_prefix('.').unwrap_or(trimmed);
    let head = trimmed.split(':').next().unwrap_or(trimmed).trim();
    match head.to_ascii_lowercase().as_str() {
        "text" => "txt".to_string(),
        "markdown" => "md".to_string(),
        "htm" => "html".to_string(),
        "mml" => "mathml".to_string(),
        "odfmath" | "odf-formula" => "odf".to_string(),
        other => other.to_string(),
    }
}

/// Infer a format hint from a file path by looking at its extension.
pub fn sniff_format_from_path(path: &str) -> Option<String> {
    let ext = Path::new(path).extension()?.to_str()?;
    Some(canonical_format_hint(ext))
}

/// Infer a format from raw bytes.
///
/// Covers the Clark-heavy cases:
/// - OOXML packages (`docx`/`xlsx`/`pptx`)
/// - ODF packages (`odt`/`ods`/`odp`)
/// - legacy CFB files (`doc`/`xls`/`ppt`)
/// - PDF documents (`pdf`)
/// - plain-text-ish payloads (`txt`/`md`/`html`/`csv`/`svg`)
pub fn sniff_format_from_bytes(bytes: &[u8]) -> Option<String> {
    if bytes.len() >= 4 && &bytes[..4] == b"PK\x03\x04" {
        let zip = ZipArchive::new(bytes).ok()?;
        if zip.contains("[Content_Types].xml") {
            let content_types = zip.read_string("[Content_Types].xml").ok()?;
            let lower = content_types.to_ascii_lowercase();
            if lower.contains("wordprocessingml") {
                return Some("docx".to_string());
            }
            if lower.contains("spreadsheetml") {
                return Some("xlsx".to_string());
            }
            if lower.contains("presentationml") {
                return Some("pptx".to_string());
            }
        }
        if zip.contains("mimetype") {
            let mimetype = zip.read_string("mimetype").ok()?.to_ascii_lowercase();
            if mimetype.contains("opendocument.text") {
                return Some("odt".to_string());
            }
            if mimetype.contains("opendocument.spreadsheet") {
                return Some("ods".to_string());
            }
            if mimetype.contains("opendocument.presentation") {
                return Some("odp".to_string());
            }
        }
    }
    if bytes.len() >= 8 && bytes[..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        if find_bytes(bytes, b"WordDocument") {
            return Some("doc".to_string());
        }
        if find_bytes(bytes, b"Workbook") {
            return Some("xls".to_string());
        }
        if find_bytes(bytes, b"PowerPoint Document") {
            return Some("ppt".to_string());
        }
    }
    let header_len = bytes.len().min(1024);
    if bytes[..header_len].windows(5).any(|window| window == b"%PDF-") {
        return Some("pdf".to_string());
    }
    let text = std::str::from_utf8(bytes).ok()?.trim_start_matches('\u{feff}');
    if text.starts_with("<svg") || text.contains("<svg") {
        return Some("svg".to_string());
    }
    if text.starts_with("<!doctype html") || text.starts_with("<html") || text.contains("<body") {
        return Some("html".to_string());
    }
    if text.starts_with('#') || text.contains("\n# ") || text.contains("\n- ") {
        return Some("md".to_string());
    }
    if text.contains(',') && text.lines().count() > 1 {
        return Some("csv".to_string());
    }
    if !text.is_empty() {
        return Some("txt".to_string());
    }
    None
}

fn find_bytes(haystack: &[u8], needle: &[u8]) -> bool {
    haystack.windows(needle.len()).any(|window| window == needle)
}

fn family_for_source(source: &str) -> Option<Family> {
    match canonical_format_hint(source).as_str() {
        "txt" | "md" | "html" | "docx" | "doc" | "odt" | "pdf" => Some(Family::Writer),
        "csv" | "xlsx" | "ods" | "xls" => Some(Family::Calc),
        "pptx" | "odp" | "ppt" => Some(Family::Impress),
        "svg" | "odg" => Some(Family::Draw),
        "latex" | "mathml" | "odf" => Some(Family::Math),
        "odb" => Some(Family::Base),
        _ => None,
    }
}

/// Convert a writer-format byte stream from `from` to `to`.
pub fn writer_convert_bytes(input: &[u8], from: &str, to: &str) -> Result<Vec<u8>> {
    let from = canonical_format_hint(from);
    let to = canonical_format_hint(to);
    let doc = lo_writer::load_bytes("document", input, &from)?;
    lo_writer::save_as(&doc, &to)
}

/// Convert a calc-format byte stream from `from` to `to`.
pub fn calc_convert_bytes(input: &[u8], from: &str, to: &str) -> Result<Vec<u8>> {
    let from = canonical_format_hint(from);
    let to = canonical_format_hint(to);
    let workbook = lo_calc::load_bytes("workbook", input, &from)?;
    lo_calc::save_as(&workbook, &to)
}

/// Convert an impress-format byte stream from `from` to `to`.
pub fn impress_convert_bytes(input: &[u8], from: &str, to: &str) -> Result<Vec<u8>> {
    let from = canonical_format_hint(from);
    let to = canonical_format_hint(to);
    let deck = lo_impress::load_bytes("presentation", input, &from)?;
    lo_impress::save_as(&deck, &to)
}

/// Convert a draw-format byte stream from `from` to `to`.
pub fn draw_convert_bytes(input: &[u8], from: &str, to: &str) -> Result<Vec<u8>> {
    let from = canonical_format_hint(from);
    let to = canonical_format_hint(to);
    let drawing = lo_draw::load_bytes("drawing", input, &from)?;
    lo_draw::save_as(&drawing, &to)
}

/// Convert a math-format byte stream from `from` to `to`.
///
/// The generic `lo_math::save_as` handles `mathml`/`svg`/`pdf`. The ODF
/// formula package (`.odf`) lives in `lo_odf` — we route it here so the
/// `convert --to odf` router stays consistent with every other family.
pub fn math_convert_bytes(input: &[u8], from: &str, to: &str) -> Result<Vec<u8>> {
    let from = canonical_format_hint(from);
    let to = canonical_format_hint(to);
    let document = lo_math::load_bytes("formula", input, &from)?;
    if to == "odf" {
        return lo_odf::save_formula_document_bytes(&document);
    }
    lo_math::save_as(&document, &to)
}

/// Convert a base-format byte stream from `from` to `to`.
pub fn base_convert_bytes(input: &[u8], from: &str, to: &str) -> Result<Vec<u8>> {
    let from = canonical_format_hint(from);
    let to = canonical_format_hint(to);
    let database = lo_base::load_bytes("database", input, &from, None)?;
    lo_base::save_as(&database, &to)
}

/// Convert any supported office-format byte stream from `from` to `to`.
pub fn convert_bytes(input: &[u8], from: &str, to: &str) -> Result<Vec<u8>> {
    let from = canonical_format_hint(from);
    let to = canonical_format_hint(to);
    match family_for_source(&from) {
        Some(Family::Writer) => writer_convert_bytes(input, &from, &to),
        Some(Family::Calc) => calc_convert_bytes(input, &from, &to),
        Some(Family::Impress) => impress_convert_bytes(input, &from, &to),
        Some(Family::Draw) => draw_convert_bytes(input, &from, &to),
        Some(Family::Math) => math_convert_bytes(input, &from, &to),
        Some(Family::Base) => base_convert_bytes(input, &from, &to),
        None => Err(LoError::Unsupported(format!(
            "generic conversion source format not supported: {from}"
        ))),
    }
}

/// Infer the source format from `path` and dispatch to [`convert_bytes`].
pub fn convert_path_bytes(path: &str, input: &[u8], to: &str) -> Result<Vec<u8>> {
    let from = sniff_format_from_path(path).ok_or_else(|| {
        LoError::InvalidInput(format!("could not infer input format from path: {path}"))
    })?;
    convert_bytes(input, &from, to)
}

/// Infer the source format from the byte payload itself and dispatch to
/// [`convert_bytes`].
pub fn convert_bytes_auto(input: &[u8], to: &str) -> Result<Vec<u8>> {
    let from = sniff_format_from_bytes(input).ok_or_else(|| {
        LoError::InvalidInput("could not infer input format from byte stream".to_string())
    })?;
    convert_bytes(input, &from, to)
}

// ---- Writer shortcuts -----------------------------------------------------

pub fn docx_to_html_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "docx", "html")
}
pub fn docx_to_txt_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "docx", "txt")
}
pub fn pdf_to_txt_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "pdf", "txt")
}
pub fn pdf_to_md_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "pdf", "md")
}
pub fn pdf_to_html_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "pdf", "html")
}
pub fn docx_to_odt_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "docx", "odt")
}
pub fn odt_to_pdf_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "odt", "pdf")
}
pub fn odt_to_docx_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "odt", "docx")
}
pub fn odt_to_html_bytes(input: &[u8]) -> Result<Vec<u8>> {
    writer_convert_bytes(input, "odt", "html")
}

// ---- Calc shortcuts -------------------------------------------------------

pub fn xlsx_to_pdf_bytes(input: &[u8]) -> Result<Vec<u8>> {
    calc_convert_bytes(input, "xlsx", "pdf")
}
pub fn xlsx_to_html_bytes(input: &[u8]) -> Result<Vec<u8>> {
    calc_convert_bytes(input, "xlsx", "html")
}
pub fn xlsx_to_csv_bytes(input: &[u8]) -> Result<Vec<u8>> {
    calc_convert_bytes(input, "xlsx", "csv")
}
pub fn xlsx_to_ods_bytes(input: &[u8]) -> Result<Vec<u8>> {
    calc_convert_bytes(input, "xlsx", "ods")
}
pub fn ods_to_pdf_bytes(input: &[u8]) -> Result<Vec<u8>> {
    calc_convert_bytes(input, "ods", "pdf")
}
pub fn ods_to_xlsx_bytes(input: &[u8]) -> Result<Vec<u8>> {
    calc_convert_bytes(input, "ods", "xlsx")
}
pub fn ods_to_csv_bytes(input: &[u8]) -> Result<Vec<u8>> {
    calc_convert_bytes(input, "ods", "csv")
}

// ---- Impress shortcuts ----------------------------------------------------

pub fn pptx_to_html_bytes(input: &[u8]) -> Result<Vec<u8>> {
    impress_convert_bytes(input, "pptx", "html")
}
pub fn pptx_to_svg_bytes(input: &[u8]) -> Result<Vec<u8>> {
    impress_convert_bytes(input, "pptx", "svg")
}
pub fn pptx_to_odp_bytes(input: &[u8]) -> Result<Vec<u8>> {
    impress_convert_bytes(input, "pptx", "odp")
}
pub fn odp_to_pdf_bytes(input: &[u8]) -> Result<Vec<u8>> {
    impress_convert_bytes(input, "odp", "pdf")
}
pub fn odp_to_pptx_bytes(input: &[u8]) -> Result<Vec<u8>> {
    impress_convert_bytes(input, "odp", "pptx")
}

// ---------------------------------------------------------------------------
// XLSX recalc and report generation
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct RecalcErrorBucket {
    pub count: usize,
    pub locations: Vec<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct RecalcCheckReport {
    pub status: String,
    pub total_errors: usize,
    pub total_formulas: usize,
    pub error_summary: BTreeMap<String, RecalcErrorBucket>,
}

impl RecalcCheckReport {
    pub fn to_json(&self) -> String {
        fn esc(input: &str) -> String {
            input
                .replace('\\', "\\\\")
                .replace('"', "\\\"")
                .replace('\n', "\\n")
        }
        let mut json = String::new();
        json.push('{');
        json.push_str(&format!("\"status\":\"{}\"", esc(&self.status)));
        json.push_str(&format!(",\"total_errors\":{}", self.total_errors));
        json.push_str(&format!(",\"total_formulas\":{}", self.total_formulas));
        json.push_str(",\"error_summary\":{");
        let mut first_kind = true;
        for (kind, bucket) in &self.error_summary {
            if !first_kind {
                json.push(',');
            }
            first_kind = false;
            json.push_str(&format!("\"{}\":{{\"count\":{},\"locations\":[", esc(kind), bucket.count));
            for (index, location) in bucket.locations.iter().enumerate() {
                if index > 0 {
                    json.push(',');
                }
                json.push_str(&format!("\"{}\"", esc(location)));
            }
            json.push_str("]}");
        }
        json.push_str("}}");
        json
    }

    fn record_error(&mut self, kind: String, location: String) {
        self.total_errors += 1;
        let bucket = self.error_summary.entry(kind).or_default();
        bucket.count += 1;
        bucket.locations.push(location);
    }
}

/// Re-evaluate every formula in an XLSX workbook and rewrite the cached
/// `<v>` values inside the existing sheet XML. The result is a fresh
/// XLSX byte stream with the same shape as the input, minus
/// `xl/calcChain.xml`.
pub fn xlsx_recalc_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    let zip = ZipArchive::new(bytes)?;
    let workbook = lo_calc::from_xlsx_bytes("workbook", bytes)?;
    let evaluator = WorkbookEvaluator::new(&workbook);
    let sheet_targets = parse_xlsx_sheet_targets(&zip)?;

    let mut entries: Vec<lo_zip::ZipEntry> = Vec::new();
    for entry_name in zip.entries() {
        let path = normalize_zip_path(entry_name);
        if path == "xl/calcChain.xml" {
            continue;
        }
        if path == "[Content_Types].xml" {
            let xml = zip.read_string(&path)?;
            let mut root = parse_xml_document(&xml)?;
            remove_content_type_override(&mut root, "/xl/calcChain.xml");
            entries.push(lo_zip::ZipEntry::new(path, serialize_xml_document(&root).into_bytes()));
            continue;
        }
        if path == "xl/_rels/workbook.xml.rels" {
            let xml = zip.read_string(&path)?;
            let mut root = parse_xml_document(&xml)?;
            remove_calc_chain_relationships(&mut root);
            entries.push(lo_zip::ZipEntry::new(path, serialize_xml_document(&root).into_bytes()));
            continue;
        }
        if path == "xl/workbook.xml" {
            let xml = zip.read_string(&path)?;
            let mut root = parse_xml_document(&xml)?;
            mark_workbook_recalculated(&mut root);
            entries.push(lo_zip::ZipEntry::new(path, serialize_xml_document(&root).into_bytes()));
            continue;
        }
        if let Some(sheet_index) = sheet_targets.iter().position(|(target, _)| target == &path) {
            let xml = zip.read_string(&path)?;
            let mut root = parse_xml_document(&xml)?;
            patch_xlsx_sheet_formula_cache(&mut root, &workbook, sheet_index, &evaluator)?;
            entries.push(lo_zip::ZipEntry::new(path, serialize_xml_document(&root).into_bytes()));
            continue;
        }
        entries.push(lo_zip::ZipEntry::new(path, zip.read(entry_name)?));
    }
    lo_zip::ooxml_package(&entries)
}

/// Produce a Clark-shaped JSON report for an existing XLSX workbook.
pub fn xlsx_recalc_check_json(bytes: &[u8]) -> Result<String> {
    Ok(xlsx_recalc_report(bytes)?.to_json())
}

/// Produce the structured recalc report used by
/// [`xlsx_recalc_check_json`].
pub fn xlsx_recalc_report(bytes: &[u8]) -> Result<RecalcCheckReport> {
    let zip = ZipArchive::new(bytes)?;
    let workbook = lo_calc::from_xlsx_bytes("workbook", bytes)?;
    let evaluator = WorkbookEvaluator::new(&workbook);
    let sheet_targets = parse_xlsx_sheet_targets(&zip)?;
    let mut report = RecalcCheckReport {
        status: "ok".to_string(),
        ..RecalcCheckReport::default()
    };

    for (sheet_index, (path, sheet_name)) in sheet_targets.iter().enumerate() {
        if !zip.contains(path) {
            continue;
        }
        let xml = zip.read_string(path)?;
        let root = parse_xml_document(&xml)?;
        walk_formula_cells(&root, sheet_name, sheet_index, &evaluator, &mut report)?;
    }

    if report.total_errors > 0 {
        report.status = "error".to_string();
    }
    Ok(report)
}

fn parse_xlsx_sheet_targets(zip: &ZipArchive) -> Result<Vec<(String, String)>> {
    let workbook_root = parse_xml_document(&zip.read_string("xl/workbook.xml")?)?;
    let rels = parse_relationships(zip, "xl/workbook.xml")?;
    let mut out = Vec::new();
    if let Some(sheets) = workbook_root.child("sheets") {
        for (index, sheet) in sheets.children_named("sheet").enumerate() {
            let name = sheet.attr("name").unwrap_or("Sheet").to_string();
            let target = sheet
                .attr("id")
                .or_else(|| sheet.attr("r:id"))
                .and_then(|id| rels.get(id))
                .cloned()
                .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", index + 1));
            out.push((normalize_zip_path(&target), name));
        }
    }
    Ok(out)
}

fn parse_relationships(zip: &ZipArchive, part: &str) -> Result<BTreeMap<String, String>> {
    let rels_path = rels_path_for(part);
    if !zip.contains(&rels_path) {
        return Ok(BTreeMap::new());
    }
    let root = parse_xml_document(&zip.read_string(&rels_path)?)?;
    let mut map = BTreeMap::new();
    for rel in root.children_named("Relationship") {
        if let (Some(id), Some(target)) = (rel.attr("Id"), rel.attr("Target")) {
            map.insert(id.to_string(), resolve_part_target(part, target));
        }
    }
    Ok(map)
}

fn remove_content_type_override(root: &mut XmlNode, part_name: &str) {
    root.items.retain(|item| match item {
        XmlItem::Node(node) if node.local_name() == "Override" => node.attr("PartName") != Some(part_name),
        _ => true,
    });
    sync_node_children(root);
}

fn remove_calc_chain_relationships(root: &mut XmlNode) {
    root.items.retain(|item| match item {
        XmlItem::Node(node) if node.local_name() == "Relationship" => {
            let target = node.attr("Target").unwrap_or("");
            let rel_type = node.attr("Type").unwrap_or("");
            !target.ends_with("calcChain.xml")
                && !rel_type.to_ascii_lowercase().contains("calcchain")
        }
        _ => true,
    });
    sync_node_children(root);
}

fn mark_workbook_recalculated(root: &mut XmlNode) {
    let mut found = false;
    for item in &mut root.items {
        if let XmlItem::Node(node) = item {
            if node.local_name() == "calcPr" {
                node.attributes.insert("calcCompleted".to_string(), "1".to_string());
                node.attributes.insert("fullCalcOnLoad".to_string(), "0".to_string());
                node.attributes.remove("calcMode");
                found = true;
            }
        }
    }
    if !found {
        let mut attrs = BTreeMap::new();
        attrs.insert("calcCompleted".to_string(), "1".to_string());
        attrs.insert("fullCalcOnLoad".to_string(), "0".to_string());
        root.items.push(XmlItem::Node(XmlNode {
            name: "calcPr".to_string(),
            attributes: attrs,
            children: Vec::new(),
            items: Vec::new(),
            text: String::new(),
        }));
    }
    sync_node_children(root);
}

fn patch_xlsx_sheet_formula_cache(
    root: &mut XmlNode,
    workbook: &Workbook,
    sheet_index: usize,
    evaluator: &WorkbookEvaluator<'_>,
) -> Result<()> {
    let Some(sheet_data) = child_mut(root, "sheetData") else {
        return Ok(());
    };
    let mut shared_formulas: BTreeMap<String, (CellAddr, String)> = BTreeMap::new();
    for row in &mut sheet_data.children {
        if row.local_name() != "row" {
            continue;
        }
        let row_number = row
            .attr("r")
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(1);
        for cell in &mut row.children {
            if cell.local_name() == "c" {
                patch_formula_cell(cell, row_number, workbook, sheet_index, evaluator, &mut shared_formulas)?;
            }
        }
        sync_node_items_from_children(row);
    }
    sync_node_items_from_children(sheet_data);
    sync_node_items_from_children(root);
    Ok(())
}

fn walk_formula_cells(
    root: &XmlNode,
    sheet_name: &str,
    sheet_index: usize,
    evaluator: &WorkbookEvaluator<'_>,
    report: &mut RecalcCheckReport,
) -> Result<()> {
    let Some(sheet_data) = root.child("sheetData") else {
        return Ok(());
    };
    let mut shared_formulas: BTreeMap<String, (CellAddr, String)> = BTreeMap::new();
    for row in &sheet_data.children {
        if row.local_name() != "row" {
            continue;
        }
        let row_number = row
            .attr("r")
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(1);
        for cell in &row.children {
            if cell.local_name() != "c" {
                continue;
            }
            let (row_1, col_1) = cell
                .attr("r")
                .and_then(parse_a1_cell_ref)
                .unwrap_or((row_number, 1));
            let addr = CellAddr::new(row_1.saturating_sub(1) as u32, col_1.saturating_sub(1) as u32);
            let Some(formula) = resolve_formula_for_cell(cell, addr, &mut shared_formulas) else {
                continue;
            };
            report.total_formulas += 1;
            let value = evaluator
                .evaluate_formula(sheet_index, &formula)
                .unwrap_or_else(|_| EvalValue::Error("#VALUE!".to_string()));
            if let EvalValue::Error(kind) = value {
                report.record_error(kind, format!("{}!{}", sheet_name, addr.to_a1()));
            }
        }
    }
    Ok(())
}

fn patch_formula_cell(
    cell: &mut XmlNode,
    fallback_row: usize,
    workbook: &Workbook,
    sheet_index: usize,
    evaluator: &WorkbookEvaluator<'_>,
    shared_formulas: &mut BTreeMap<String, (CellAddr, String)>,
) -> Result<()> {
    let (row_1, col_1) = cell
        .attr("r")
        .and_then(parse_a1_cell_ref)
        .unwrap_or((fallback_row, 1));
    let addr = CellAddr::new(row_1.saturating_sub(1) as u32, col_1.saturating_sub(1) as u32);
    let Some(formula) = resolve_formula_for_cell(cell, addr, shared_formulas) else {
        return Ok(());
    };
    if formula.trim().is_empty() {
        return Ok(());
    }
    let _ = workbook; // kept for signature symmetry and future named-range resolution.
    let value = evaluator
        .evaluate_formula(sheet_index, &formula)
        .unwrap_or_else(|_| EvalValue::Error("#VALUE!".to_string()));
    let mut new_items = Vec::new();
    for item in &cell.items {
        match item {
            XmlItem::Text(text) => new_items.push(XmlItem::Text(text.clone())),
            XmlItem::Node(node) if matches!(node.local_name(), "v" | "is") => {}
            XmlItem::Node(node) => new_items.push(XmlItem::Node(node.clone())),
        }
    }
    new_items.push(XmlItem::Node(make_value_node(&value)));
    cell.items = new_items;
    sync_node_children(cell);
    apply_formula_cache_type(cell, &value);
    Ok(())
}

fn resolve_formula_for_cell(
    cell: &XmlNode,
    addr: CellAddr,
    shared_formulas: &mut BTreeMap<String, (CellAddr, String)>,
) -> Option<String> {
    let mut formula_text = None;
    let mut formula_kind = None;
    let mut shared_index = None;
    for child in &cell.children {
        if child.local_name() == "f" {
            formula_text = Some(text_content(child));
            formula_kind = child.attr("t").map(str::to_string);
            shared_index = child.attr("si").map(str::to_string);
            break;
        }
    }
    let text = formula_text.unwrap_or_default();
    if !text.trim().is_empty() {
        if formula_kind.as_deref() == Some("shared") {
            if let Some(si) = shared_index.clone() {
                shared_formulas.insert(si, (addr, text.clone()));
            }
        }
        return Some(text);
    }
    if formula_kind.as_deref() == Some("shared") {
        if let Some(si) = shared_index {
            if let Some((base_addr, base_formula)) = shared_formulas.get(&si) {
                return Some(translate_shared_formula(base_formula, *base_addr, addr));
            }
        }
    }
    None
}

fn apply_formula_cache_type(cell: &mut XmlNode, value: &EvalValue) {
    match value {
        EvalValue::Number(_) | EvalValue::Blank => {
            cell.attributes.remove("t");
        }
        EvalValue::Text(_) => {
            cell.attributes.insert("t".to_string(), "str".to_string());
        }
        EvalValue::Bool(_) => {
            cell.attributes.insert("t".to_string(), "b".to_string());
        }
        EvalValue::Error(_) => {
            cell.attributes.insert("t".to_string(), "e".to_string());
        }
    }
}

fn make_value_node(value: &EvalValue) -> XmlNode {
    let text = match value {
        EvalValue::Blank => String::new(),
        EvalValue::Number(number) => {
            if number.fract() == 0.0 && number.is_finite() {
                format!("{}", *number as i64)
            } else {
                number.to_string()
            }
        }
        EvalValue::Text(text) => text.clone(),
        EvalValue::Bool(value) => {
            if *value { "1".to_string() } else { "0".to_string() }
        }
        EvalValue::Error(text) => text.clone(),
    };
    XmlNode {
        name: "v".to_string(),
        attributes: BTreeMap::new(),
        children: Vec::new(),
        items: if text.is_empty() { Vec::new() } else { vec![XmlItem::Text(text.clone())] },
        text,
    }
}

fn parse_a1_cell_ref(input: &str) -> Option<(usize, usize)> {
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in input.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return None;
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let row = digits.parse().ok()?;
    let mut col = 0usize;
    for ch in letters.chars() {
        col = col * 26 + ((ch.to_ascii_uppercase() as u8 - b'A' + 1) as usize);
    }
    Some((row, col))
}

fn text_content(node: &XmlNode) -> String {
    let mut out = String::new();
    if !node.text.is_empty() {
        out.push_str(&node.text);
    }
    for child in &node.children {
        out.push_str(&text_content(child));
    }
    out
}

// ---------------------------------------------------------------------------
// Accept all tracked changes
// ---------------------------------------------------------------------------

/// Walk every WordprocessingML part inside a DOCX, accept all common
/// tracked revisions, then re-emit the package.
///
/// This keeps inserted content (`w:ins`, `w:moveTo`), drops deleted
/// content (`w:del`, `w:moveFrom`, deleted rows/cells), strips
/// formatting-history `*Change` elements, removes `w:trackRevisions`
/// from settings, and prunes unreferenced comments from
/// `word/comments.xml`.
pub fn accept_all_tracked_changes_docx_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    let zip = ZipArchive::new(bytes)?;
    let mut xml_parts: Vec<(String, XmlNode)> = Vec::new();
    let mut passthrough: Vec<lo_zip::ZipEntry> = Vec::new();

    for entry_name in zip.entries() {
        let path = normalize_zip_path(entry_name);
        if is_wordprocessing_xml(&path) {
            let xml = zip.read_string(&path)?;
            let root = parse_xml_document(&xml)?;
            let accepted = accept_revision_root(&root, &path);
            xml_parts.push((path, accepted));
        } else {
            passthrough.push(lo_zip::ZipEntry::new(path, zip.read(entry_name)?));
        }
    }

    let mut live_comment_ids = BTreeSet::new();
    for (path, root) in &xml_parts {
        if !path.ends_with("comments.xml") {
            collect_comment_ids(root, &mut live_comment_ids);
        }
    }

    let mut entries: Vec<lo_zip::ZipEntry> =
        Vec::with_capacity(xml_parts.len() + passthrough.len());
    for (path, root) in xml_parts {
        let root = if path.ends_with("comments.xml") {
            filter_comment_part(&root, &live_comment_ids)
        } else {
            root
        };
        entries.push(lo_zip::ZipEntry::new(
            path,
            serialize_xml_document(&root).into_bytes(),
        ));
    }
    entries.extend(passthrough);
    lo_zip::ooxml_package(&entries)
}

/// Back-compat alias; prefer [`accept_all_tracked_changes_docx_bytes`].
#[deprecated(note = "use accept_all_tracked_changes_docx_bytes")]
pub fn accept_tracked_changes_docx_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    accept_all_tracked_changes_docx_bytes(bytes)
}

/// Back-compat alias; prefer [`xlsx_recalc_bytes`].
#[deprecated(note = "use xlsx_recalc_bytes")]
pub fn recalc_existing_xlsx_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    xlsx_recalc_bytes(bytes)
}

fn is_wordprocessing_xml(path: &str) -> bool {
    path.starts_with("word/")
        && path.ends_with(".xml")
        && !path.contains("_rels/")
        && !path.ends_with("fontTable.xml")
}

fn accept_revision_root(root: &XmlNode, path: &str) -> XmlNode {
    let items = accept_revision_items(&root.items);
    let mut node = rebuild_node(root, items, root.attributes.clone());
    if path.ends_with("settings.xml") {
        node.items.retain(
            |item| !matches!(item, XmlItem::Node(child) if child.local_name() == "trackRevisions"),
        );
        sync_node_children(&mut node);
    }
    node
}

fn accept_revision_items(items: &[XmlItem]) -> Vec<XmlItem> {
    let mut out = Vec::new();
    for item in items {
        match item {
            XmlItem::Text(text) => out.push(XmlItem::Text(text.clone())),
            XmlItem::Node(node) => out.extend(accept_revision_node(node)),
        }
    }
    out
}

fn accept_revision_node(node: &XmlNode) -> Vec<XmlItem> {
    let local = node.local_name();
    if matches!(
        local,
        "del"
            | "delText"
            | "delInstrText"
            | "cellDel"
            | "moveFrom"
            | "moveFromRangeStart"
            | "moveFromRangeEnd"
            | "moveToRangeStart"
            | "moveToRangeEnd"
            | "customXmlDelRangeStart"
            | "customXmlDelRangeEnd"
            | "customXmlMoveFromRangeStart"
            | "customXmlMoveFromRangeEnd"
            | "customXmlMoveToRangeStart"
            | "customXmlMoveToRangeEnd"
            | "trackRevisions"
    ) {
        return Vec::new();
    }
    if matches!(
        local,
        "ins"
            | "moveTo"
            | "customXmlInsRangeStart"
            | "customXmlInsRangeEnd"
            | "cellIns"
    ) {
        return accept_revision_items(&node.items);
    }
    if local.ends_with("Change") || local == "numberingChange" || local == "cellMerge" {
        return Vec::new();
    }
    if row_deleted(node) || cell_deleted(node) {
        return Vec::new();
    }
    let items = accept_revision_items(&node.items);
    vec![XmlItem::Node(rebuild_node(node, items, node.attributes.clone()))]
}

fn row_deleted(node: &XmlNode) -> bool {
    if node.local_name() != "tr" {
        return false;
    }
    node.child("trPr")
        .map(|trpr| {
            trpr.children
                .iter()
                .any(|child| matches!(child.local_name(), "del" | "cellDel" | "cellMerge"))
        })
        .unwrap_or(false)
}

fn cell_deleted(node: &XmlNode) -> bool {
    if node.local_name() != "tc" {
        return false;
    }
    node.child("tcPr")
        .map(|tcpr| {
            tcpr.children
                .iter()
                .any(|child| matches!(child.local_name(), "cellDel" | "del"))
        })
        .unwrap_or(false)
}

fn collect_comment_ids(node: &XmlNode, out: &mut BTreeSet<String>) {
    let local = node.local_name();
    if matches!(
        local,
        "commentRangeStart" | "commentRangeEnd" | "commentReference"
    ) {
        if let Some(id) = attribute_local(node, "id") {
            out.insert(id.to_string());
        }
    }
    for child in &node.children {
        collect_comment_ids(child, out);
    }
}

fn filter_comment_part(root: &XmlNode, live_comment_ids: &BTreeSet<String>) -> XmlNode {
    if root.local_name() != "comments" {
        return root.clone();
    }
    let items = root
        .items
        .iter()
        .filter_map(|item| match item {
            XmlItem::Text(text) => Some(XmlItem::Text(text.clone())),
            XmlItem::Node(node) => filter_comment_node(node, live_comment_ids).map(XmlItem::Node),
        })
        .collect();
    rebuild_node(root, items, root.attributes.clone())
}

fn filter_comment_node(node: &XmlNode, live_comment_ids: &BTreeSet<String>) -> Option<XmlNode> {
    if node.local_name() == "comment" {
        let keep = attribute_local(node, "id")
            .map(|id| live_comment_ids.contains(id))
            .unwrap_or(true);
        if !keep {
            return None;
        }
    }
    let items = node
        .items
        .iter()
        .filter_map(|item| match item {
            XmlItem::Text(text) => Some(XmlItem::Text(text.clone())),
            XmlItem::Node(child) => filter_comment_node(child, live_comment_ids).map(XmlItem::Node),
        })
        .collect();
    Some(rebuild_node(node, items, node.attributes.clone()))
}

fn attribute_local<'a>(node: &'a XmlNode, local_name: &str) -> Option<&'a str> {
    let suffix = format!(":{local_name}");
    node.attributes.iter().find_map(|(key, value)| {
        if key == local_name || key.ends_with(&suffix) {
            Some(value.as_str())
        } else {
            None
        }
    })
}

// ---------------------------------------------------------------------------
// Shared XmlNode mutation helpers
// ---------------------------------------------------------------------------

fn rebuild_node(
    template: &XmlNode,
    items: Vec<XmlItem>,
    attributes: BTreeMap<String, String>,
) -> XmlNode {
    let mut node = XmlNode {
        name: template.name.clone(),
        attributes,
        children: Vec::new(),
        items,
        text: String::new(),
    };
    sync_node_children(&mut node);
    node
}

fn sync_node_children(node: &mut XmlNode) {
    node.children = node
        .items
        .iter()
        .filter_map(|item| match item {
            XmlItem::Node(child) => Some(child.clone()),
            _ => None,
        })
        .collect();
    node.text = node
        .items
        .iter()
        .filter_map(|item| match item {
            XmlItem::Text(text) => Some(text.clone()),
            _ => None,
        })
        .collect::<Vec<_>>()
        .join("");
}

fn sync_node_items_from_children(node: &mut XmlNode) {
    let mut child_index = 0usize;
    let mut new_items = Vec::with_capacity(node.items.len().max(node.children.len()));
    for item in &node.items {
        match item {
            XmlItem::Text(text) => new_items.push(XmlItem::Text(text.clone())),
            XmlItem::Node(_) => {
                if let Some(updated) = node.children.get(child_index) {
                    new_items.push(XmlItem::Node(updated.clone()));
                    child_index += 1;
                }
            }
        }
    }
    while let Some(updated) = node.children.get(child_index) {
        new_items.push(XmlItem::Node(updated.clone()));
        child_index += 1;
    }
    node.items = new_items;
    sync_node_children(node);
}

fn child_mut<'a>(node: &'a mut XmlNode, name: &str) -> Option<&'a mut XmlNode> {
    node.children
        .iter_mut()
        .find(|child| child.local_name() == name || child.name == name)
}

#[allow(dead_code)]
fn _assert_send_sync() {
    fn assert<T: Send + Sync>() {}
    assert::<Result<Vec<u8>>>();
    let _ = LoError::Parse(String::new());
}


// ---- Markdown extraction --------------------------------------------------

/// Extract Markdown from an existing DOCX file using the native Writer importer.
pub fn docx_to_md_bytes(input: &[u8]) -> Result<Vec<u8>> {
    let doc = lo_writer::from_docx_bytes("document", input)?;
    Ok(lo_writer::to_markdown(&doc).into_bytes())
}

/// Extract Markdown from an existing PPTX file using the native Impress importer.
pub fn pptx_to_md_bytes(input: &[u8]) -> Result<Vec<u8>> {
    let deck = lo_impress::from_pptx_bytes("presentation", input)?;
    Ok(lo_impress::to_markdown(&deck).into_bytes())
}

/// Extract Markdown from an existing XLSX file using the native Calc importer.
pub fn xlsx_to_md_bytes(input: &[u8]) -> Result<Vec<u8>> {
    let workbook = lo_calc::from_xlsx_bytes("workbook", input)?;
    Ok(lo_calc::to_markdown(&workbook).into_bytes())
}

// ---- Direct raster output -------------------------------------------------

/// Rasterize a DOCX document directly to PNG pages at the requested DPI.
pub fn docx_to_png_pages(input: &[u8], dpi: u32) -> Result<Vec<Vec<u8>>> {
    let doc = lo_writer::from_docx_bytes("document", input)?;
    Ok(lo_writer::render_png_pages(&doc, dpi.max(72)))
}

/// Rasterize a DOCX document directly to JPEG pages at the requested DPI.
pub fn docx_to_jpeg_pages(input: &[u8], dpi: u32, quality: u8) -> Result<Vec<Vec<u8>>> {
    let doc = lo_writer::from_docx_bytes("document", input)?;
    Ok(lo_writer::render_jpeg_pages(&doc, dpi.max(72), quality.max(1)))
}

/// Rasterize a PPTX deck directly to PNG slide images at the requested DPI.
pub fn pptx_to_png_pages(input: &[u8], dpi: u32) -> Result<Vec<Vec<u8>>> {
    let deck = lo_impress::from_pptx_bytes("presentation", input)?;
    Ok(lo_impress::render_png_pages(&deck, dpi.max(72)))
}

/// Rasterize a PPTX deck directly to JPEG slide images at the requested DPI.
pub fn pptx_to_jpeg_pages(input: &[u8], dpi: u32, quality: u8) -> Result<Vec<Vec<u8>>> {
    let deck = lo_impress::from_pptx_bytes("presentation", input)?;
    Ok(lo_impress::render_jpeg_pages(&deck, dpi.max(72), quality.max(1)))
}
