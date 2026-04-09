//! In-process LibreOfficeKit-like runtime.
//!
//! This crate is the bridge between the per-document crates (`lo_writer`,
//! `lo_calc`, `lo_impress`, `lo_draw`, `lo_math`, `lo_base`) and a calling
//! application that wants to do the things real LOK can do:
//!
//! - hold an [`Office`] handle that owns multiple documents
//! - open empty documents or load them from text/markdown/CSV bytes
//! - export each document via the per-crate `save_as` dispatchers
//! - render an SVG "tile" preview of any document
//! - dispatch UNO-style commands (`.uno:InsertText`, `.uno:SetCell`, …)
//! - publish [`KitEvent`] callbacks for opens, saves, commands and tiles
//! - share a [`lo_uno::ComponentContext`] across all documents so commands
//!   can themselves invoke services like `TextTransformations`
//!
//! It is intentionally small (no IPC, no threading model beyond `RwLock`),
//! but the surface area mirrors what a LOK consumer typically uses.

use std::collections::BTreeMap;
use std::fmt::{Display, Formatter};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, RwLock};

use lo_core::{
    units::Length, CellAddr, CellValue, DatabaseDocument, Drawing, FormulaDocument, LoError,
    Presentation, Result, Size, TextDocument, Workbook,
};
use lo_uno::{ComponentContext, PropertyMap, UnoValue};

pub type Callback = Arc<dyn Fn(&KitEvent) + Send + Sync>;

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct DocumentId(u64);

impl Display for DocumentId {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "doc-{}", self.0)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum DocumentKind {
    Writer,
    Calc,
    Impress,
    Draw,
    Math,
    Base,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct LoadOptions {
    /// Source format hint, e.g. `"md"`, `"txt"`, `"csv"`, `"latex"`.
    pub format: Option<String>,
    /// Table name to use when loading CSV into a Base document.
    pub table_name: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TileRequest {
    pub width: u32,
    pub height: u32,
}

impl Default for TileRequest {
    fn default() -> Self {
        Self {
            width: 1024,
            height: 768,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Tile {
    pub mime_type: String,
    pub bytes: Vec<u8>,
    pub width: u32,
    pub height: u32,
}

#[derive(Clone, Debug, PartialEq)]
pub enum KitEvent {
    DocumentOpened {
        id: DocumentId,
        kind: DocumentKind,
    },
    DocumentLoaded {
        id: DocumentId,
        kind: DocumentKind,
        format: String,
    },
    CommandExecuted {
        id: DocumentId,
        command: String,
    },
    DocumentSaved {
        id: DocumentId,
        format: String,
    },
    TileRendered {
        id: DocumentId,
        width: u32,
        height: u32,
    },
}

/// The actual document state held inside the office. Each variant wraps the
/// canonical typed model from `lo_core`.
#[derive(Clone, Debug, PartialEq)]
pub enum DocumentBackend {
    Writer(TextDocument),
    Calc(Workbook),
    Impress(Presentation),
    Draw(Drawing),
    Math(FormulaDocument),
    Base(DatabaseDocument),
}

impl DocumentBackend {
    pub fn kind(&self) -> DocumentKind {
        match self {
            Self::Writer(_) => DocumentKind::Writer,
            Self::Calc(_) => DocumentKind::Calc,
            Self::Impress(_) => DocumentKind::Impress,
            Self::Draw(_) => DocumentKind::Draw,
            Self::Math(_) => DocumentKind::Math,
            Self::Base(_) => DocumentKind::Base,
        }
    }

    pub fn title(&self) -> &str {
        match self {
            Self::Writer(d) => &d.meta.title,
            Self::Calc(w) => &w.meta.title,
            Self::Impress(p) => &p.meta.title,
            Self::Draw(d) => &d.meta.title,
            Self::Math(d) => &d.meta.title,
            Self::Base(d) => &d.meta.title,
        }
    }

    pub fn save_as(&self, format: &str) -> Result<Vec<u8>> {
        match self {
            Self::Writer(d) => lo_writer::save_as(d, format),
            Self::Calc(w) => lo_calc::save_as(w, format),
            Self::Impress(p) => lo_impress::save_as(p, format),
            Self::Draw(d) => lo_draw::save_as(d, format),
            Self::Math(d) => lo_math::save_as(d, format),
            Self::Base(d) => lo_base::save_as(d, format),
        }
    }

    /// Render an SVG preview tile of the document. The width/height in the
    /// request are interpreted as PDF/SVG points.
    pub fn render_tile(&self, request: &TileRequest) -> Result<Tile> {
        let size = Size::new(
            Length::pt(request.width as f32),
            Length::pt(request.height as f32),
        );
        let bytes = match self {
            Self::Writer(d) => lo_writer::render_svg(d, size).into_bytes(),
            Self::Calc(w) => lo_calc::render_svg(w, size).into_bytes(),
            Self::Impress(p) => lo_impress::render_svg(p).into_bytes(),
            Self::Draw(d) => lo_draw::render_svg(d).into_bytes(),
            Self::Math(d) => lo_math::render_svg(&d.root, size).into_bytes(),
            Self::Base(d) => lo_base::render_svg(d).into_bytes(),
        };
        Ok(Tile {
            mime_type: "image/svg+xml".to_string(),
            bytes,
            width: request.width,
            height: request.height,
        })
    }

    pub fn execute_command(
        &mut self,
        command: &str,
        arguments: &PropertyMap,
        ctx: &ComponentContext,
    ) -> Result<UnoValue> {
        match self {
            Self::Writer(d) => execute_writer_command(d, command, arguments, ctx),
            Self::Calc(w) => execute_calc_command(w, command, arguments),
            Self::Impress(p) => execute_impress_command(p, command, arguments),
            Self::Draw(d) => execute_draw_command(d, command, arguments),
            Self::Math(d) => execute_math_command(d, command, arguments),
            Self::Base(d) => execute_base_command(d, command, arguments),
        }
    }
}

struct DocumentState {
    backend: DocumentBackend,
}

struct OfficeInner {
    next_id: AtomicU64,
    ctx: ComponentContext,
    documents: RwLock<BTreeMap<DocumentId, Arc<RwLock<DocumentState>>>>,
    callbacks: RwLock<Vec<Callback>>,
}

/// Cloneable office handle. Cloning shares the same underlying document
/// registry, callback list and component context.
#[derive(Clone)]
pub struct Office {
    inner: Arc<OfficeInner>,
}

impl Default for Office {
    fn default() -> Self {
        Self::new()
    }
}

impl Office {
    pub fn new() -> Self {
        Self {
            inner: Arc::new(OfficeInner {
                next_id: AtomicU64::new(1),
                ctx: ComponentContext::new(),
                documents: RwLock::new(BTreeMap::new()),
                callbacks: RwLock::new(Vec::new()),
            }),
        }
    }

    pub fn context(&self) -> &ComponentContext {
        &self.inner.ctx
    }

    pub fn register_callback(&self, callback: Callback) {
        self.inner
            .callbacks
            .write()
            .expect("callback registry lock poisoned")
            .push(callback);
    }

    pub fn list_documents(&self) -> Vec<DocumentId> {
        self.inner
            .documents
            .read()
            .expect("document registry lock poisoned")
            .keys()
            .cloned()
            .collect()
    }

    /// Open an empty document of the given kind.
    pub fn open_empty(
        &self,
        kind: DocumentKind,
        title: impl Into<String>,
    ) -> Result<DocumentHandle> {
        let title = title.into();
        let backend = match kind {
            DocumentKind::Writer => DocumentBackend::Writer(TextDocument::new(title)),
            DocumentKind::Calc => DocumentBackend::Calc(Workbook::new(title)),
            DocumentKind::Impress => DocumentBackend::Impress(Presentation::new(title)),
            DocumentKind::Draw => DocumentBackend::Draw(Drawing::new(title)),
            DocumentKind::Math => {
                // Empty formula = a single empty group, which round-trips
                // through both MathML and the ODF formula serializer.
                let doc = lo_math::from_latex(title, "")?;
                DocumentBackend::Math(doc)
            }
            DocumentKind::Base => DocumentBackend::Base(DatabaseDocument::new(title)),
        };
        self.insert_document(backend, true, "empty")
    }

    /// Load a document from in-memory bytes. The bytes must be UTF-8 text.
    /// `format` controls how Writer/Math interpret the input.
    pub fn load_from_bytes(
        &self,
        kind: DocumentKind,
        title: impl Into<String>,
        bytes: &[u8],
        options: LoadOptions,
    ) -> Result<DocumentHandle> {
        let title = title.into();
        let format = options
            .format
            .clone()
            .unwrap_or_else(|| default_format_for_kind(kind).to_string());
        let text = std::str::from_utf8(bytes)
            .map_err(|err| LoError::Parse(format!("utf8 decode failed: {err}")))?
            .to_string();
        let backend = match kind {
            DocumentKind::Writer => {
                let doc = if format.eq_ignore_ascii_case("md")
                    || format.eq_ignore_ascii_case("markdown")
                {
                    lo_writer::from_markdown(title, &text)
                } else {
                    lo_writer::from_plain_text(title, &text)
                };
                DocumentBackend::Writer(doc)
            }
            DocumentKind::Calc => {
                let workbook = lo_calc::workbook_from_csv(title, "Sheet1", &text)?;
                DocumentBackend::Calc(workbook)
            }
            DocumentKind::Impress => {
                // No real importer; treat the loaded text as a single bullet
                // slide so the document at least carries the input content.
                let mut p = Presentation::new(title);
                p.slides.push(lo_core::Slide {
                    name: "Loaded".to_string(),
                    elements: vec![lo_core::SlideElement::TextBox(lo_core::TextBox {
                        frame: lo_core::Rect::new(
                            Length::mm(20.0),
                            Length::mm(30.0),
                            Length::mm(220.0),
                            Length::mm(120.0),
                        ),
                        text,
                        style: lo_core::TextBoxStyle::default(),
                    })],
                    notes: Vec::new(),
                chart_tokens: Vec::new(),
                });
                DocumentBackend::Impress(p)
            }
            DocumentKind::Draw => {
                let mut d = Drawing::new(title);
                d.pages[0]
                    .elements
                    .push(lo_core::DrawElement::TextBox(lo_core::TextBox {
                        frame: lo_core::Rect::new(
                            Length::mm(20.0),
                            Length::mm(30.0),
                            Length::mm(220.0),
                            Length::mm(120.0),
                        ),
                        text,
                        style: lo_core::TextBoxStyle::default(),
                    }));
                DocumentBackend::Draw(d)
            }
            DocumentKind::Math => DocumentBackend::Math(lo_math::from_latex(title, &text)?),
            DocumentKind::Base => {
                let table_name = options.table_name.unwrap_or_else(|| "data".to_string());
                let db = lo_base::database_from_csv(title, &table_name, &text)?;
                DocumentBackend::Base(db)
            }
        };
        self.insert_document(backend, false, &format)
    }

    fn insert_document(
        &self,
        backend: DocumentBackend,
        opened: bool,
        format: &str,
    ) -> Result<DocumentHandle> {
        let id = DocumentId(self.inner.next_id.fetch_add(1, Ordering::Relaxed));
        let kind = backend.kind();
        let state = Arc::new(RwLock::new(DocumentState { backend }));
        self.inner
            .documents
            .write()
            .expect("document registry lock poisoned")
            .insert(id.clone(), state);
        if opened {
            self.emit(KitEvent::DocumentOpened {
                id: id.clone(),
                kind,
            });
        } else {
            self.emit(KitEvent::DocumentLoaded {
                id: id.clone(),
                kind,
                format: format.to_string(),
            });
        }
        Ok(DocumentHandle {
            office: self.clone(),
            id,
        })
    }

    fn emit(&self, event: KitEvent) {
        let callbacks = self
            .inner
            .callbacks
            .read()
            .expect("callback registry lock poisoned")
            .clone();
        for callback in callbacks {
            callback(&event);
        }
    }

    fn state(&self, id: &DocumentId) -> Result<Arc<RwLock<DocumentState>>> {
        self.inner
            .documents
            .read()
            .expect("document registry lock poisoned")
            .get(id)
            .cloned()
            .ok_or_else(|| LoError::InvalidInput(format!("document not found: {id}")))
    }
}

#[derive(Clone)]
pub struct DocumentHandle {
    office: Office,
    id: DocumentId,
}

impl DocumentHandle {
    pub fn id(&self) -> &DocumentId {
        &self.id
    }

    pub fn save_as(&self, format: &str) -> Result<Vec<u8>> {
        let state = self.office.state(&self.id)?;
        let guard = state.read().expect("document lock poisoned");
        let bytes = guard.backend.save_as(format)?;
        drop(guard);
        self.office.emit(KitEvent::DocumentSaved {
            id: self.id.clone(),
            format: format.to_string(),
        });
        Ok(bytes)
    }

    pub fn render_tile(&self, request: TileRequest) -> Result<Tile> {
        let state = self.office.state(&self.id)?;
        let guard = state.read().expect("document lock poisoned");
        let tile = guard.backend.render_tile(&request)?;
        drop(guard);
        self.office.emit(KitEvent::TileRendered {
            id: self.id.clone(),
            width: tile.width,
            height: tile.height,
        });
        Ok(tile)
    }

    pub fn execute_command(&self, command: &str, arguments: &PropertyMap) -> Result<UnoValue> {
        let state = self.office.state(&self.id)?;
        let mut guard = state.write().expect("document lock poisoned");
        let value = guard
            .backend
            .execute_command(command, arguments, self.office.context())?;
        drop(guard);
        self.office.emit(KitEvent::CommandExecuted {
            id: self.id.clone(),
            command: command.to_string(),
        });
        Ok(value)
    }

    /// Return the text the document considers "selected" — this lets the
    /// application surface implement copy/paste without each crate having to
    /// know about clipboards. Writer returns its plain-text body, Math
    /// returns its source MathML, other kinds currently return `None`.
    pub fn selected_text(&self) -> Result<Option<String>> {
        let state = self.office.state(&self.id)?;
        let guard = state.read().expect("document lock poisoned");
        Ok(match &guard.backend {
            DocumentBackend::Writer(d) => Some(d.plain_text()),
            DocumentBackend::Math(d) => Some(lo_math::to_mathml_string(&d.root)),
            _ => None,
        })
    }
}

fn default_format_for_kind(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Writer => "txt",
        DocumentKind::Calc => "csv",
        DocumentKind::Impress => "txt",
        DocumentKind::Draw => "txt",
        DocumentKind::Math => "latex",
        DocumentKind::Base => "csv",
    }
}

fn get_string(arguments: &PropertyMap, key: &str) -> Option<String> {
    arguments
        .get(key)
        .and_then(UnoValue::as_str)
        .map(str::to_string)
}

fn get_i64(arguments: &PropertyMap, key: &str) -> Option<i64> {
    arguments.get(key).and_then(UnoValue::as_i64)
}

// ---- per-kind command handlers --------------------------------------------

fn execute_writer_command(
    doc: &mut TextDocument,
    command: &str,
    arguments: &PropertyMap,
    ctx: &ComponentContext,
) -> Result<UnoValue> {
    use lo_core::{Block, Inline, Paragraph, Table, TableCell, TableRow};
    match command {
        ".uno:InsertText" => {
            doc.push_paragraph(get_string(arguments, "text").unwrap_or_default());
            Ok(UnoValue::Bool(true))
        }
        ".uno:AppendHeading" => {
            let level = get_i64(arguments, "level").unwrap_or(1).clamp(1, 6) as u8;
            doc.push_heading(level, get_string(arguments, "text").unwrap_or_default());
            Ok(UnoValue::Bool(true))
        }
        ".uno:AppendTable" => {
            // `rows` is a multi-line string with `|`-separated cells per line
            // (matches the parity attempt's calling convention).
            let rows_text = get_string(arguments, "rows").unwrap_or_default();
            let rows: Vec<TableRow> = rows_text
                .lines()
                .map(|line| {
                    let cells = line
                        .split('|')
                        .map(|cell| TableCell {
                            paragraphs: vec![Paragraph {
                                spans: vec![Inline::Text(cell.trim().to_string())],
                                ..Paragraph::default()
                            }],
                        })
                        .collect();
                    TableRow { cells }
                })
                .collect();
            doc.body.push(Block::Table(Table {
                name: "Table1".to_string(),
                rows,
            }));
            Ok(UnoValue::Bool(true))
        }
        ".uno:SelectAll" => {
            // Selection state is implicit in this runtime (every command that
            // takes "selection" reads the whole body), so SelectAll is a no-op
            // success that exists so the menu/toolbar models can dispatch it.
            Ok(UnoValue::Bool(true))
        }
        ".uno:Bold" => {
            // "Bold the selection" → walk every inline span in the body and
            // wrap text runs in Inline::Bold.
            for block in &mut doc.body {
                if let Block::Paragraph(p) = block {
                    bold_inlines(&mut p.spans);
                } else if let Block::Heading(h) = block {
                    bold_inlines(&mut h.content.spans);
                }
            }
            Ok(UnoValue::Bool(true))
        }
        ".uno:GetSelectionText" => Ok(UnoValue::String(doc.plain_text())),
        ".uno:Uppercase" => {
            // Round-trips through the TextTransformations service so callers
            // can see the event bus deliver the result.
            let response = ctx.invoke(
                "com.libreoffice_rs.TextTransformations",
                "uppercase",
                &[UnoValue::String(doc.plain_text())],
            )?;
            // Replace the document body with one paragraph containing the
            // transformed text. Good enough for a "select-all + change case".
            if let Some(text) = response.as_str() {
                doc.body.clear();
                doc.push_paragraph(text);
            }
            Ok(response)
        }
        other => Err(LoError::Unsupported(format!("writer command {other}"))),
    }
}

fn bold_inlines(spans: &mut Vec<lo_core::Inline>) {
    use lo_core::Inline;
    for span in spans.iter_mut() {
        match span {
            Inline::Text(text) => *span = Inline::Bold(std::mem::take(text)),
            Inline::Italic(text) => *span = Inline::Bold(std::mem::take(text)),
            Inline::Code(_) | Inline::Bold(_) | Inline::Link { .. } | Inline::LineBreak => {}
        }
    }
}

fn execute_calc_command(
    workbook: &mut Workbook,
    command: &str,
    arguments: &PropertyMap,
) -> Result<UnoValue> {
    let sheet = workbook
        .sheet_mut(0)
        .ok_or_else(|| LoError::InvalidInput("workbook has no sheets".to_string()))?;
    match command {
        ".uno:SetCell" => {
            let row = get_i64(arguments, "row").unwrap_or(0).max(0) as u32;
            let col = get_i64(arguments, "col").unwrap_or(0).max(0) as u32;
            let raw = get_string(arguments, "value").unwrap_or_default();
            let value = if let Some(rest) = raw.strip_prefix('=') {
                CellValue::Formula(format!("={rest}"))
            } else if raw.eq_ignore_ascii_case("true") || raw.eq_ignore_ascii_case("false") {
                CellValue::Bool(raw.eq_ignore_ascii_case("true"))
            } else if let Ok(number) = raw.parse::<f64>() {
                CellValue::Number(number)
            } else if raw.is_empty() {
                CellValue::Empty
            } else {
                CellValue::Text(raw)
            };
            sheet.set(CellAddr::new(row, col), value);
            Ok(UnoValue::Bool(true))
        }
        ".uno:AppendRow" => {
            // `row` is a comma-separated string. Append it as a new row at
            // the bottom of the sheet, one cell per comma-separated entry.
            let row_text = get_string(arguments, "row").unwrap_or_default();
            let (max_row, _) = sheet.max_extent();
            let row_index = if sheet.cells.is_empty() {
                0
            } else {
                max_row + 1
            };
            for (col_idx, cell) in row_text.split(',').enumerate() {
                let trimmed = cell.trim();
                if trimmed.is_empty() {
                    continue;
                }
                let value = if let Ok(n) = trimmed.parse::<f64>() {
                    CellValue::Number(n)
                } else {
                    CellValue::Text(trimmed.to_string())
                };
                sheet.set(CellAddr::new(row_index, col_idx as u32), value);
            }
            Ok(UnoValue::Bool(true))
        }
        ".uno:EvaluateCell" => {
            let row = get_i64(arguments, "row").unwrap_or(0).max(0) as u32;
            let col = get_i64(arguments, "col").unwrap_or(0).max(0) as u32;
            // Borrow the sheet immutably for evaluation.
            let sheet_ref = workbook
                .sheet(0)
                .ok_or_else(|| LoError::InvalidInput("workbook has no sheets".to_string()))?;
            let cell = sheet_ref
                .get(CellAddr::new(row, col))
                .ok_or_else(|| LoError::InvalidInput("cell is empty".to_string()))?;
            match &cell.value {
                CellValue::Formula(formula) => {
                    let value = lo_calc::evaluate_formula(formula, sheet_ref)?;
                    Ok(UnoValue::String(format!("{value:?}")))
                }
                other => Ok(UnoValue::String(format!("{other:?}"))),
            }
        }
        other => Err(LoError::Unsupported(format!("calc command {other}"))),
    }
}

fn execute_impress_command(
    deck: &mut Presentation,
    command: &str,
    arguments: &PropertyMap,
) -> Result<UnoValue> {
    match command {
        ".uno:InsertSlide" => {
            let title = get_string(arguments, "title").unwrap_or_else(|| "Slide".to_string());
            deck.slides.push(lo_core::Slide {
                name: title,
                elements: Vec::new(),
                notes: Vec::new(),
                chart_tokens: Vec::new(),
            });
            Ok(UnoValue::Int(deck.slides.len() as i64))
        }
        ".uno:InsertTextBox" => {
            let slide_index = get_i64(arguments, "slide").unwrap_or(1).max(1) as usize - 1;
            while deck.slides.len() <= slide_index {
                deck.slides.push(lo_core::Slide::default());
            }
            deck.slides[slide_index]
                .elements
                .push(lo_core::SlideElement::TextBox(lo_core::TextBox {
                    frame: lo_core::Rect::new(
                        Length::mm(get_i64(arguments, "x").unwrap_or(20) as f32),
                        Length::mm(get_i64(arguments, "y").unwrap_or(20) as f32),
                        Length::mm(get_i64(arguments, "width").unwrap_or(220) as f32),
                        Length::mm(get_i64(arguments, "height").unwrap_or(40) as f32),
                    ),
                    text: get_string(arguments, "text").unwrap_or_default(),
                    style: lo_core::TextBoxStyle::default(),
                }));
            Ok(UnoValue::Bool(true))
        }
        ".uno:InsertBullets" => {
            // `items` is a `|`-separated list. We render bullets as a single
            // text box whose text is `• item\n• item\n…`, matching how the
            // Impress builder's bullet slide already presents bullets.
            let slide_index = get_i64(arguments, "slide").unwrap_or(1).max(1) as usize - 1;
            while deck.slides.len() <= slide_index {
                deck.slides.push(lo_core::Slide::default());
            }
            let body = get_string(arguments, "items")
                .unwrap_or_default()
                .split('|')
                .map(|item| format!("• {}", item.trim()))
                .filter(|line| line != "• ")
                .collect::<Vec<_>>()
                .join("\n");
            deck.slides[slide_index]
                .elements
                .push(lo_core::SlideElement::TextBox(lo_core::TextBox {
                    frame: lo_core::Rect::new(
                        Length::mm(20.0),
                        Length::mm(45.0),
                        Length::mm(220.0),
                        Length::mm(80.0),
                    ),
                    text: body,
                    style: lo_core::TextBoxStyle::default(),
                }));
            Ok(UnoValue::Bool(true))
        }
        ".uno:InsertShape" => {
            let slide_index = get_i64(arguments, "slide").unwrap_or(1).max(1) as usize - 1;
            while deck.slides.len() <= slide_index {
                deck.slides.push(lo_core::Slide::default());
            }
            let kind = match get_string(arguments, "kind")
                .unwrap_or_else(|| "rect".to_string())
                .to_ascii_lowercase()
                .as_str()
            {
                "ellipse" => lo_core::ShapeKind::Ellipse,
                "line" => lo_core::ShapeKind::Line,
                _ => lo_core::ShapeKind::Rectangle,
            };
            deck.slides[slide_index]
                .elements
                .push(lo_core::SlideElement::Shape(lo_core::Shape {
                    frame: lo_core::Rect::new(
                        Length::mm(get_i64(arguments, "x").unwrap_or(80) as f32),
                        Length::mm(get_i64(arguments, "y").unwrap_or(60) as f32),
                        Length::mm(get_i64(arguments, "width").unwrap_or(80) as f32),
                        Length::mm(get_i64(arguments, "height").unwrap_or(45) as f32),
                    ),
                    style: lo_core::ShapeStyle {
                        fill: "#d9eaf7".to_string(),
                        stroke: "#1f4e79".to_string(),
                        stroke_width_mm: 1,
                    },
                    kind,
                }));
            Ok(UnoValue::Bool(true))
        }
        other => Err(LoError::Unsupported(format!("impress command {other}"))),
    }
}

fn execute_draw_command(
    drawing: &mut Drawing,
    command: &str,
    arguments: &PropertyMap,
) -> Result<UnoValue> {
    match command {
        ".uno:InsertPage" => {
            let name = get_string(arguments, "name")
                .unwrap_or_else(|| format!("Page{}", drawing.pages.len() + 1));
            drawing.pages.push(lo_core::DrawPage {
                name,
                elements: Vec::new(),
            });
            Ok(UnoValue::Int(drawing.pages.len() as i64))
        }
        ".uno:InsertShape" => {
            if drawing.pages.is_empty() {
                drawing.pages.push(lo_core::DrawPage::default());
            }
            let page_index = get_i64(arguments, "page").unwrap_or(1).max(1) as usize - 1;
            while drawing.pages.len() <= page_index {
                drawing.pages.push(lo_core::DrawPage::default());
            }
            let kind = match get_string(arguments, "kind")
                .unwrap_or_else(|| "rect".to_string())
                .to_ascii_lowercase()
                .as_str()
            {
                "ellipse" => lo_core::ShapeKind::Ellipse,
                "line" => lo_core::ShapeKind::Line,
                _ => lo_core::ShapeKind::Rectangle,
            };
            drawing.pages[page_index]
                .elements
                .push(lo_core::DrawElement::Shape(lo_core::Shape {
                    frame: lo_core::Rect::new(
                        Length::mm(get_i64(arguments, "x").unwrap_or(20) as f32),
                        Length::mm(get_i64(arguments, "y").unwrap_or(40) as f32),
                        Length::mm(get_i64(arguments, "width").unwrap_or(60) as f32),
                        Length::mm(get_i64(arguments, "height").unwrap_or(40) as f32),
                    ),
                    style: lo_core::ShapeStyle {
                        fill: "#d9eaf7".to_string(),
                        stroke: "#1f4e79".to_string(),
                        stroke_width_mm: 1,
                    },
                    kind,
                }));
            // Optional label as a sibling text box.
            if let Some(text) = get_string(arguments, "text") {
                drawing.pages[page_index]
                    .elements
                    .push(lo_core::DrawElement::TextBox(lo_core::TextBox {
                        frame: lo_core::Rect::new(
                            Length::mm(get_i64(arguments, "x").unwrap_or(20) as f32 + 2.0),
                            Length::mm(get_i64(arguments, "y").unwrap_or(40) as f32 + 2.0),
                            Length::mm(get_i64(arguments, "width").unwrap_or(60) as f32 - 4.0),
                            Length::mm(10.0),
                        ),
                        text,
                        style: lo_core::TextBoxStyle::default(),
                    }));
            }
            Ok(UnoValue::Bool(true))
        }
        other => Err(LoError::Unsupported(format!("draw command {other}"))),
    }
}

fn execute_math_command(
    document: &mut FormulaDocument,
    command: &str,
    arguments: &PropertyMap,
) -> Result<UnoValue> {
    match command {
        ".uno:SetFormula" => {
            let formula = get_string(arguments, "formula").unwrap_or_default();
            *document = lo_math::from_latex(document.meta.title.clone(), &formula)?;
            Ok(UnoValue::Bool(true))
        }
        ".uno:ToMathML" => Ok(UnoValue::String(lo_math::to_mathml_string(&document.root))),
        other => Err(LoError::Unsupported(format!("math command {other}"))),
    }
}

fn execute_base_command(
    db: &mut DatabaseDocument,
    command: &str,
    arguments: &PropertyMap,
) -> Result<UnoValue> {
    use lo_core::{ColumnDef, ColumnType, DbValue, TableData};
    match command {
        ".uno:CreateTable" => {
            let name = get_string(arguments, "name").unwrap_or_else(|| "data".to_string());
            let columns: Vec<ColumnDef> = get_string(arguments, "columns")
                .unwrap_or_default()
                .split(',')
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty())
                .map(|name| ColumnDef {
                    name,
                    column_type: ColumnType::Text,
                })
                .collect();
            db.tables.push(TableData {
                name: name.clone(),
                columns,
                rows: Vec::new(),
            });
            Ok(UnoValue::String(name))
        }
        ".uno:InsertRow" => {
            let table_name = get_string(arguments, "table").unwrap_or_else(|| "data".to_string());
            let row_text = get_string(arguments, "row").unwrap_or_default();
            let table = db
                .tables
                .iter_mut()
                .find(|t| t.name == table_name)
                .ok_or_else(|| LoError::InvalidInput(format!("table not found: {table_name}")))?;
            let values: Vec<DbValue> = row_text
                .split(',')
                .map(|cell| {
                    let trimmed = cell.trim();
                    if trimmed.is_empty() {
                        DbValue::Null
                    } else if trimmed.eq_ignore_ascii_case("true") {
                        DbValue::Bool(true)
                    } else if trimmed.eq_ignore_ascii_case("false") {
                        DbValue::Bool(false)
                    } else if let Ok(v) = trimmed.parse::<i64>() {
                        DbValue::Integer(v)
                    } else if let Ok(v) = trimmed.parse::<f64>() {
                        DbValue::Float(v)
                    } else {
                        DbValue::Text(trimmed.to_string())
                    }
                })
                .collect();
            table.rows.push(values);
            Ok(UnoValue::Bool(true))
        }
        ".uno:ExecuteQuery" => {
            let sql = get_string(arguments, "sql").unwrap_or_default();
            let result = lo_base::execute_select(db, &sql)?;
            Ok(UnoValue::Int(result.rows.len() as i64))
        }
        other => Err(LoError::Unsupported(format!("base command {other}"))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;

    fn args(pairs: &[(&str, UnoValue)]) -> PropertyMap {
        pairs
            .iter()
            .map(|(k, v)| ((*k).to_string(), v.clone()))
            .collect()
    }

    #[test]
    fn writer_open_command_save() {
        let office = Office::new();
        let doc = office.open_empty(DocumentKind::Writer, "doc").unwrap();
        doc.execute_command(
            ".uno:InsertText",
            &args(&[("text", UnoValue::String("hello world".to_string()))]),
        )
        .unwrap();
        let pdf = doc.save_as("pdf").unwrap();
        assert!(pdf.starts_with(b"%PDF-1.4"));
    }

    #[test]
    fn calc_set_cell_and_evaluate() {
        let office = Office::new();
        let doc = office.open_empty(DocumentKind::Calc, "wb").unwrap();
        doc.execute_command(
            ".uno:SetCell",
            &args(&[
                ("row", UnoValue::Int(0)),
                ("col", UnoValue::Int(0)),
                ("value", UnoValue::String("2".to_string())),
            ]),
        )
        .unwrap();
        doc.execute_command(
            ".uno:SetCell",
            &args(&[
                ("row", UnoValue::Int(1)),
                ("col", UnoValue::Int(0)),
                ("value", UnoValue::String("3".to_string())),
            ]),
        )
        .unwrap();
        doc.execute_command(
            ".uno:SetCell",
            &args(&[
                ("row", UnoValue::Int(2)),
                ("col", UnoValue::Int(0)),
                ("value", UnoValue::String("=SUM(A1:A2)".to_string())),
            ]),
        )
        .unwrap();
        let result = doc
            .execute_command(
                ".uno:EvaluateCell",
                &args(&[("row", UnoValue::Int(2)), ("col", UnoValue::Int(0))]),
            )
            .unwrap();
        let s = result.as_str().unwrap_or_default().to_string();
        assert!(s.contains("5.0") || s.contains("Number(5"));
    }

    #[test]
    fn callbacks_fire_on_open_command_save() {
        let office = Office::new();
        let log = Arc::new(Mutex::new(Vec::<String>::new()));
        let log_for_callback = Arc::clone(&log);
        office.register_callback(Arc::new(move |event| {
            log_for_callback.lock().unwrap().push(format!("{event:?}"));
        }));
        let doc = office.open_empty(DocumentKind::Writer, "cb").unwrap();
        doc.execute_command(
            ".uno:InsertText",
            &args(&[("text", UnoValue::String("hi".to_string()))]),
        )
        .unwrap();
        doc.save_as("txt").unwrap();
        let entries = log.lock().unwrap();
        assert!(entries.iter().any(|m| m.contains("DocumentOpened")));
        assert!(entries.iter().any(|m| m.contains("CommandExecuted")));
        assert!(entries.iter().any(|m| m.contains("DocumentSaved")));
    }

    #[test]
    fn render_tile_returns_svg_bytes() {
        let office = Office::new();
        let doc = office.open_empty(DocumentKind::Writer, "tile").unwrap();
        doc.execute_command(
            ".uno:InsertText",
            &args(&[("text", UnoValue::String("tile me".to_string()))]),
        )
        .unwrap();
        let tile = doc.render_tile(TileRequest::default()).unwrap();
        assert_eq!(tile.mime_type, "image/svg+xml");
        let svg = std::str::from_utf8(&tile.bytes).unwrap();
        assert!(svg.starts_with("<svg"));
        assert!(svg.contains("tile me"));
    }
}
