# Changelog

## 0.4.5 — Full evidence benchmark + `math → odf` convert router

### `math → odf` via the generic convert router

`libreoffice_pure::math_convert_bytes` now special-cases `to == "odf"` and
routes through a new `lo_odf::save_formula_document_bytes` (which shares
its layout with `save_formula_document` via a new
`package_document_bytes` helper in `lo_odf::common`). Previously
`convert --to odf` on a math input returned
`Unsupported("math format not supported: odf")` because `lo_math::save_as`
deliberately only handles `mathml` / `svg` / `pdf` — ODF formula packaging
lives in `lo_odf` to keep the dependency arrow pointing one way. The new
router closes the loop so the generic `convert` CLI, `convert_bytes`, and
`convert_bytes_auto` can all land a formula straight into an ODF archive
that real LibreOffice loads as a Math document.

Three new tests in `crates/libreoffice-rs/tests/math_convert.rs` cover
`latex → odf`, `mathml → odf`, and `odf → odf` round-trips. `cargo test
--workspace` reports **116 passed / 0 failed** (up from 113).

### Full Evidence Benchmark (`tests/full_evidence_benchmark.sh`)

A new end-to-end script exercises every supported feature, saves every
produced file to `benchmark_evidence/` for manual inspection, and scores
output against real LibreOffice.

- **Conversion matrix**: 193 attempts across Writer / Calc / Impress /
  Draw / Math / Base, covering every supported `(from, to)` pair per
  family. Result: **193 / 193 pass** after the `math → odf` fix (was
  190 / 193).
- **Pipeline commands**: `docx-to-{pdf,md,pngs,jpegs}`,
  `pptx-to-{pdf,md,pngs,jpegs}`, `xlsx-{recalc,recalc-check,to-md}`,
  `accept-changes`, `doc-to-docx`, `--headless --convert-to`,
  `office-demo`, `desktop-demo`, `package inspect`, formula `eval`,
  SQL `query`.
- **Real-world corpus**: US Census Bureau XLSX (state + county
  population 2020-2023), US CDC COVID-19 Deaths CSV (25 MB),
  python-openxml / PHPWord / PHPSpreadsheet / python-pptx / Calibre
  fixtures. Tracked-changes DOCX synthesized via Python zipfile
  (mirrors the unit-test fixtures) because no reliable public sample
  exists. Legacy `.doc` generated via `soffice --convert-to doc` so
  the round-trip test uses a real Word 97-2003 binary container.
- **Quality vs real LibreOffice**: DOCX PDF text Jaccard mean 99.0%
  (5 / 7 fixtures at 100 %), XLSX CSV mean 100 % across all real
  corpus files, PPTX text Jaccard varies from 42-100 % (charts are
  rendered differently).
- `.gitignore` now excludes `/benchmark_evidence` so the ~560 MB of
  outputs never land in git.

BLS XLSX endpoints return HTTP 403 (Akamai-blocked) regardless of the
User-Agent, so Census URLs are used for the US government XLSX corpus
and the failure is documented in the run log.

## 0.4.3 — Native PDF input

`lo_core::pdf` now ships a dependency-free PDF reader alongside the existing
writer helpers: `parse_pdf`, `extract_text_from_pdf`, and
`extract_pages_from_pdf` walk the cross-reference, expand object streams,
decode `FlateDecode` / `ASCIIHexDecode` / `ASCII85Decode` / `RunLengthDecode` /
`LZWDecode` filters, traverse the page tree (with inherited resources and
`MediaBox`), and replay content streams (`Tj`/`TJ`/`'`/`"` show operators,
`Td`/`TD`/`Tm`/`T*` text positioning, `cm`/`q`/`Q` graphics state, and `Do`
recursion into Form XObjects). Glyph runs are decoded through embedded
`ToUnicode` CMaps when present (with `bfchar`/`bfrange` plus surrogate-aware
UTF-16BE handling) and fall back to WinAnsi otherwise. A typed `PdfBuilder`
+ `PdfValue`/`PdfStream`/`PdfObjectId` model is also exported for
programmatic PDF construction.

`lo_writer::import` adds `from_pdf_bytes`, mapping each extracted page into
`Block::Paragraph` chunks with explicit `Block::PageBreak` separators, and
`load_bytes` now routes the `pdf` format hint there. The `libreoffice-pure`
top-level crate gains `pdf_to_txt_bytes` / `pdf_to_md_bytes` /
`pdf_to_html_bytes` wrappers, `sniff_format_from_bytes` recognises `%PDF-`
headers, and `family_for_source` routes `pdf` to Writer so generic
`convert_bytes` / `convert_bytes_auto` / `--convert-to` accept PDF input
end-to-end.

The CLI gains three commands — `pdf-to-txt`, `pdf-to-md`, `pdf-to-html` —
plus the existing `convert --to <fmt> input.pdf` path. Five new tests in
`crates/libreoffice-rs/tests/pdf_input.rs` cover the round-trip through
`pdf_to_*_bytes`, `convert_bytes`, `convert_bytes_auto`, and a hex-encoded
ToUnicode CMap, on top of two new unit tests in `lo_core::pdf` for the
parser and the `extract_text_from_pdf` happy path.

## 0.4.2 — README cleanup + Agent Skill bundle

- Trimmed `README.md`: per-version "**New in …**" bullet wall and the
  v0.3.8/v0.3.9 deep-dive sections moved to this `CHANGELOG.md`. Quality
  benchmark table reduced to current numbers only.
- New `skills/libreoffice-rs/SKILL.md` Agent Skill bundle following the
  [agentskills.io](https://agentskills.io/specification) spec, so LLM
  coding agents can drive `libreoffice-pure` for office-document
  conversion, raster rendering, Markdown extraction, and XLSX recalc
  checks.
- Path-dep versions in every crate `Cargo.toml` realigned with the
  workspace version (were stuck at `0.4.0`).

## 0.4.1 — Real PPTX chart rendering

`lo_impress::chart` parses `c:barChart` / `c:lineChart` / `c:areaChart` /
`c:pieChart` / `c:doughnutChart` / `c:scatterChart` / `c:radarChart` /
`c:bubbleChart` / `c:stockChart` (including combo charts) into a structured
`ChartSpec`, then draws axis grids, gridlines, value/category tick labels,
legend swatches, bar/column clusters, line/area polylines, scatter markers,
and pie/doughnut wedges natively into the PDF and PNG/JPEG renderers.
Date-formatted serial axes are skipped so Excel epoch numbers no longer leak
into extracted text. The Markdown extractor also now emits structured
`[chart: type]` blocks with axis titles, categories, and per-series value
lists.

## 0.4.0 — Hardened DOCX tracked-changes + legacy `.doc` reader

- `accept_all_tracked_changes_docx_bytes` now keeps inserted/moved-to content
  (`w:ins`, `w:moveTo`, `w:cellIns`), drops deleted content (`w:del`,
  `w:moveFrom`, deleted rows + cells via `w:cellDel`/`w:cellMerge`), strips
  formatting-history `*Change`/`numberingChange`/`cellMerge` records, removes
  `w:trackRevisions` from `word/settings.xml`, and **preserves live comment
  anchors while pruning unreferenced `<w:comment>` entries from
  `word/comments.xml`**.
- `lo_writer::legacy_doc` now falls back across `0Table`/`1Table`, scans for
  plausible CLX offsets when the canonical FIB location is stale, parses the
  piece table structurally, and normalises common WordPad / Word control
  codes into paragraphs, line breaks, and tabs. Compressed (8-bit) and
  Unicode (16-bit) pieces both round-trip.
- 8 new regression tests in `tests/tracked_changes_full.rs` and
  `lo_writer::legacy_doc::tests` cover insertion/deletion acceptance,
  live-vs-orphan comment pruning, formatting-only change records, deleted
  row/cell removal, move-from/move-to handling, CLX fallback, and
  control-code normalisation.

## 0.3.9 — Quality benchmark closer to 100%

Pushed the head-to-head quality benchmark to **5/7 DOCX files at 100% PDF
Jaccard, all 3 XLSX at 100% CSV + 100% Markdown, 1 of 3 PPTX charts at 100%
PDF Jaccard**, mean PDF Jaccard up from 82.7% → 91.7% and mean Markdown
Jaccard up from 80.4% → 90.9%.

- **DOCX import**: walk `<w:ins>`/`<w:del>`/`<w:moveTo>`/`<w:moveFrom>`/
  `<w:smartTag>`/`<w:customXml>`/`<w:sdt>` recursively. Recurse into nested
  `<w:tbl>` inside `<w:tc>`. Drop `<w:instrText>` field instructions from
  visible text. Parse `word/numbering.xml` and emit numbered-list markers
  (`1.`, `2.`, …) instead of bullets when the list's `numFmt` is `decimal`/
  `lowerRoman`/`lowerLetter`. Sequentially renumber `<w:footnoteReference>`/
  `<w:endnoteReference>` markers in document order. Drop the synthesized
  `dc:title` from rendered output to match `soffice --convert-to pdf`
  behaviour.
- **`lo_writer` PDF text rendering**: merge consecutive same-style runs into
  a single `Tj` operator and serialize tabs as four spaces. Transliterate
  Unicode typographic punctuation to ASCII equivalents in `pdf_escape`.
- **`lo_writer::render_table`**: bounds-check cell rendering when a row has
  fewer cells than the column count.
- **PPTX import**: synthesize "nice" axis tick labels from numeric value
  caches. Detect `<c:formatCode>` date formats and skip generating ticks for
  date axes. Honor chart-level `<c:dLbls><c:showVal>`.
- **`lo_impress` PDF backend**: render every chart token as its own
  positioned `Tj` operator so single-character tokens like axis tick labels
  are picked up as separate words by `pdftotext`.
- **`lo_calc::to_markdown`**: drop the synthesized workbook title and the
  `## SheetN` headers so XLSX Markdown matches LO CSV exactly.

## 0.3.8 — Close every gap surfaced by the v0.3.7 quality benchmark

- **UTF-8 fix in the XML parser** — `decode_entities` walked the input by
  byte rather than by char, mojibake-ing every multi-byte UTF-8 sequence.
  Foundational bug that touched DOCX, XLSX, PPTX, ODT, ODS, ODP, ODG and ODB
  import.
- **PPTX chart text** — `from_pptx_bytes` now walks `<p:graphicFrame>`,
  follows the chart relationship, and harvests every `<c:v>` / `<a:t>`
  string from the chart XML; nested `<p:grpSp>` are also walked recursively.
  Slide tables (`<a:tbl>`) are also extracted.
- **DOCX header / footer / footnote / endnote text** — `from_docx_bytes` now
  follows `word/_rels/document.xml.rels` and pulls plain text out of every
  header / footer / footnotes.xml / endnotes.xml part.
- **Real document titles** — DOCX, XLSX, and PPTX importers now read
  `docProps/core.xml` for the actual `dc:title`.
- **DOCX table layout panic** — `lo_writer::layout::render_table` no longer
  indexes past the end when a row has fewer cells than the table's column
  count.
- **Markdown noise** — `Block::PageBreak` now serializes to `---` instead of
  an HTML `<div class="page-break"></div>`, and single-sheet workbooks no
  longer emit a redundant `## Sheet1` header.

## 0.3.7 — Pure-Rust raster pipeline

- Pure-Rust raster pipeline in `lo_core::raster` (PNG + JPEG encoders, no
  external crates).
- Direct DOCX/PPTX → PNG/JPEG page rendering (`docx_to_png_pages`,
  `docx_to_jpeg_pages`, `pptx_to_png_pages`, `pptx_to_jpeg_pages`).
- Markdown extraction from DOCX/PPTX/XLSX (`docx_to_md_bytes`,
  `pptx_to_md_bytes`, `xlsx_to_md_bytes`).
- New `libreoffice-pure` CLI subcommands: `docx-to-md`, `pptx-to-md`,
  `xlsx-to-md`, `docx-to-pngs`, `docx-to-jpegs`, `pptx-to-pngs`,
  `pptx-to-jpegs`.

## 0.3.6 — Byte-sniffing and richer recalc

- Byte-sniffing (`sniff_format_from_bytes`, `convert_bytes_auto`).
- Workbook-aware XLSX recalc with cross-sheet & whole-row/column
  shared-formula translation.
- JSON recalc reports (`xlsx_recalc_check_json` / `xlsx_recalc_report`).
- Richer Writer PDF layout (colored headings, borders, table grids).
- Impress PDF renderer with slide chrome.
- Expanded `pdf_canvas` (RGB text/lines/rects/ellipses, line widths).
- Tracked-changes acceptance for comments, formatting-only changes, and
  deleted table rows.

## 0.3.5 — `soffice --headless --convert-to`-compatible CLI

- `office-demo`/`desktop-demo` end-to-end commands.
- `soffice --headless --convert-to`-compatible conversion mode.
- Generic `libreoffice_pure::convert_bytes` / `convert_path_bytes` helpers.
