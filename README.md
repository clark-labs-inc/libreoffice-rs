# libreoffice-rs

`libreoffice-rs` is a pure-Rust, std-only monorepo that builds a serious foundation for an office suite:

- Writer-like rich text documents with Markdown ingestion, **PDF text input**, and **TXT/HTML/SVG/PDF/ODT/DOCX** export
- Calc-like spreadsheets with a formula engine and **CSV/HTML/SVG/PDF/ODS/XLSX** export
- Impress-like slide decks with **HTML/SVG/PDF/ODP/PPTX** export, including native PPTX chart rendering
- Draw-like vector pages with **SVG/PDF/ODG** export
- Math-like TeX-style formula parsing with **MathML/SVG/PDF/ODF** export
- Base-like tabular data with `SELECT` / `WHERE` / `ORDER BY` / `LIMIT` queries and **HTML/SVG/PDF/ODB** export
- ODF and OOXML packaging on top of a native Rust ZIP writer
- A pure-Rust raster pipeline (PNG + JPEG, no external crates) plus direct
  DOCX/PPTX → PNG/JPEG page rendering and DOCX/PPTX/XLSX → Markdown extraction
- Hardened DOCX tracked-changes acceptance and a legacy binary `.doc` reader
- Byte-sniffing helpers (`sniff_format_from_bytes`, `convert_bytes_auto`) and
  generic `libreoffice_pure::convert_bytes` / `convert_path_bytes` routers,
  plus per-domain `writer_/calc_/impress_/draw_/math_/base_convert_bytes`
- A UNO-like service runtime with property maps, an event bus, and built-in services
- A LibreOfficeKit-like in-process runtime (`lo_lok`): document handles, command dispatch, callbacks, tile rendering
- A desktop application surface (`lo_app`): start center, templates, recent files, preferences, autosave + recovery, macro recording/replay, per-window HTML shells with menubar/toolbar/sidebar
- A command-line front-end with `office-demo`/`desktop-demo` end-to-end commands and a `soffice --headless --convert-to`-compatible conversion mode

See [`CHANGELOG.md`](CHANGELOG.md) for per-version details.

## Important status note

This repository is **not honest-to-goodness feature parity with LibreOffice**. LibreOffice is a massive decades-old office suite with millions of lines of code, deep import/export compatibility, UI toolkits, rendering, printing, macros, scripting, accessibility, collaboration, databases, and platform integration. This codebase provides a coherent Rust architecture and a meaningful amount of real functionality, but it does **not** fully replace LibreOffice.

That said, it is intentionally structured to be a practical starting point rather than a toy:

- every crate is pure Rust and uses only the standard library
- ODF packages are written with a custom ZIP implementation
- the spreadsheet engine parses and evaluates non-trivial formulas
- the CLI can create/export actual `.odt`, `.ods`, `.odp`, `.odg`, `.odf`, and `.odb`-style archives

## Benchmark Results vs Real LibreOffice

Tested against **LibreOffice 26.2.2** on macOS (Apple Silicon). All generated documents were validated by opening/converting them with the real LibreOffice CLI (`soffice --headless`).

### Speed Comparison

libreoffice-rs generates documents **15-43× faster** than LibreOffice processes them. Sample numbers from `tests/benchmark.sh`:

| Operation | libreoffice-rs | LibreOffice | Speedup |
|-----------|---------------|-------------|---------|
| 10 paragraphs → ODT | 21ms | ~700ms | **~33×** |
| 100 paragraphs → ODT | 19ms | ~680ms | **~36×** |
| 1,000 paragraphs → ODT | 25ms | ~880ms | **~35×** |
| 10 rows → ODS | 22ms | ~660ms | **~30×** |
| 100 rows → ODS | 44ms | ~580ms | **~13×** |
| 1,000 rows → ODS | 31ms | ~680ms | **~22×** |
| 5,000 rows × 10 cols → ODS | 76ms | 1,158ms | **15×** |
| EN ODT → PDF (real LibreOffice render) | 231ms | 1,492ms | **6×** |
| ZH ODT → PDF (real LibreOffice render) | 18ms | 779ms | **43×** |
| ES ODT → PDF (real LibreOffice render) | 18ms | 664ms | **37×** |

### Importer Microbenchmarks

From `cargo run --release -p lo_cli --example bench` (Apple Silicon, release build):

| Importer | Iterations | Per-iter |
|----------|-----------:|---------:|
| `lo_zip::ZipArchive::new` (big.docx) | 200 | 10.2µs |
| DOCX → `TextDocument` (big.docx, 2k paras) | 50 | 5.39ms |
| DOCX → `TextDocument` (real.docx) | 5,000 | 32.7µs |
| XLSX → `Workbook` (sheet.xlsx) | 5,000 | 71.7µs |
| ODS → `Workbook` (sheet.ods) | 5,000 | 166.0µs |
| PPTX → `Presentation` (deck.pptx) | 5,000 | 361.8µs |
| ODP → `Presentation` (deck.odp) | 5,000 | 231.1µs |
| ODG → `Drawing` (draw.odg) | 5,000 | 97.0µs |
| MathML parser (formula.mathml) | 20,000 | 10.2µs |

### Multilingual Support

| Language | Generation | Content Preserved through soffice round-trip |
|----------|-----------|---|
| English (bold, italic, tables, links) | ~20ms | All text, formatting, tables |
| Chinese (中文, CJK characters) | 18ms | All characters, idioms, table data |
| Spanish (accents, ñ, ¿¡) | 20ms | All diacritics, special punctuation |

### Test Suite Totals

| Suite | Result |
|---|---|
| `cargo test --workspace` (unit tests) | **57 passed / 0 failed** |
| `tests/libreoffice_integration.sh` | **35 passed / 0 failed** |
| `tests/office_demo_integration.sh` (DOCX/XLSX/PPTX/ODT/ODS/ODP/ODG/ODF/ODB) | **40 passed / 0 failed** |
| `tests/desktop_demo_integration.sh` (lo_app desktop surface) | **38 passed / 0 failed** |
| `tests/benchmark.sh` (speed + accuracy + edge cases + multilingual) | **88 passed / 0 failed** |

### Per-Feature Accuracy (88/88 benchmark tests pass — 100%)

**Writer (ODT) Features:**
- Bold → `Strong` style
- Italic → `Emphasis` style
- Inline code → `Code` style
- Hyperlinks with URL and label
- Unordered lists
- Tables (arbitrary columns/rows)
- Headings (all 6 levels)
- Horizontal rules
- Page breaks

**Calc (ODS) Features:**
- Numeric, string, and boolean cell types
- Formulas with `of:=` ODF prefix notation
- All 16 formula functions: SUM, AVERAGE, MIN, MAX, COUNT, IF, AND, OR, NOT, ABS, ROUND, LEN, CONCAT, cell arithmetic, exponentiation

**Cross-Format Compatibility (validated by `soffice --convert-to`):**
| Conversion | Status |
|------------|--------|
| ODT → PDF | Pass |
| ODT → DOCX (round-trip) | Pass |
| DOCX → TXT (content preserved: heading, bold, link, table) | Pass |
| ODS → PDF | Pass |
| ODS → XLSX (round-trip) | Pass |
| XLSX → CSV (formulas re-evaluated by real LibreOffice) | Pass |
| ODP → PDF | Pass |
| ODP → PPTX (round-trip preserves slide titles + bullets) | Pass |
| ODG → PDF / SVG | Pass |
| ODF (Math) → PDF (rendered by real LibreOffice Math) | Pass |
| ODB → PDF | Pass |

**Edge Cases (all pass):**
- Empty documents
- XML-sensitive characters (`< > & " '`)
- Unicode: Arabic (RTL), CJK, math symbols, currency, diacritics
- 50,000-character single line
- 5,000 × 10 spreadsheet (5.8 MB ODS)
- Minimal single-column CSV
- Negative numbers, large numbers, quoted strings with embedded commas

### Running the Benchmarks

```bash
# Workspace unit tests
cargo test --workspace

# Integration tests against real LibreOffice
bash tests/libreoffice_integration.sh    # 35 assertions
bash tests/office_demo_integration.sh    # 40 assertions  (DOCX/XLSX/PPTX/ODT/ODS/ODP/ODG/ODF/ODB)
bash tests/desktop_demo_integration.sh   # 38 assertions  (lo_app desktop surface + autosave + macros)

# Full performance + accuracy benchmark
bash tests/benchmark.sh                  # 88 assertions

# Comprehensive 1:1 quality benchmark vs real LibreOffice (downloads 13 public
# DOCX/PPTX/XLSX fixtures from python-docx, scanny/python-pptx, PHPOffice and
# Calibre, runs both engines, scores PDF/CSV/Markdown text similarity, page
# counts, raster output, and wall-clock speed)
bash tests/quality_benchmark.sh
```

### Real-world Quality Benchmark (v0.4.1)

`tests/quality_benchmark.sh` downloads 13 public DOCX/PPTX/XLSX test files
from python-openxml, scanny/python-pptx, PHPOffice and Calibre, runs both
engines on each, then scores:

- **DOCX**: native `docx-to-pdf`, `docx-to-md`, and `docx-to-pngs` vs
  `soffice --convert-to pdf` / `txt`
- **PPTX**: native `pptx-to-pdf`, `pptx-to-md`, and `pptx-to-pngs` vs
  `soffice --convert-to pdf` / `txt`
- **XLSX**: native `convert --to csv`, `xlsx-to-md`, and `xlsx-recalc-check`
  vs `soffice --convert-to csv`

Results (LibreOffice 26.2.2 on Apple Silicon, 13/13 conversions succeeded):

| Metric | libreoffice-rs vs LibreOffice |
|---|---|
| Mean wall-clock per conversion | **18ms vs 902ms (~51× faster)** |
| PDF page-count agreement | **9/10 files exact match** |
| Mean PDF text Jaccard (DOCX, 7 files) | **99.1%** (5 at 100%, 1 at 99.2%, 1 at 94.4%) |
| Mean PDF text Jaccard (PPTX, 3 files) | **74.4%** (1 at 100%, 1 at 72%, 1 at 51%) |
| Mean CSV cell Jaccard (XLSX, 3 files) | **100.0%** (3/3) |
| Mean Markdown extraction Jaccard (13 files) | **90.9%** (DOCX mean 97.8% with 5/7 at 100%, XLSX 3/3 at 100%) |
| Native PNG raster | All 13 files rendered at 96 DPI |
| `xlsx-recalc-check` | All 3 XLSX files reported `status=ok` |

Per-file results: see `docs/quality_benchmark_results.txt` and the raw TSV
in `docs/quality_benchmark_results.tsv`.

Requires LibreOffice installed (`brew install --cask libreoffice` on macOS).

## Workspace layout

- `lo_core` — common models, XML/PDF/SVG/HTML helpers, styles, document data structures
- `lo_zip` — minimal ZIP reader/writer in pure Rust + ODF/OOXML packaging helpers
- `lo_odf` — ODF package serializers for every document type
- `lo_writer` — Writer-style editing, Markdown/plain-text ingestion, **TXT/HTML/SVG/PDF/ODT/DOCX** export
- `lo_calc` — spreadsheet formula parser/evaluator, CSV import/export, **HTML/SVG/PDF/ODS/XLSX** export
- `lo_impress` — presentation builders, **HTML/SVG/PDF/ODP/PPTX** export
- `lo_draw` — vector drawing builders, **SVG/PDF/ODG** export
- `lo_math` — TeX-style formula parser, **MathML/SVG/PDF** export (ODF formula via `lo_odf`)
- `lo_base` — typed table model, `SELECT` / `WHERE` / `ORDER BY` / `LIMIT` query execution, **HTML/SVG/PDF/ODB** export
- `lo_uno` — UNO-like service runtime: `ComponentContext`, factories, event bus, property maps, built-in `Echo`/`Info`/`TextTransformations` services
- `lo_lok` — LibreOfficeKit-like in-process runtime: `Office` handle, `DocumentHandle`, command dispatch, callbacks, SVG tile rendering
- `lo_app` — desktop application surface over `lo_lok`: windows, preferences, recent files, templates, clipboard, autosave + recovery, macro recording/replay, start-center HTML and per-window HTML shells
- `libreoffice-rs` — CLI binary with `office-demo` and `desktop-demo` end-to-end commands

## Examples

```bash
# Single-format conversions
cargo run -p libreoffice-pure -- writer new out.odt --title "Hello" --text "Hello from Rust"
cargo run -p libreoffice-pure -- writer markdown-to-odt notes.md notes.odt
cargo run -p libreoffice-pure -- writer convert notes.md notes.docx
cargo run -p libreoffice-pure -- writer convert notes.md notes.html
cargo run -p libreoffice-pure -- writer convert notes.md notes.pdf
cargo run -p libreoffice-pure -- calc csv-to-ods sheet.csv sheet.ods --sheet Data --has-header
cargo run -p libreoffice-pure -- calc convert sheet.csv sheet.xlsx --has-header
cargo run -p libreoffice-pure -- calc eval "=SUM(B2:B4)" --csv numbers.csv --has-header
cargo run -p libreoffice-pure -- impress demo deck.odp
cargo run -p libreoffice-pure -- draw demo diagram.odg
cargo run -p libreoffice-pure -- math latex-to-odf formula.txt formula.odf
cargo run -p libreoffice-pure -- base csv-to-odb table.csv People out.odb
cargo run -p libreoffice-pure -- package inspect out.odt

# soffice-compatible bytes→bytes conversion
cargo run -p libreoffice-pure -- --headless --convert-to pdf report.docx
cargo run -p libreoffice-pure -- --headless --convert-to pdf slide.pptx --outdir out
cargo run -p libreoffice-pure -- convert --to pdf report.docx
cargo run -p libreoffice-pure -- convert --from ods --to xlsx sheet.ods sheet.xlsx
cargo run -p libreoffice-pure -- --convert-to "pdf:writer_pdf_Export" notes.odt   # filter strings normalized

# Format auto-detection and recalc check
cargo run -p libreoffice-pure -- xlsx-recalc-check sheet.xlsx                      # JSON recalc report
# library API: sniff_format_from_bytes / convert_bytes_auto / xlsx_recalc_check_json

# Direct raster + Markdown extraction
cargo run -p libreoffice-pure -- docx-to-pngs report.docx pages_png --dpi 144
cargo run -p libreoffice-pure -- docx-to-jpegs report.docx pages_jpg --dpi 150 --quality 88
cargo run -p libreoffice-pure -- pptx-to-pngs slides.pptx slides_png --dpi 96
cargo run -p libreoffice-pure -- pptx-to-jpegs slides.pptx slides_jpg --dpi 150 --quality 85
cargo run -p libreoffice-pure -- docx-to-md report.docx report.md
cargo run -p libreoffice-pure -- pptx-to-md slides.pptx slides.md
cargo run -p libreoffice-pure -- xlsx-to-md sheet.xlsx sheet.md

# End-to-end demos
cargo run -p libreoffice-pure -- office-demo  ./demo_out          # writes 28 files (every kind × every format)
cargo run -p libreoffice-pure -- desktop-demo ./demo_profile      # full lo_app desktop surface (start center, autosave, macros, …)

# In-process runtime examples
cargo run -p lo_lok --example demo                                # open / command / save / tile via the LOK runtime
cargo run -p lo_app --example desktop_demo                        # open template / save / shell-render via DesktopApp
```

## Agent skill

An [Agent Skill](https://agentskills.io/specification) bundle for LLM-based
coding agents lives in [`skills/libreoffice-rs/`](skills/libreoffice-rs/).
Drop the directory into your agent's skills path (e.g. `~/.claude/skills/`)
and the agent will know how to drive `libreoffice-pure` for office-document
conversion, raster rendering, and Markdown extraction.

## Design goals

1. No C or C++ shims.
2. Native Rust models for all core document types.
3. Self-contained packaging and XML generation.
4. Easy-to-extend crate boundaries.
5. Honest status reporting.

## Feature status

See `STATUS.md` for a candid feature matrix.
See `CHANGELOG.md` for the per-version history.
