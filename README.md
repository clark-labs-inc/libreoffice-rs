# libreoffice-rs — Pure Rust Office Document Toolkit (DOCX, XLSX, PPTX, ODF, PDF)

[![crates.io](https://img.shields.io/crates/v/libreoffice-pure.svg)](https://crates.io/crates/libreoffice-pure)
[![docs.rs](https://docs.rs/libreoffice-pure/badge.svg)](https://docs.rs/libreoffice-pure)
[![license: MIT](https://img.shields.io/crates/l/libreoffice-pure.svg)](LICENSE)

> **Pure-Rust, std-only** library and CLI for reading, writing, converting, and rendering office documents — **DOCX, XLSX, PPTX, ODT, ODS, ODP, ODG, ODF, ODB, PDF, HTML, SVG, Markdown, CSV, and plain text** — with **no C/C++ shims, no FFI, no LibreOffice install, and zero non-std dependencies**.

## Quick start

```bash
cargo install libreoffice-pure

# Convert DOCX to PDF — no soffice, no Java, no system deps
libreoffice-pure --headless --convert-to pdf report.docx

# Spreadsheet to CSV
libreoffice-pure --headless --convert-to csv spreadsheet.xlsx

# PPTX to PDF
libreoffice-pure --headless --convert-to pdf slides.pptx

# Markdown to DOCX
libreoffice-pure convert --to docx notes.md notes.docx

# Extract text from DOCX/PPTX/XLSX as Markdown
libreoffice-pure docx-to-md report.docx report.md
libreoffice-pure pptx-to-md slides.pptx slides.md
libreoffice-pure xlsx-to-md sheet.xlsx sheet.md

# Render pages as images
libreoffice-pure docx-to-pngs report.docx pages/ --dpi 144
libreoffice-pure pptx-to-jpegs slides.pptx slides/ --dpi 150 --quality 85
```

The `--headless --convert-to` interface mirrors `soffice` — drop it into existing scripts as a faster, dependency-free replacement.

## Why libreoffice-rs?

- **No LibreOffice install required.** Convert DOCX → PDF, XLSX → CSV, PPTX → Markdown, and more directly from Rust.
- **Pure Rust, std only.** Compiles with `cargo build` on a stock toolchain. Drop-in for WASM, serverless, sandboxes, and minimal containers.
- **Fast.** 10–187× faster than real LibreOffice across the full N×M conversion matrix (mean ~116× across 63 head-to-head pairs). See [benchmarks](#benchmarks).
- **Honest.** This is **not** a full replacement for LibreOffice. See [project status](#project-status).

## Project status

This is **not feature-parity with LibreOffice**. LibreOffice is a massive decades-old suite with deep compatibility, UI toolkits, printing, macros, scripting, and more. This codebase provides a coherent Rust architecture and meaningful real functionality — every crate is pure Rust/std-only, the spreadsheet engine evaluates non-trivial formulas, and the CLI creates actual ODF/OOXML archives — but it does **not** fully replace LibreOffice. See [`STATUS.md`](STATUS.md) for the candid feature matrix.

## Benchmarks

Tested against **LibreOffice 26.2.2** on macOS (Apple Silicon). All generated documents validated by `soffice --headless`.

### N×M Conversion Matrix (111 format pairs)

`tests/matrix_speed_comparison.py` runs every supported *(from → to)* pair through both engines on the same inputs:

| Family | Pairs | libreoffice-rs mean | LibreOffice mean | Mean speedup |
|---|---:|---:|---:|---:|
| writer  | 27 | 23 ms |   955 ms | **122×** |
| calc    | 20 |  7 ms |   671 ms | **111×** |
| impress | 10 | 12 ms |   766 ms |  **95×** |
| draw    |  6 |  5 ms |   752 ms | **139×** |

Head-to-head (n = 63): libreoffice-rs **14 ms** vs LibreOffice **816 ms** — mean **115.8×**. Realistic fixtures land in the 10–80× band; small synthetic inputs reach 100–180× because LibreOffice's ~700 ms startup dominates. libreoffice-rs also handles **35 pairs** whose formats `soffice --convert-to` cannot produce at all (Markdown, MathML, ODF-Math, ODB).

Full table: [`benchmark_evidence/matrix_speed_comparison.md`](benchmark_evidence/matrix_speed_comparison.md)

### Real-world Quality (13 public fixtures)

`tests/quality_benchmark.sh` downloads public DOCX/PPTX/XLSX files from python-openxml, python-pptx, PHPOffice and Calibre, then scores both engines:

- **51× faster** mean wall-clock (18 ms vs 902 ms)
- **99.1%** mean PDF text Jaccard on DOCX (5/7 files at 100%)
- **100%** CSV cell Jaccard on XLSX (3/3)
- **90.9%** Markdown extraction Jaccard across all 13 files
- PDF page count exact match on 9/10 files

Details: [`docs/quality_benchmark_results.txt`](docs/quality_benchmark_results.txt)

### Test suite

258 tests pass across 5 test suites (`cargo test --workspace`, integration, benchmarks). 88/88 benchmark assertions cover Writer features (bold, italic, headings, tables, links, lists), Calc features (16 formula functions, CSV round-trips), cross-format compatibility (11 conversion paths validated by real LibreOffice), multilingual support (English, Chinese, Spanish), and edge cases (empty docs, Unicode, 5K-row spreadsheets).

## CLI reference

```bash
# Format conversions (soffice-compatible interface)
libreoffice-pure --headless --convert-to pdf report.docx
libreoffice-pure convert --from ods --to xlsx sheet.ods sheet.xlsx

# Per-module commands
libreoffice-pure writer new out.odt --title "Hello" --text "Hello from Rust"
libreoffice-pure writer markdown-to-odt notes.md notes.odt
libreoffice-pure writer convert notes.md notes.pdf
libreoffice-pure calc csv-to-ods sheet.csv sheet.ods --sheet Data --has-header
libreoffice-pure calc eval "=SUM(B2:B4)" --csv numbers.csv --has-header
libreoffice-pure impress demo deck.odp
libreoffice-pure draw demo diagram.odg
libreoffice-pure math latex-to-odf formula.txt formula.odf
libreoffice-pure base csv-to-odb table.csv People out.odb

# Raster + Markdown extraction
libreoffice-pure docx-to-pngs report.docx pages/ --dpi 144
libreoffice-pure pptx-to-jpegs slides.pptx slides/ --dpi 150 --quality 85
libreoffice-pure docx-to-md report.docx report.md
libreoffice-pure xlsx-recalc-check sheet.xlsx

# End-to-end demos
libreoffice-pure office-demo ./demo_out
libreoffice-pure desktop-demo ./demo_profile
```

## Crate architecture

<details>
<summary>Workspace crates (14 crates, all pure Rust / std-only)</summary>

| Crate | Description |
|---|---|
| `lo_core` | Common models, XML/PDF/SVG/HTML helpers, styles |
| `lo_zip` | ZIP reader/writer + ODF/OOXML packaging |
| `lo_odf` | ODF package serializers |
| `lo_writer` | Writer: Markdown ingestion, TXT/HTML/SVG/PDF/ODT/DOCX export |
| `lo_calc` | Calc: formula engine, CSV/HTML/SVG/PDF/ODS/XLSX export |
| `lo_impress` | Impress: slide decks, HTML/SVG/PDF/ODP/PPTX export (incl. chart rendering) |
| `lo_draw` | Draw: vector pages, SVG/PDF/ODG export |
| `lo_math` | Math: TeX parser, MathML/SVG/PDF/ODF export |
| `lo_base` | Base: typed tables, SELECT/WHERE/ORDER BY queries, HTML/SVG/PDF/ODB export |
| `lo_uno` | UNO-like service runtime |
| `lo_lok` | LibreOfficeKit-like in-process runtime |
| `lo_app` | Desktop app surface (start center, autosave, macros, HTML shells) |
| `lo_cli` | CLI argument parsing |
| `libreoffice-pure` | CLI binary + `convert_bytes` / `convert_path_bytes` library API |

</details>

## FAQ

**Does libreoffice-rs require LibreOffice to be installed?**
No. It parses, writes, and converts DOCX/XLSX/PPTX/ODF/PDF files directly in pure Rust. LibreOffice only appears in benchmarks and integration tests as a reference.

**Can I use it as a library?**
Yes. Every crate is on [crates.io](https://crates.io/crates/libreoffice-pure). The `libreoffice_pure::convert_bytes` / `convert_path_bytes` API provides format-auto-detecting conversion. Per-domain crates (`lo_writer`, `lo_calc`, etc.) give fine-grained control.

**Does it work with WASM / serverless / containers?**
Yes. Pure Rust, std-only, no `build.rs` / FFI / system libraries.

## Agent skill

An [Agent Skill](https://agentskills.io/specification) bundle lives in [`skills/libreoffice-rs/`](skills/libreoffice-rs/). Drop it into your agent's skills path and it will know how to drive `libreoffice-pure` for document conversion, rendering, and extraction.

## Maintainer

Built and maintained by **Clark Labs Inc.**

See [`CHANGELOG.md`](CHANGELOG.md) for per-version details.
