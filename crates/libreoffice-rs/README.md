# libreoffice-rs

[![crates.io](https://img.shields.io/crates/v/libreoffice-pure.svg)](https://crates.io/crates/libreoffice-pure)
[![docs.rs](https://docs.rs/libreoffice-pure/badge.svg)](https://docs.rs/libreoffice-pure)

**Pure-Rust, std-only** library and CLI for reading, writing, converting, and rendering office documents — DOCX, XLSX, PPTX, ODT, ODS, ODP, ODG, ODF, ODB, PDF, HTML, SVG, Markdown, CSV, and plain text — with no C/C++ shims, no FFI, no LibreOffice install, and zero non-std dependencies.

## Quick start

```bash
cargo install libreoffice-pure

# Convert DOCX to PDF — no soffice, no Java, no system deps
libreoffice-pure --headless --convert-to pdf report.docx

# Spreadsheet to CSV
libreoffice-pure --headless --convert-to csv spreadsheet.xlsx

# Markdown to DOCX
libreoffice-pure convert --to docx notes.md notes.docx

# Extract text as Markdown
libreoffice-pure docx-to-md report.docx report.md

# Render pages as images
libreoffice-pure docx-to-pngs report.docx pages/ --dpi 144
```

The `--headless --convert-to` interface mirrors `soffice` — drop it into existing scripts as a faster, dependency-free replacement.

## Library API

```rust
use libreoffice_pure::{convert_bytes, sniff_format_from_bytes};

let docx = std::fs::read("report.docx").unwrap();
let pdf = convert_bytes(&docx, "docx", "pdf").unwrap();
std::fs::write("report.pdf", pdf).unwrap();
```

Per-domain crates (`lo_writer`, `lo_calc`, `lo_impress`, `lo_draw`, `lo_math`, `lo_base`) give fine-grained control over document creation and export.

## Performance

10–187× faster than real LibreOffice across the full N×M conversion matrix (mean ~116× across 63 head-to-head format pairs). See the [full README](https://github.com/clark-labs-inc/libreoffice-rs#benchmarks) for benchmark details.

## Status

This is **not** feature-parity with LibreOffice. It provides a coherent Rust architecture and meaningful real functionality, but does not fully replace LibreOffice. See [`STATUS.md`](https://github.com/clark-labs-inc/libreoffice-rs/blob/main/STATUS.md) for the candid feature matrix.

## License

MIT
