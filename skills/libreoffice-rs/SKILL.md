---
name: libreoffice-rs
description: Convert, render, and extract text from office documents (DOCX, PPTX, XLSX, ODT, ODS, ODP, ODG, CSV, Markdown, PDF) using the pure-Rust libreoffice-rs CLI. Use whenever the user wants to turn one office format into another, render slides/documents to PDF or PNG/JPEG, extract Markdown from Word/PowerPoint/Excel files, recalc XLSX formulas, or run a fast headless soffice-compatible conversion without installing LibreOffice.
license: See repository LICENSE
metadata:
  homepage: https://github.com/stanvanrooy/libreoffice-rs
  version: "0.4.1"
compatibility: Requires a Rust toolchain (cargo) and a checkout of the libreoffice-rs repository. No C/C++ libraries needed; LibreOffice itself is NOT required.
---

# libreoffice-rs

`libreoffice-rs` is a pure-Rust office suite. The CLI binary is
`libreoffice-pure` and is invoked from inside a checkout with
`cargo run -p libreoffice-pure -- <subcommand>`. Add `--release` for
production-speed conversions on large inputs.

Use this skill when the user asks for any of:

- "convert this docx/pptx/xlsx/odt/ods/odp/odg/csv/md/pdf to X"
- "extract text/Markdown from this Word/PowerPoint/Excel file"
- "render these slides to PNG/JPEG/PDF"
- "headless soffice replacement", `--convert-to`, batch document conversion
- "recalculate / sanity-check XLSX formulas"
- "generate a Word/Excel/PowerPoint file from Markdown / CSV / a template"

The binary is fast (typically 15–50× faster than `soffice --headless`) and
needs no LibreOffice installation, so prefer it over shelling out to
`soffice` whenever the user already has the repo checked out.

## When to use vs not use

Use this skill when:
- The repo `libreoffice-rs` is checked out (or you can clone it).
- The user wants common office-format conversions, text/Markdown
  extraction, or page rasterization.

Do NOT use this skill when:
- The user needs full feature-parity LibreOffice behaviour (macros,
  complex layout edge cases, exotic charts). Fall back to real `soffice`.
- The user only wants to read/write plain text or JSON — use ordinary
  file tools.

## Quickstart

Always run from the repo root. Two equivalent invocation styles work:

```bash
# soffice-compatible style
cargo run -p libreoffice-pure -- --headless --convert-to pdf report.docx
cargo run -p libreoffice-pure -- --headless --convert-to pdf slide.pptx --outdir out

# explicit subcommand style
cargo run -p libreoffice-pure -- convert --to pdf report.docx
cargo run -p libreoffice-pure -- convert --from ods --to xlsx sheet.ods sheet.xlsx
```

For large or batch jobs add `--release`:

```bash
cargo run --release -p libreoffice-pure -- --headless --convert-to pdf big.docx
```

## Common recipes

### DOCX → PDF / TXT / Markdown / PNG / JPEG

```bash
cargo run -p libreoffice-pure -- --headless --convert-to pdf report.docx
cargo run -p libreoffice-pure -- docx-to-md   report.docx report.md
cargo run -p libreoffice-pure -- docx-to-pngs report.docx pages_png --dpi 144
cargo run -p libreoffice-pure -- docx-to-jpegs report.docx pages_jpg --dpi 150 --quality 88
```

### PPTX → PDF / Markdown / PNG / JPEG (with native chart rendering)

```bash
cargo run -p libreoffice-pure -- --headless --convert-to pdf slides.pptx
cargo run -p libreoffice-pure -- pptx-to-md   slides.pptx slides.md
cargo run -p libreoffice-pure -- pptx-to-pngs slides.pptx slides_png --dpi 96
cargo run -p libreoffice-pure -- pptx-to-jpegs slides.pptx slides_jpg --dpi 150 --quality 85
```

PPTX charts (`barChart`, `lineChart`, `areaChart`, `pieChart`,
`doughnutChart`, `scatterChart`, `radarChart`, `bubbleChart`, `stockChart`,
combo charts) are parsed and drawn natively into the PDF/raster output.

### XLSX / ODS → CSV / Markdown / round-trip

```bash
cargo run -p libreoffice-pure -- --headless --convert-to csv  sheet.xlsx
cargo run -p libreoffice-pure -- xlsx-to-md sheet.xlsx sheet.md
cargo run -p libreoffice-pure -- convert --from ods --to xlsx sheet.ods sheet.xlsx
```

Sanity-check XLSX formula recalculation against the file's cached values:

```bash
cargo run -p libreoffice-pure -- xlsx-recalc-check sheet.xlsx   # JSON report
```

### Generate documents from Markdown / CSV

```bash
cargo run -p libreoffice-pure -- writer markdown-to-odt notes.md notes.odt
cargo run -p libreoffice-pure -- writer convert         notes.md notes.docx
cargo run -p libreoffice-pure -- writer convert         notes.md notes.pdf
cargo run -p libreoffice-pure -- writer convert         notes.md notes.html
cargo run -p libreoffice-pure -- calc   csv-to-ods     sheet.csv sheet.ods --sheet Data --has-header
cargo run -p libreoffice-pure -- calc   convert        sheet.csv sheet.xlsx --has-header
cargo run -p libreoffice-pure -- calc   eval "=SUM(B2:B4)" --csv numbers.csv --has-header
```

### Other domains

```bash
cargo run -p libreoffice-pure -- impress  demo                  deck.odp
cargo run -p libreoffice-pure -- draw     demo                  diagram.odg
cargo run -p libreoffice-pure -- math     latex-to-odf          formula.txt formula.odf
cargo run -p libreoffice-pure -- base     csv-to-odb            people.csv People out.odb
cargo run -p libreoffice-pure -- package  inspect               out.odt
```

### Filter strings

`--convert-to "pdf:writer_pdf_Export"` and similar `soffice` filter
strings are accepted and normalized; the part after `:` is ignored when it
maps onto a known target.

## Library API (when embedding instead of shelling out)

If the user wants to call the library directly from Rust, point them at
these entry points:

- `libreoffice_pure::convert_bytes(input: &[u8], from: Format, to: Format)`
- `libreoffice_pure::convert_path_bytes(path, to: Format)`
- `libreoffice_pure::sniff_format_from_bytes(input: &[u8])`
- `libreoffice_pure::convert_bytes_auto(input: &[u8], to: Format)` —
  byte-sniffs the source format
- Per-domain helpers: `writer_convert_bytes`, `calc_convert_bytes`,
  `impress_convert_bytes`, `draw_convert_bytes`, `math_convert_bytes`,
  `base_convert_bytes`
- `lo_calc::xlsx_recalc_check_json` / `xlsx_recalc_report`
- Raster helpers in `lo_core::raster` and the high-level
  `docx_to_png_pages`, `docx_to_jpeg_pages`, `pptx_to_png_pages`,
  `pptx_to_jpeg_pages`
- `lo_writer::accept_all_tracked_changes_docx_bytes` for tracked-changes
  acceptance (insertions kept, deletions/move-from dropped, comment anchors
  preserved, orphan `<w:comment>` entries pruned)
- `lo_writer::legacy_doc` for the binary `.doc` reader

## Workspace layout (for code edits, not just CLI use)

- `lo_core` — shared models, XML/PDF/SVG/HTML helpers, raster pipeline
- `lo_zip` — pure-Rust ZIP + ODF/OOXML packaging
- `lo_odf` — ODF package serializers
- `lo_writer` — Writer (DOCX/ODT/HTML/PDF/SVG/TXT), legacy `.doc` reader,
  tracked-changes acceptance
- `lo_calc` — formulas, CSV/XLSX/ODS, recalc check
- `lo_impress` — slides, PPTX/ODP, native chart rendering
- `lo_draw` — vector drawings, ODG/SVG/PDF
- `lo_math` — TeX-style formulas, MathML/ODF/SVG/PDF
- `lo_base` — typed table model, SELECT/WHERE/ORDER BY/LIMIT, ODB
- `lo_uno` — UNO-like service runtime
- `lo_lok` — LibreOfficeKit-like in-process runtime
- `lo_app` — desktop application surface (windows, autosave, macros)
- `libreoffice-rs` — CLI binary (`libreoffice-pure`)

## Verifying your output

After a conversion, prefer one of:

- `cargo test --workspace` for code-level changes
- `bash tests/libreoffice_integration.sh` (35 assertions; needs `soffice`)
- `bash tests/office_demo_integration.sh` (40 assertions, no soffice)
- `bash tests/desktop_demo_integration.sh` (38 assertions)
- `bash tests/benchmark.sh` (88 assertions)
- `bash tests/quality_benchmark.sh` for the full 13-file 1:1 quality
  comparison vs real LibreOffice

For one-off output sanity-checking, run the produced file through
`pdftotext` or open it in a viewer and compare against the source.

## Edge cases & gotchas

- **Multi-byte UTF-8** is fully supported in DOCX/XLSX/PPTX/ODT/ODS/ODP/
  ODG/ODB import (post-0.3.8 fix). Treat any mojibake as a real bug, not
  expected behaviour.
- **DOCX tracked changes**: use
  `accept_all_tracked_changes_docx_bytes` if the user wants a clean copy.
  It keeps `w:ins`/`w:moveTo`/`w:cellIns`, drops `w:del`/`w:moveFrom`/
  deleted rows, strips formatting-history records, removes
  `w:trackRevisions`, and prunes orphan `<w:comment>` entries.
- **Numbered lists** in DOCX import respect `numFmt` (`decimal`,
  `lowerRoman`, `lowerLetter`) and emit `1.`, `2.`, … instead of bullets.
- **PPTX charts** with date-formatted serial axes intentionally skip tick
  labels (so Excel epoch numbers don't leak into extracted text).
- **XLSX Markdown** drops the synthesized workbook title and `## SheetN`
  headers so it matches what `soffice --convert-to csv` produces.
- **Filter strings** like `pdf:writer_pdf_Export` are normalized — pass
  them through unchanged when mirroring an existing soffice command line.
- **Performance**: always add `--release` for files larger than a few
  hundred KB; debug builds are dramatically slower.
- **Not feature-parity**: fall back to real `soffice` for macros, complex
  page layouts, or any feature not listed in `STATUS.md`.

## Reporting results

When you finish a conversion, summarise:

1. The exact command you ran.
2. Where the output landed (path).
3. Anything you noticed in the output (page count, warnings, recalc
   `status=ok`, Jaccard score if you ran a comparison).

If a conversion fails or produces obviously wrong output, capture the
stderr and check `STATUS.md` and `CHANGELOG.md` to see whether the failing
feature is known to be unsupported before filing a bug.
