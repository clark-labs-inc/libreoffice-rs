# Feature status

## Implemented in this repo
- Core office document models in Rust
- Writer-style document building plus TXT/HTML/SVG/PDF/ODT/DOCX export
- Calc-style spreadsheet building plus CSV/HTML/SVG/PDF/ODS/XLSX export
- Impress-style slide deck building plus HTML/SVG/PDF/ODP/PPTX export
- Draw-style vector page building plus SVG/PDF/ODG export
- Math parsing plus MathML/SVG/PDF import/export helpers
- Base-like tabular data plus HTML/SVG/PDF/ODB import/export helpers
- ODF archive generation for multiple document types
- ZIP writing and central-directory listing
- UNO-like service runtime and desktop/application surface crates
- `libreoffice-pure` high-level bytes→bytes helpers for:
  - DOCX → PDF
  - DOC → DOCX
  - PPTX → PDF
  - XLSX formula recalc-in-place
  - DOCX tracked-change acceptance
  - generic Writer/Calc/Impress/Draw/Math/Base conversion routing
- `libreoffice-pure` CLI surfaces for:
  - legacy one-command-per-conversion entry points
  - `convert --from X --to Y input [output]`
  - `--headless --convert-to pdf file.docx [--outdir dir]` soffice-style conversion
- Extension sniffing and filter-string normalization (for example `pdf:writer_pdf_Export` → `pdf`)

## Still not implemented to LibreOffice parity
- Native GUI parity with real LibreOffice
- Full ODF fidelity
- Full DOC/DOCX/XLS/XLSX/PPT/PPTX import/export parity
- Rendering/layout/pagination parity
- Printing/PDF parity with the real suite
- Macro runtime / full UNO compatibility
- Collaborative editing
- Accessibility stack parity
- Database driver ecosystem parity
- Grammar/spell checking parity
- Style/layout parity across formats
- Charting suite parity
- Formula compatibility parity
- VBA / StarBasic / Java / Python macro compatibility
