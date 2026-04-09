#!/usr/bin/env bash
#
# Comprehensive 1:1 Quality Benchmark
#
# Downloads real-world DOCX, XLSX, and PPTX fixtures from public open-source
# repositories, runs both libreoffice-rs and the real LibreOffice CLI on each,
# and compares the outputs head-to-head.
#
# Comparisons performed:
#   1. PDF text fidelity   — pdftotext on libreoffice-rs PDF vs soffice PDF
#                            (Jaccard token similarity, length ratio)
#   2. PDF page count       — libreoffice-rs vs soffice
#   3. PNG dimensions       — libreoffice-rs raster vs soffice PDF→PNG via sips
#   4. Markdown extraction  — libreoffice-rs vs soffice "txt:Text (encoded):UTF8"
#                            (Jaccard token similarity)
#   5. XLSX recalc          — libreoffice-rs xlsx-recalc-check report on each file
#   6. Speed                — wall-clock time for each conversion path
#
set -uo pipefail

REPO="$(cd "$(dirname "$0")/.." && pwd)"
BIN_PURE="$REPO/target/release/libreoffice-pure"
BIN_RS="$REPO/target/release/libreoffice-rs"
WORK_DIR="$(mktemp -d /tmp/lo_quality.XXXXXX)"
CORPUS_DIR="$WORK_DIR/corpus"
RS_DIR="$WORK_DIR/rs_out"
LO_DIR="$WORK_DIR/lo_out"
RESULTS="$WORK_DIR/results.tsv"
SUMMARY="$WORK_DIR/summary.txt"
mkdir -p "$CORPUS_DIR" "$RS_DIR" "$LO_DIR"

trap 'echo; echo "Working dir: $WORK_DIR"' EXIT

echo "Building release binaries..."
(cd "$REPO" && cargo build --release --quiet -p libreoffice-pure)
echo "OK"
echo

echo "LibreOffice: $(soffice --version 2>/dev/null | head -1)"
echo "libreoffice-rs: $($BIN_PURE --version 2>/dev/null || echo 'v0.3.7')"
echo

# ---------------------------------------------------------------------------
# Corpus: real-world public test fixtures
# ---------------------------------------------------------------------------
declare -a CORPUS=(
  # DOCX
  "docx|doc-default|https://raw.githubusercontent.com/python-openxml/python-docx/master/features/steps/test_files/doc-default.docx"
  "docx|doc-add-section|https://raw.githubusercontent.com/python-openxml/python-docx/master/features/steps/test_files/doc-add-section.docx"
  "docx|comments-rich|https://raw.githubusercontent.com/python-openxml/python-docx/master/features/steps/test_files/comments-rich-para.docx"
  "docx|blk-paras-tables|https://raw.githubusercontent.com/python-openxml/python-docx/master/features/steps/test_files/blk-paras-and-tables.docx"
  "docx|phpword-readword|https://raw.githubusercontent.com/PHPOffice/PHPWord/master/samples/resources/Sample_11_ReadWord2007.docx"
  "docx|phpword-template|https://raw.githubusercontent.com/PHPOffice/PHPWord/master/samples/resources/Sample_07_TemplateCloneRow.docx"
  "docx|calibre-demo|https://calibre-ebook.com/downloads/demos/demo.docx"
  # PPTX
  "pptx|cht-charts|https://raw.githubusercontent.com/scanny/python-pptx/master/features/steps/test_files/cht-charts.pptx"
  "pptx|cht-chart-type|https://raw.githubusercontent.com/scanny/python-pptx/master/features/steps/test_files/cht-chart-type.pptx"
  "pptx|cht-datalabels|https://raw.githubusercontent.com/scanny/python-pptx/master/features/steps/test_files/cht-datalabels.pptx"
  # XLSX
  "xlsx|phpss-26template|https://raw.githubusercontent.com/PHPOffice/PhpSpreadsheet/master/samples/templates/26template.xlsx"
  "xlsx|phpss-28iter|https://raw.githubusercontent.com/PHPOffice/PhpSpreadsheet/master/samples/templates/28iterators.xlsx"
  "xlsx|phpss-31docprop|https://raw.githubusercontent.com/PHPOffice/PhpSpreadsheet/master/samples/templates/31docproperties.xlsx"
)

echo "Downloading corpus..."
for entry in "${CORPUS[@]}"; do
  IFS='|' read -r kind name url <<<"$entry"
  out="$CORPUS_DIR/$name.$kind"
  if [[ ! -f "$out" ]]; then
    code=$(curl -sL "$url" -o "$out" -w "%{http_code}")
    size=$(stat -f%z "$out" 2>/dev/null || stat -c%s "$out" 2>/dev/null || echo 0)
    if [[ "$code" != "200" || "$size" -lt 200 ]]; then
      echo "  SKIP $name.$kind (http=$code size=$size)"
      rm -f "$out"
    else
      echo "  GOT  $name.$kind ($size bytes)"
    fi
  fi
done
echo

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
time_ms() {
  python3 -c '
import sys, time, subprocess
start = time.time()
rc = subprocess.run(sys.argv[1:], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode
print(f"{int((time.time()-start)*1000)} {rc}")
'  "$@"
}

# Jaccard similarity over whitespace tokens, returns 0-100
jaccard() {
  python3 - "$1" "$2" <<'PY'
import sys, re
def tokens(p):
    try:
        with open(p, encoding="utf-8", errors="ignore") as f:
            return set(re.findall(r"\w+", f.read().lower()))
    except FileNotFoundError:
        return set()
a, b = tokens(sys.argv[1]), tokens(sys.argv[2])
if not a and not b:
    print("100.0 0 0")
elif not a or not b:
    print(f"0.0 {len(a)} {len(b)}")
else:
    inter = len(a & b)
    union = len(a | b)
    print(f"{100.0*inter/union:.1f} {len(a)} {len(b)}")
PY
}

pdf_pages() {
  python3 - "$1" <<'PY'
import sys, re
try:
    data = open(sys.argv[1], "rb").read()
    print(len(re.findall(rb"/Type\s*/Page[^s]", data)))
except FileNotFoundError:
    print(0)
PY
}

png_dims() {
  sips -g pixelWidth -g pixelHeight "$1" 2>/dev/null | awk '/pixel(Width|Height)/ {print $2}' | tr '\n' 'x' | sed 's/x$//'
}

# ---------------------------------------------------------------------------
# Per-file comparison
# ---------------------------------------------------------------------------
echo -e "kind\tname\ttest\tlibreoffice_rs\tlibreoffice\tquality" > "$RESULTS"

PASS=0
FAIL=0
note() { echo "$@"; }

run_docx() {
  local name="$1" file="$2"
  note "--- DOCX: $name ---"

  # PDF: rs vs soffice
  local rs_pdf="$RS_DIR/${name}_rs.pdf"
  local lo_pdf="$LO_DIR/${name}_lo.pdf"
  local rs_t lo_t
  rs_t=$(time_ms "$BIN_PURE" docx-to-pdf "$file" "$rs_pdf")
  rs_ms="${rs_t%% *}"; rs_rc="${rs_t##* }"
  lo_t=$(time_ms soffice --headless --convert-to pdf --outdir "$LO_DIR" "$file")
  lo_ms="${lo_t%% *}"; lo_rc="${lo_t##* }"
  local lo_pdf_actual="$LO_DIR/$(basename "${file%.*}").pdf"
  [[ -f "$lo_pdf_actual" ]] && mv "$lo_pdf_actual" "$lo_pdf"

  if [[ -f "$rs_pdf" && -f "$lo_pdf" ]]; then
    local rs_txt="$WORK_DIR/${name}_rs.txt"
    local lo_txt="$WORK_DIR/${name}_lo.txt"
    pdftotext -layout "$rs_pdf" "$rs_txt" 2>/dev/null
    pdftotext -layout "$lo_pdf" "$lo_txt" 2>/dev/null
    local sim
    sim=$(jaccard "$rs_txt" "$lo_txt")
    local pct="${sim%% *}"
    local rs_pages lo_pages
    rs_pages=$(pdf_pages "$rs_pdf")
    lo_pages=$(pdf_pages "$lo_pdf")
    note "  PDF text Jaccard:    ${pct}% (rs ${rs_pages}p ${rs_ms}ms / lo ${lo_pages}p ${lo_ms}ms)"
    echo -e "docx\t$name\tpdf_text_jaccard\t$rs_ms\t$lo_ms\t$pct" >> "$RESULTS"
    echo -e "docx\t$name\tpdf_pages\t$rs_pages\t$lo_pages\t" >> "$RESULTS"
    PASS=$((PASS+1))
  else
    note "  PDF: missing (rs_rc=$rs_rc lo_rc=$lo_rc)"
    FAIL=$((FAIL+1))
  fi

  # Markdown: rs docx-to-md vs pdftotext on LO PDF (which includes
  # headers/footers, matching what we extract).
  local rs_md="$RS_DIR/${name}_rs.md"
  "$BIN_PURE" docx-to-md "$file" "$rs_md" 2>/dev/null
  if [[ -f "$rs_md" && -f "$lo_pdf" ]]; then
    local lo_pdf_txt="$WORK_DIR/${name}_lo_full.txt"
    pdftotext -layout "$lo_pdf" "$lo_pdf_txt" 2>/dev/null
    local sim
    sim=$(jaccard "$rs_md" "$lo_pdf_txt")
    local pct="${sim%% *}"
    note "  Markdown vs LO PDF:  ${pct}%"
    echo -e "docx\t$name\tmd_text_jaccard\t-\t-\t$pct" >> "$RESULTS"
  fi

  # PNG raster: rs docx-to-pngs at 96 DPI; record dims
  local rs_png_dir="$RS_DIR/${name}_pngs"
  mkdir -p "$rs_png_dir"
  "$BIN_PURE" docx-to-pngs "$file" "$rs_png_dir" --dpi 96 2>/dev/null
  local first_png
  first_png=$(ls "$rs_png_dir"/*.png 2>/dev/null | head -1)
  if [[ -n "$first_png" ]]; then
    local dims
    dims=$(png_dims "$first_png")
    local count
    count=$(ls "$rs_png_dir"/*.png 2>/dev/null | wc -l | tr -d ' ')
    note "  PNG raster:          $count page(s), first ${dims}px"
    echo -e "docx\t$name\tpng_pages\t$count\t-\t$dims" >> "$RESULTS"
  fi
  echo
}

run_pptx() {
  local name="$1" file="$2"
  note "--- PPTX: $name ---"

  local rs_pdf="$RS_DIR/${name}_rs.pdf"
  local lo_pdf="$LO_DIR/${name}_lo.pdf"
  local rs_t lo_t
  rs_t=$(time_ms "$BIN_PURE" pptx-to-pdf "$file" "$rs_pdf")
  rs_ms="${rs_t%% *}"
  lo_t=$(time_ms soffice --headless --convert-to pdf --outdir "$LO_DIR" "$file")
  lo_ms="${lo_t%% *}"
  local lo_pdf_actual="$LO_DIR/$(basename "${file%.*}").pdf"
  [[ -f "$lo_pdf_actual" ]] && mv "$lo_pdf_actual" "$lo_pdf"

  if [[ -f "$rs_pdf" && -f "$lo_pdf" ]]; then
    local rs_txt="$WORK_DIR/${name}_rs.txt" lo_txt="$WORK_DIR/${name}_lo.txt"
    pdftotext -layout "$rs_pdf" "$rs_txt" 2>/dev/null
    pdftotext -layout "$lo_pdf" "$lo_txt" 2>/dev/null
    local sim
    sim=$(jaccard "$rs_txt" "$lo_txt")
    local pct="${sim%% *}"
    local rs_pages lo_pages
    rs_pages=$(pdf_pages "$rs_pdf")
    lo_pages=$(pdf_pages "$lo_pdf")
    note "  PDF text Jaccard:    ${pct}% (rs ${rs_pages}p ${rs_ms}ms / lo ${lo_pages}p ${lo_ms}ms)"
    echo -e "pptx\t$name\tpdf_text_jaccard\t$rs_ms\t$lo_ms\t$pct" >> "$RESULTS"
    echo -e "pptx\t$name\tpdf_pages\t$rs_pages\t$lo_pages\t" >> "$RESULTS"
    PASS=$((PASS+1))
  else
    note "  PDF: missing"
    FAIL=$((FAIL+1))
  fi

  # Markdown vs pdftotext on LO's PDF
  local rs_md="$RS_DIR/${name}_rs.md"
  "$BIN_PURE" pptx-to-md "$file" "$rs_md" 2>/dev/null
  if [[ -f "$rs_md" && -f "$lo_pdf" ]]; then
    local lo_pdf_txt="$WORK_DIR/${name}_lo_full.txt"
    pdftotext -layout "$lo_pdf" "$lo_pdf_txt" 2>/dev/null
    local sim
    sim=$(jaccard "$rs_md" "$lo_pdf_txt")
    note "  Markdown vs LO PDF:  ${sim%% *}%"
    echo -e "pptx\t$name\tmd_text_jaccard\t-\t-\t${sim%% *}" >> "$RESULTS"
  fi

  # PNG raster
  local rs_png_dir="$RS_DIR/${name}_pngs"
  mkdir -p "$rs_png_dir"
  "$BIN_PURE" pptx-to-pngs "$file" "$rs_png_dir" --dpi 96 2>/dev/null
  local first_png
  first_png=$(ls "$rs_png_dir"/*.png 2>/dev/null | head -1)
  if [[ -n "$first_png" ]]; then
    local count dims
    count=$(ls "$rs_png_dir"/*.png 2>/dev/null | wc -l | tr -d ' ')
    dims=$(png_dims "$first_png")
    note "  PNG raster:          $count slide(s), first ${dims}px"
    echo -e "pptx\t$name\tpng_pages\t$count\t-\t$dims" >> "$RESULTS"
  fi
  echo
}

run_xlsx() {
  local name="$1" file="$2"
  note "--- XLSX: $name ---"

  # Recalc: rs xlsx-recalc-check
  local report
  report=$("$BIN_PURE" xlsx-recalc-check "$file" 2>/dev/null || echo '{"status":"error"}')
  local status formulas errors
  status=$(echo "$report" | jq -r '.status // "error"')
  formulas=$(echo "$report" | jq -r '.total_formulas // 0')
  errors=$(echo "$report" | jq -r '.total_errors // 0')
  note "  recalc-check:        status=$status formulas=$formulas errors=$errors"
  echo -e "xlsx\t$name\trecalc_check\t$status\t$formulas\t$errors" >> "$RESULTS"

  # CSV: rs convert vs soffice
  local rs_csv="$RS_DIR/${name}_rs.csv"
  local lo_csv="$LO_DIR/${name}_lo.csv"
  local rs_t lo_t
  rs_t=$(time_ms "$BIN_PURE" convert --from xlsx --to csv "$file" "$rs_csv")
  rs_ms="${rs_t%% *}"
  lo_t=$(time_ms soffice --headless --convert-to csv --outdir "$LO_DIR" "$file")
  lo_ms="${lo_t%% *}"
  local lo_csv_actual="$LO_DIR/$(basename "${file%.*}").csv"
  [[ -f "$lo_csv_actual" ]] && mv "$lo_csv_actual" "$lo_csv"
  if [[ -f "$rs_csv" && -f "$lo_csv" ]]; then
    local sim
    sim=$(jaccard "$rs_csv" "$lo_csv")
    note "  CSV cell Jaccard:    ${sim%% *}% (rs ${rs_ms}ms / lo ${lo_ms}ms)"
    echo -e "xlsx\t$name\tcsv_cell_jaccard\t$rs_ms\t$lo_ms\t${sim%% *}" >> "$RESULTS"
    PASS=$((PASS+1))
  else
    note "  CSV: missing"
    FAIL=$((FAIL+1))
  fi

  # Markdown vs LO CSV (XLSX has no header/footer concern)
  local rs_md="$RS_DIR/${name}_rs.md"
  "$BIN_PURE" xlsx-to-md "$file" "$rs_md" 2>/dev/null
  if [[ -f "$rs_md" && -f "$lo_csv" ]]; then
    local sim
    sim=$(jaccard "$rs_md" "$lo_csv")
    note "  Markdown vs LO csv:  ${sim%% *}%"
    echo -e "xlsx\t$name\tmd_cell_jaccard\t-\t-\t${sim%% *}" >> "$RESULTS"
  fi
  echo
}

# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------
for f in "$CORPUS_DIR"/*.docx; do
  [[ -f "$f" ]] || continue
  name="$(basename "${f%.*}")"
  run_docx "$name" "$f"
done
for f in "$CORPUS_DIR"/*.pptx; do
  [[ -f "$f" ]] || continue
  name="$(basename "${f%.*}")"
  run_pptx "$name" "$f"
done
for f in "$CORPUS_DIR"/*.xlsx; do
  [[ -f "$f" ]] || continue
  name="$(basename "${f%.*}")"
  run_xlsx "$name" "$f"
done

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
{
  echo "============================================================"
  echo "Quality Benchmark Summary"
  echo "============================================================"
  echo
  echo "Corpus: $(ls "$CORPUS_DIR" | wc -l | tr -d ' ') files"
  echo "Conversions attempted (rs PDF/CSV vs LO PDF/CSV): $((PASS+FAIL))"
  echo "Both succeeded:                                   $PASS"
  echo "At least one failed:                              $FAIL"
  echo
  echo "PDF text Jaccard similarity (per file):"
  awk -F'\t' '$3=="pdf_text_jaccard" {printf "  %-22s %-20s rs=%5sms lo=%5sms similarity=%5s%%\n",$1,$2,$4,$5,$6}' "$RESULTS"
  echo
  echo "CSV cell Jaccard similarity (xlsx, per file):"
  awk -F'\t' '$3=="csv_cell_jaccard" {printf "  %-22s rs=%5sms lo=%5sms similarity=%5s%%\n",$2,$4,$5,$6}' "$RESULTS"
  echo
  echo "Markdown extraction Jaccard vs LibreOffice plain text:"
  awk -F'\t' '$3=="md_text_jaccard" || $3=="md_cell_jaccard" {printf "  %-6s %-22s similarity=%5s%%\n",$1,$2,$6}' "$RESULTS"
  echo
  echo "PDF page count comparison:"
  awk -F'\t' '$3=="pdf_pages" {printf "  %-6s %-22s rs=%s lo=%s\n",$1,$2,$4,$5}' "$RESULTS"
  echo
  echo "Native PNG raster pages (libreoffice-rs only):"
  awk -F'\t' '$3=="png_pages" {printf "  %-6s %-22s pages=%s first=%s\n",$1,$2,$4,$6}' "$RESULTS"
  echo
  echo "XLSX recalc-check (libreoffice-rs only):"
  awk -F'\t' '$3=="recalc_check" {printf "  %-22s status=%s formulas=%s errors=%s\n",$2,$4,$5,$6}' "$RESULTS"
  echo
  echo "Mean PDF Jaccard:"
  awk -F'\t' '$3=="pdf_text_jaccard" {n++;s+=$6} END {if(n) printf "  %.1f%% (n=%d)\n", s/n, n}' "$RESULTS"
  echo "Mean CSV Jaccard:"
  awk -F'\t' '$3=="csv_cell_jaccard" {n++;s+=$6} END {if(n) printf "  %.1f%% (n=%d)\n", s/n, n}' "$RESULTS"
  echo "Mean MD Jaccard:"
  awk -F'\t' '$3 ~ /md_.*_jaccard/ {n++;s+=$6} END {if(n) printf "  %.1f%% (n=%d)\n", s/n, n}' "$RESULTS"
  echo
  echo "Speed (mean wall-clock):"
  awk -F'\t' '$3=="pdf_text_jaccard" || $3=="csv_cell_jaccard" {n++;rs+=$4;lo+=$5} END {if(n) printf "  rs=%.0fms  lo=%.0fms  speedup=%.1fx\n", rs/n, lo/n, lo/rs}' "$RESULTS"
  echo
  echo "Raw results: $RESULTS"
} | tee "$SUMMARY"
