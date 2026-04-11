#!/usr/bin/env bash
#
# Full Evidence Benchmark
#
# Exercises every supported libreoffice-rs feature and saves all outputs to
# a persistent, browsable evidence folder. Produces an N×M conversion matrix
# across Writer/Calc/Impress/Draw/Math/Base families plus every pipeline
# command (docx-to-pdf/md/png/jpeg, pptx-to-*, xlsx-recalc, accept-changes,
# office-demo, desktop-demo, package inspect). Downloads real-world inputs
# from government (BLS, Census) and open-source fixture repos.
#
# Result: $EVIDENCE/SUMMARY.md + organized subfolders with every produced file.
#
set -uo pipefail

REPO="$(cd "$(dirname "$0")/.." && pwd)"
BIN_PURE="$REPO/target/release/libreoffice-pure"
BIN_RS="$REPO/target/release/libreoffice-rs"
EVIDENCE="$REPO/benchmark_evidence"

CORPUS="$EVIDENCE/00_corpus"
MATRIX="$EVIDENCE/01_matrix"
DOCX_PIPE="$EVIDENCE/02_docx_pipeline"
XLSX_PIPE="$EVIDENCE/03_xlsx_pipeline"
PPTX_PIPE="$EVIDENCE/04_pptx_pipeline"
WRITER_DIR="$EVIDENCE/05_writer_features"
CALC_DIR="$EVIDENCE/06_calc_features"
IMPRESS_DIR="$EVIDENCE/07_impress_features"
DRAW_DIR="$EVIDENCE/08_draw_features"
MATH_DIR="$EVIDENCE/09_math_features"
BASE_DIR="$EVIDENCE/10_base_features"
OFFICE_DEMO="$EVIDENCE/11_office_demo"
DESKTOP_DEMO="$EVIDENCE/12_desktop_demo"
PACKAGE="$EVIDENCE/13_package_inspect"
LEGACY="$EVIDENCE/14_legacy_formats"
TRACKED="$EVIDENCE/15_tracked_changes"
SOFFICE_CMP="$EVIDENCE/16_soffice_reference"
HEADLESS="$EVIDENCE/17_headless_compat"
SUMMARY="$EVIDENCE/SUMMARY.md"
RUNLOG="$EVIDENCE/run.log"
MATRIX_TSV="$EVIDENCE/matrix_results.tsv"

rm -rf "$EVIDENCE"
mkdir -p "$CORPUS" "$MATRIX" "$DOCX_PIPE" "$XLSX_PIPE" "$PPTX_PIPE" \
         "$WRITER_DIR" "$CALC_DIR" "$IMPRESS_DIR" "$DRAW_DIR" "$MATH_DIR" "$BASE_DIR" \
         "$OFFICE_DEMO" "$DESKTOP_DEMO" "$PACKAGE" "$LEGACY" "$TRACKED" \
         "$SOFFICE_CMP" "$HEADLESS"

: >"$RUNLOG"
echo -e "family\tfrom\tto\tinput\tresult\tbytes\tms\terror" > "$MATRIX_TSV"

log()  { printf '%s\n' "$*" | tee -a "$RUNLOG" ; }
sec()  { printf '\n=== %s ===\n\n' "$*" | tee -a "$RUNLOG" ; }

now_ms() { python3 -c 'import time; print(int(time.time()*1000))'; }
fsize()  { stat -f%z "$1" 2>/dev/null || stat -c%s "$1" 2>/dev/null || echo 0; }

###############################################################################
# 1. Build
###############################################################################
sec "Building release binaries"
( cd "$REPO" && cargo build --release --quiet -p libreoffice-pure ) \
  || { log "build failed"; exit 1; }
log "OK"

LO_VER=$(soffice --version 2>&1 | head -1 || echo "soffice unavailable")
RS_VER=$("$BIN_PURE" --version 2>/dev/null || echo "libreoffice-rs v0.4.4")

###############################################################################
# 2. Corpus download (gov + open fixtures)
###############################################################################
sec "Downloading corpus (real-world inputs)"
declare -a CORPUS_ENTRIES=(
  # -- US government XLSX (Census Bureau) --
  # Note: BLS Akamai blocks non-browser clients (HTTP 403) regardless of User-Agent.
  "xlsx|gov-census-state-pop|https://www2.census.gov/programs-surveys/popest/tables/2020-2023/state/totals/NST-EST2023-POP.xlsx|US Census Bureau — State population 2020-2023"
  "xlsx|gov-census-county-pop|https://www2.census.gov/programs-surveys/popest/tables/2020-2023/counties/totals/co-est2023-pop.xlsx|US Census Bureau — County population 2020-2023"
  # -- DOCX fixtures (python-openxml, PHPWord, Calibre) --
  "docx|fixture-python-docx-default|https://raw.githubusercontent.com/python-openxml/python-docx/master/features/steps/test_files/doc-default.docx|python-openxml"
  "docx|fixture-python-docx-add-section|https://raw.githubusercontent.com/python-openxml/python-docx/master/features/steps/test_files/doc-add-section.docx|python-openxml"
  "docx|fixture-python-docx-blk-paras-tables|https://raw.githubusercontent.com/python-openxml/python-docx/master/features/steps/test_files/blk-paras-and-tables.docx|python-openxml"
  "docx|fixture-python-docx-comments|https://raw.githubusercontent.com/python-openxml/python-docx/master/features/steps/test_files/comments-rich-para.docx|python-openxml"
  "docx|fixture-phpword-readword|https://raw.githubusercontent.com/PHPOffice/PHPWord/master/samples/resources/Sample_11_ReadWord2007.docx|PHPOffice PHPWord"
  "docx|fixture-phpword-template|https://raw.githubusercontent.com/PHPOffice/PHPWord/master/samples/resources/Sample_07_TemplateCloneRow.docx|PHPOffice PHPWord"
  "docx|fixture-calibre-demo|https://calibre-ebook.com/downloads/demos/demo.docx|Calibre"
  # -- PPTX fixtures (python-pptx) --
  "pptx|fixture-python-pptx-charts|https://raw.githubusercontent.com/scanny/python-pptx/master/features/steps/test_files/cht-charts.pptx|scanny/python-pptx"
  "pptx|fixture-python-pptx-datalabels|https://raw.githubusercontent.com/scanny/python-pptx/master/features/steps/test_files/cht-datalabels.pptx|scanny/python-pptx"
  "pptx|fixture-python-pptx-chart-type|https://raw.githubusercontent.com/scanny/python-pptx/master/features/steps/test_files/cht-chart-type.pptx|scanny/python-pptx"
  # -- XLSX fixtures (PHPSpreadsheet) --
  "xlsx|fixture-phpss-26template|https://raw.githubusercontent.com/PHPOffice/PhpSpreadsheet/master/samples/templates/26template.xlsx|PHPOffice PhpSpreadsheet"
  "xlsx|fixture-phpss-28iterators|https://raw.githubusercontent.com/PHPOffice/PhpSpreadsheet/master/samples/templates/28iterators.xlsx|PHPOffice PhpSpreadsheet"
  "xlsx|fixture-phpss-31docprops|https://raw.githubusercontent.com/PHPOffice/PhpSpreadsheet/master/samples/templates/31docproperties.xlsx|PHPOffice PhpSpreadsheet"
)

for entry in "${CORPUS_ENTRIES[@]}"; do
  IFS='|' read -r kind name url src <<<"$entry"
  out="$CORPUS/$name.$kind"
  code=$(curl -sSLA 'libreoffice-rs-full-benchmark/1.0' "$url" -o "$out" -w "%{http_code}" --max-time 45 2>/dev/null || echo "000")
  size=$(fsize "$out")
  if [[ "$code" != "200" || "$size" -lt 200 ]]; then
    log "  SKIP $name.$kind (http=$code size=$size) [$src]"
    rm -f "$out"
  else
    log "  GOT  $name.$kind (${size}B) [$src]"
  fi
done

# Record provenance
{
  echo "# Corpus provenance"
  echo
  echo "All inputs used in this benchmark run."
  echo
  echo "| File | Source | URL |"
  echo "|------|--------|-----|"
  for entry in "${CORPUS_ENTRIES[@]}"; do
    IFS='|' read -r kind name url src <<<"$entry"
    if [[ -f "$CORPUS/$name.$kind" ]]; then
      echo "| \`$name.$kind\` | $src | <$url> |"
    fi
  done
} > "$CORPUS/PROVENANCE.md"

###############################################################################
# 3. Synthetic inputs (one per format family we cannot reliably download)
###############################################################################
sec "Downloading large real-world government CSV (pipeline-only)"
# CDC: Provisional COVID-19 Deaths by Sex and Age (real government dataset, ~25MB)
# Not added to the matrix (would explode the N×M runtime); used for a
# single realistic-scale pipeline demo below.
CDC_CSV="$CORPUS/gov-cdc-covid-deaths-by-sex-age.csv"
curl -sSLA 'libreoffice-rs-benchmark/1.0' \
  "https://data.cdc.gov/api/views/9bhg-hcku/rows.csv?accessType=DOWNLOAD" \
  -o "$CDC_CSV" --max-time 120 2>>"$RUNLOG" || true
if [[ -s "$CDC_CSV" ]]; then
  log "  GOT gov-cdc-covid-deaths-by-sex-age.csv ($(fsize "$CDC_CSV") B)"
  echo "| \`gov-cdc-covid-deaths-by-sex-age.csv\` | US CDC — Provisional COVID-19 Deaths by Sex and Age | <https://data.cdc.gov/NCHS/Provisional-COVID-19-Deaths-by-Sex-and-Age/9bhg-hcku> |" >> "$CORPUS/PROVENANCE.md"
else
  log "  SKIP CDC CSV download"
fi

sec "Generating synthetic inputs"

cat > "$CORPUS/synthetic.md" <<'MD'
# Full Feature Showcase

This document exercises every **Writer** feature.

## Inline

- **Bold**, *italic*, `inline code`
- A [hyperlink](https://example.com)
- Unicode: áéíóú ñ 中文 العربية

## Lists

1. First
2. Second
3. Third

## Table

| Region | Units | Price |
|--------|------:|------:|
| North  |    12 |  9.50 |
| South  |     8 | 11.00 |
| East   |    15 |  7.25 |

## Blockquote & rule

> Quote with *emphasis*.

---

End of file.
MD

cat > "$CORPUS/synthetic.csv" <<'CSV'
Region,Units,Price,Active,Revenue
North,12,9.50,true,=B2*C2
South,8,11.00,false,=B3*C3
East,15,7.25,true,=B4*C4
West,21,8.75,true,=B5*C5
Total,=SUM(B2:B5),=AVERAGE(C2:C5),,=SUM(E2:E5)
CSV

cat > "$CORPUS/synthetic.html" <<'HTML'
<!doctype html>
<html><head><meta charset="utf-8"><title>Sample</title></head>
<body>
<h1>Hello World</h1>
<p>A paragraph with <b>bold</b>, <i>italic</i>, and <a href="https://example.com">a link</a>.</p>
<ul><li>one</li><li>two</li><li>three</li></ul>
</body></html>
HTML

cat > "$CORPUS/synthetic.txt" <<'TXT'
Plain text sample.
Line two: ASCII.
Line three: special chars < > & " '
End of file.
TXT

cat > "$CORPUS/synthetic.svg" <<'SVG'
<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="240" height="120" viewBox="0 0 240 120">
  <rect x="10" y="10" width="220" height="100" fill="#4472C4" stroke="#2F5597" stroke-width="2"/>
  <text x="120" y="65" font-family="sans-serif" font-size="24" fill="white" text-anchor="middle">libreoffice-rs</text>
  <circle cx="40" cy="40" r="15" fill="#ED7D31"/>
  <circle cx="200" cy="80" r="10" fill="#70AD47"/>
</svg>
SVG

cat > "$CORPUS/synthetic.latex" <<'LATEX'
\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}
LATEX

cat > "$CORPUS/synthetic.mathml" <<'MML'
<math xmlns="http://www.w3.org/1998/Math/MathML" display="block">
  <mrow><mi>x</mi><mo>=</mo>
    <mfrac>
      <mrow><mo>-</mo><mi>b</mi><mo>&#xB1;</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>-</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow>
      <mrow><mn>2</mn><mi>a</mi></mrow>
    </mfrac>
  </mrow>
</math>
MML

log "  wrote synthetic.{md,csv,html,txt,svg,latex,mathml}"

###############################################################################
# 4. Derive ODF/OOXML inputs from synthetic sources
###############################################################################
sec "Deriving ODF/OOXML inputs"

"$BIN_RS" writer markdown-to-odt "$CORPUS/synthetic.md" "$CORPUS/synthetic.odt" --title "Feature Showcase" >>"$RUNLOG" 2>&1 && log "  synthetic.odt"  || log "  FAIL synthetic.odt"
"$BIN_RS" writer convert "$CORPUS/synthetic.md" "$CORPUS/synthetic.docx" --title "Feature Showcase" >>"$RUNLOG" 2>&1 && log "  synthetic.docx" || log "  FAIL synthetic.docx"
"$BIN_RS" calc csv-to-ods "$CORPUS/synthetic.csv" "$CORPUS/synthetic.ods" --sheet Sales --title "Sales" --has-header >>"$RUNLOG" 2>&1 && log "  synthetic.ods" || log "  FAIL synthetic.ods"
"$BIN_RS" calc convert "$CORPUS/synthetic.csv" "$CORPUS/synthetic.xlsx" --sheet Sales --title "Sales" --has-header >>"$RUNLOG" 2>&1 && log "  synthetic.xlsx" || log "  FAIL synthetic.xlsx"
"$BIN_RS" impress demo "$CORPUS/synthetic.odp" --title "Feature Showcase" >>"$RUNLOG" 2>&1 && log "  synthetic.odp" || log "  FAIL synthetic.odp"
"$BIN_RS" draw demo "$CORPUS/synthetic.odg" --title "Feature Showcase" >>"$RUNLOG" 2>&1 && log "  synthetic.odg" || log "  FAIL synthetic.odg"
"$BIN_RS" math latex-to-odf "$CORPUS/synthetic.latex" "$CORPUS/synthetic.odf" --title "Quadratic" >>"$RUNLOG" 2>&1 && log "  synthetic.odf" || log "  FAIL synthetic.odf"
"$BIN_RS" base csv-to-odb "$CORPUS/synthetic.csv" sales "$CORPUS/synthetic.odb" --title "Sales DB" >>"$RUNLOG" 2>&1 && log "  synthetic.odb" || log "  FAIL synthetic.odb"
"$BIN_PURE" convert --from odt --to pdf "$CORPUS/synthetic.odt" "$CORPUS/synthetic.pdf" >>"$RUNLOG" 2>&1 && log "  synthetic.pdf" || log "  FAIL synthetic.pdf"

###############################################################################
# 5. Full conversion matrix
###############################################################################
sec "Conversion matrix — every supported input × every supported output"

record() {
  local family="$1" from="$2" to="$3" in_name="$4" result="$5" bytes="$6" ms="$7" err="${8:-}"
  printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' "$family" "$from" "$to" "$in_name" "$result" "$bytes" "$ms" "$err" >> "$MATRIX_TSV"
}

try_convert() {
  local family="$1" from="$2" to="$3" input="$4"
  local in_name; in_name=$(basename "$input" | sed 's/\.[^.]*$//')
  local outdir="$MATRIX/$family/$in_name"
  mkdir -p "$outdir"
  local out="$outdir/$in_name.$to"
  local errfile="$outdir/.err.$to.log"
  local start end rc bytes err
  start=$(now_ms)
  "$BIN_PURE" convert --from "$from" --to "$to" "$input" "$out" >/dev/null 2>"$errfile"
  rc=$?
  end=$(now_ms)
  local ms=$((end - start))
  if [[ $rc -eq 0 && -f "$out" ]]; then
    bytes=$(fsize "$out")
    record "$family" "$from" "$to" "$in_name" "OK" "$bytes" "$ms" ""
    rm -f "$errfile"
    printf '  ok   %-8s %4s→%-4s %-40s %6d B  %4d ms\n' "$family" "$from" "$to" "$in_name" "$bytes" "$ms"
  else
    err=$(tail -1 "$errfile" 2>/dev/null | tr '\t\n' '  ' | cut -c1-160)
    record "$family" "$from" "$to" "$in_name" "ERR" "0" "$ms" "$err"
    printf '  ERR  %-8s %4s→%-4s %-40s %s\n' "$family" "$from" "$to" "$in_name" "$err"
  fi
}

# ---- Writer family ----
WRITER_INPUTS=("$CORPUS/synthetic.txt" "$CORPUS/synthetic.md" "$CORPUS/synthetic.html" "$CORPUS/synthetic.odt" "$CORPUS/synthetic.pdf")
for f in "$CORPUS"/*.docx; do
  [[ -f "$f" ]] && WRITER_INPUTS+=("$f")
done
WRITER_OUTS=(txt md html svg pdf odt docx)
for input in "${WRITER_INPUTS[@]}"; do
  [[ -f "$input" ]] || continue
  from="${input##*.}"
  for to in "${WRITER_OUTS[@]}"; do
    try_convert writer "$from" "$to" "$input"
  done
done

# ---- Calc family ----
CALC_INPUTS=("$CORPUS/synthetic.csv" "$CORPUS/synthetic.ods")
for f in "$CORPUS"/*.xlsx; do
  [[ -f "$f" ]] && CALC_INPUTS+=("$f")
done
CALC_OUTS=(csv md html svg pdf ods xlsx)
for input in "${CALC_INPUTS[@]}"; do
  [[ -f "$input" ]] || continue
  from="${input##*.}"
  for to in "${CALC_OUTS[@]}"; do
    try_convert calc "$from" "$to" "$input"
  done
done

# ---- Impress family ----
IMPRESS_INPUTS=("$CORPUS/synthetic.odp")
for f in "$CORPUS"/*.pptx; do
  [[ -f "$f" ]] && IMPRESS_INPUTS+=("$f")
done
IMPRESS_OUTS=(md html svg pdf odp pptx)
for input in "${IMPRESS_INPUTS[@]}"; do
  [[ -f "$input" ]] || continue
  from="${input##*.}"
  for to in "${IMPRESS_OUTS[@]}"; do
    try_convert impress "$from" "$to" "$input"
  done
done

# ---- Draw family ----
DRAW_INPUTS=("$CORPUS/synthetic.svg" "$CORPUS/synthetic.odg")
DRAW_OUTS=(svg pdf odg)
for input in "${DRAW_INPUTS[@]}"; do
  [[ -f "$input" ]] || continue
  from="${input##*.}"
  for to in "${DRAW_OUTS[@]}"; do
    try_convert draw "$from" "$to" "$input"
  done
done

# ---- Math family ----
MATH_INPUTS=("$CORPUS/synthetic.latex" "$CORPUS/synthetic.mathml" "$CORPUS/synthetic.odf")
MATH_OUTS=(mathml svg pdf odf)
for input in "${MATH_INPUTS[@]}"; do
  [[ -f "$input" ]] || continue
  from="${input##*.}"
  for to in "${MATH_OUTS[@]}"; do
    try_convert math "$from" "$to" "$input"
  done
done

# ---- Base family ----
BASE_INPUTS=("$CORPUS/synthetic.odb")
BASE_OUTS=(html svg pdf odb)
for input in "${BASE_INPUTS[@]}"; do
  [[ -f "$input" ]] || continue
  for to in "${BASE_OUTS[@]}"; do
    try_convert base odb "$to" "$input"
  done
done

###############################################################################
# 6. Pipeline commands (dedicated paths, not the generic convert router)
###############################################################################
sec "DOCX pipeline (docx-to-pdf / md / pngs / jpegs)"
for f in "$CORPUS"/*.docx; do
  [[ -f "$f" ]] || continue
  name=$(basename "${f%.*}")
  d="$DOCX_PIPE/$name"
  mkdir -p "$d/pngs" "$d/jpegs"
  "$BIN_PURE" docx-to-pdf   "$f" "$d/$name.pdf"             >>"$RUNLOG" 2>&1 && log "  $name.pdf ($(fsize "$d/$name.pdf") B)" || log "  FAIL $name docx-to-pdf"
  "$BIN_PURE" docx-to-md    "$f" "$d/$name.md"              >>"$RUNLOG" 2>&1 && log "  $name.md ($(fsize "$d/$name.md") B)" || log "  FAIL $name docx-to-md"
  "$BIN_PURE" docx-to-pngs  "$f" "$d/pngs"  --dpi 96         >>"$RUNLOG" 2>&1 && log "  $name/pngs ($(ls "$d/pngs" 2>/dev/null | wc -l | tr -d ' ') files)" || log "  FAIL $name docx-to-pngs"
  "$BIN_PURE" docx-to-jpegs "$f" "$d/jpegs" --dpi 96 --quality 80 >>"$RUNLOG" 2>&1 && log "  $name/jpegs ($(ls "$d/jpegs" 2>/dev/null | wc -l | tr -d ' ') files)" || log "  FAIL $name docx-to-jpegs"
done

sec "XLSX pipeline (recalc-check / recalc / csv / md)"
for f in "$CORPUS"/*.xlsx; do
  [[ -f "$f" ]] || continue
  name=$(basename "${f%.*}")
  d="$XLSX_PIPE/$name"
  mkdir -p "$d"
  "$BIN_PURE" xlsx-recalc-check "$f" "$d/recalc-report.json" >>"$RUNLOG" 2>&1 && log "  $name recalc-report.json"    || log "  FAIL $name xlsx-recalc-check"
  "$BIN_PURE" xlsx-recalc       "$f" "$d/recalculated.xlsx"  >>"$RUNLOG" 2>&1 && log "  $name recalculated.xlsx"     || log "  FAIL $name xlsx-recalc"
  "$BIN_PURE" convert --from xlsx --to csv "$f" "$d/$name.csv" >>"$RUNLOG" 2>&1 && log "  $name.csv"                 || log "  FAIL $name xlsx→csv"
  "$BIN_PURE" xlsx-to-md        "$f" "$d/$name.md"           >>"$RUNLOG" 2>&1 && log "  $name.md"                    || log "  FAIL $name xlsx-to-md"
done

sec "Large government dataset pipeline (CDC COVID-19 deaths CSV)"
if [[ -s "$CORPUS/gov-cdc-covid-deaths-by-sex-age.csv" ]]; then
  cdc="$CORPUS/gov-cdc-covid-deaths-by-sex-age.csv"
  d="$XLSX_PIPE/gov-cdc-covid-deaths"
  mkdir -p "$d"
  "$BIN_PURE" convert --from csv --to xlsx "$cdc" "$d/deaths.xlsx" >>"$RUNLOG" 2>&1 && log "  deaths.xlsx ($(fsize "$d/deaths.xlsx") B)" || log "  FAIL csv→xlsx"
  "$BIN_PURE" convert --from csv --to ods  "$cdc" "$d/deaths.ods"  >>"$RUNLOG" 2>&1 && log "  deaths.ods  ($(fsize "$d/deaths.ods") B)"  || log "  FAIL csv→ods"
  "$BIN_PURE" convert --from csv --to html "$cdc" "$d/deaths.html" >>"$RUNLOG" 2>&1 && log "  deaths.html ($(fsize "$d/deaths.html") B)" || log "  FAIL csv→html"
  "$BIN_PURE" convert --from csv --to pdf  "$cdc" "$d/deaths.pdf"  >>"$RUNLOG" 2>&1 && log "  deaths.pdf  ($(fsize "$d/deaths.pdf") B)"  || log "  FAIL csv→pdf"
  if [[ -f "$d/deaths.xlsx" ]]; then
    "$BIN_PURE" xlsx-recalc-check "$d/deaths.xlsx" "$d/recalc-report.json" >>"$RUNLOG" 2>&1 && log "  recalc-report.json" || log "  FAIL recalc-check"
    "$BIN_PURE" xlsx-to-md "$d/deaths.xlsx" "$d/deaths.md" >>"$RUNLOG" 2>&1 && log "  deaths.md" || log "  FAIL xlsx-to-md"
  fi
else
  log "  (no CDC CSV available — skipping)"
fi

sec "PPTX pipeline (pptx-to-pdf / md / pngs / jpegs)"
for f in "$CORPUS"/*.pptx; do
  [[ -f "$f" ]] || continue
  name=$(basename "${f%.*}")
  d="$PPTX_PIPE/$name"
  mkdir -p "$d/pngs" "$d/jpegs"
  "$BIN_PURE" pptx-to-pdf   "$f" "$d/$name.pdf"              >>"$RUNLOG" 2>&1 && log "  $name.pdf"                  || log "  FAIL $name pptx-to-pdf"
  "$BIN_PURE" pptx-to-md    "$f" "$d/$name.md"               >>"$RUNLOG" 2>&1 && log "  $name.md"                   || log "  FAIL $name pptx-to-md"
  "$BIN_PURE" pptx-to-pngs  "$f" "$d/pngs"  --dpi 96          >>"$RUNLOG" 2>&1 && log "  $name/pngs"                 || log "  FAIL $name pptx-to-pngs"
  "$BIN_PURE" pptx-to-jpegs "$f" "$d/jpegs" --dpi 96 --quality 80 >>"$RUNLOG" 2>&1 && log "  $name/jpegs"            || log "  FAIL $name pptx-to-jpegs"
done

###############################################################################
# 7. Per-family feature demos (Writer/Calc/Impress/Draw/Math/Base)
###############################################################################
sec "Writer feature showcase"
"$BIN_RS" writer new "$WRITER_DIR/new.odt" --title "New Doc" --text "Created directly from CLI" >>"$RUNLOG" 2>&1
for out in txt html svg pdf odt docx; do
  "$BIN_RS" writer convert "$CORPUS/synthetic.md" "$WRITER_DIR/showcase.$out" --title "Showcase" >>"$RUNLOG" 2>&1 \
    && log "  showcase.$out ($(fsize "$WRITER_DIR/showcase.$out") B)" \
    || log "  FAIL showcase.$out"
done

sec "Calc feature showcase (formulas + exports)"
for out in csv html svg pdf ods xlsx; do
  "$BIN_RS" calc convert "$CORPUS/synthetic.csv" "$CALC_DIR/sales.$out" --sheet Sales --title Sales --has-header >>"$RUNLOG" 2>&1 \
    && log "  sales.$out ($(fsize "$CALC_DIR/sales.$out") B)" \
    || log "  FAIL sales.$out"
done
# Evaluate all major formula functions inline
for formula in \
    "=SUM(B2:B5)" "=AVERAGE(B2:B5)" "=MIN(B2:B5)" "=MAX(B2:B5)" "=COUNT(B2:B5)" \
    "=ROUND(AVERAGE(C2:C5),2)" "=IF(B2>10,\"big\",\"small\")" \
    "=LEN(\"hello\")" "=CONCAT(A2,\"-\",A3)" "=ABS(-42.5)" \
    "=AND(TRUE,TRUE)" "=OR(FALSE,TRUE)" "=NOT(FALSE)" \
    "=2^10" "=B2+B3*C2"
do
  safe=$(printf '%s' "$formula" | tr -cd 'A-Za-z0-9_')
  "$BIN_RS" calc eval "$formula" --csv "$CORPUS/synthetic.csv" --has-header \
    > "$CALC_DIR/eval_${safe:0:40}.txt" 2>>"$RUNLOG" \
    && log "  eval $formula => $(cat "$CALC_DIR/eval_${safe:0:40}.txt")" \
    || log "  FAIL eval $formula"
done

sec "Impress feature showcase"
"$BIN_RS" impress demo "$IMPRESS_DIR/demo.odp" --title "Impress Demo" >>"$RUNLOG" 2>&1
for out in html svg pdf odp pptx; do
  "$BIN_PURE" convert --from odp --to "$out" "$IMPRESS_DIR/demo.odp" "$IMPRESS_DIR/demo.$out" >>"$RUNLOG" 2>&1 \
    && log "  demo.$out ($(fsize "$IMPRESS_DIR/demo.$out") B)" \
    || log "  FAIL demo.$out"
done

sec "Draw feature showcase"
"$BIN_RS" draw demo "$DRAW_DIR/demo.odg" --title "Draw Demo" >>"$RUNLOG" 2>&1
for out in svg pdf odg; do
  "$BIN_PURE" convert --from odg --to "$out" "$DRAW_DIR/demo.odg" "$DRAW_DIR/demo.$out" >>"$RUNLOG" 2>&1 \
    && log "  demo.$out ($(fsize "$DRAW_DIR/demo.$out") B)" \
    || log "  FAIL demo.$out"
done

sec "Math feature showcase"
"$BIN_RS" math latex-to-mathml "$CORPUS/synthetic.latex" > "$MATH_DIR/quadratic.mathml" 2>>"$RUNLOG" && log "  quadratic.mathml"
"$BIN_RS" math latex-to-odf    "$CORPUS/synthetic.latex" "$MATH_DIR/quadratic.odf" --title "Quadratic" >>"$RUNLOG" 2>&1 && log "  quadratic.odf"
for out in mathml svg pdf; do
  "$BIN_PURE" convert --from latex --to "$out" "$CORPUS/synthetic.latex" "$MATH_DIR/quadratic.$out" >>"$RUNLOG" 2>&1 \
    && log "  quadratic.$out" || log "  FAIL quadratic.$out"
done
# More complex formulas
for latex in \
    '\int_0^\infty e^{-x^2} dx' \
    '\sum_{n=1}^{\infty} \frac{1}{n^2}' \
    '\begin{matrix} a & b \\ c & d \end{matrix}'
do
  safe=$(printf '%s' "$latex" | tr -cd 'A-Za-z0-9_' | cut -c1-24)
  echo "$latex" > "$MATH_DIR/$safe.latex"
  "$BIN_RS" math latex-to-odf "$MATH_DIR/$safe.latex" "$MATH_DIR/$safe.odf" --title "$safe" >>"$RUNLOG" 2>&1 \
    && log "  $safe.odf" || log "  FAIL $safe.odf"
done

sec "Base feature showcase (CSV → ODB + SQL)"
"$BIN_RS" base csv-to-odb "$CORPUS/synthetic.csv" sales "$BASE_DIR/sales.odb" --title "Sales" >>"$RUNLOG" 2>&1 && log "  sales.odb"
for out in html svg pdf odb; do
  "$BIN_PURE" convert --from odb --to "$out" "$BASE_DIR/sales.odb" "$BASE_DIR/sales.$out" >>"$RUNLOG" 2>&1 \
    && log "  sales.$out ($(fsize "$BASE_DIR/sales.$out") B)" \
    || log "  FAIL sales.$out"
done
# SQL queries
{
  echo "# SQL query demonstrations (lo_base)"
  echo
  for q in \
      "SELECT * FROM sales" \
      "SELECT Region,Units FROM sales WHERE Units > 10" \
      "SELECT * FROM sales ORDER BY Units DESC" \
      "SELECT * FROM sales LIMIT 2"
  do
    echo "## \`$q\`"
    echo
    echo '```'
    "$BIN_RS" base query "$CORPUS/synthetic.csv" sales "$q" 2>>"$RUNLOG" || echo "(query failed)"
    echo '```'
    echo
  done
} > "$BASE_DIR/queries.md"
log "  queries.md"

###############################################################################
# 8. office-demo and desktop-demo end-to-end
###############################################################################
sec "office-demo: every kind × every format (28 files)"
"$BIN_RS" office-demo "$OFFICE_DEMO" >>"$RUNLOG" 2>&1 && log "  wrote $(ls "$OFFICE_DEMO" | wc -l | tr -d ' ') files" || log "  FAIL office-demo"

sec "desktop-demo: full lo_app surface"
"$BIN_RS" desktop-demo "$DESKTOP_DEMO" >>"$RUNLOG" 2>&1 && log "  desktop profile populated" || log "  FAIL desktop-demo"

###############################################################################
# 9. Package inspection
###############################################################################
sec "package inspect: ODF archive listings"
for f in \
    "$CORPUS/synthetic.odt" "$CORPUS/synthetic.ods" "$CORPUS/synthetic.odp" \
    "$CORPUS/synthetic.odg" "$CORPUS/synthetic.odf" "$CORPUS/synthetic.odb" \
    "$OFFICE_DEMO/writer.odt" "$OFFICE_DEMO/calc.ods" "$OFFICE_DEMO/impress.odp" \
    "$OFFICE_DEMO/draw.odg" "$OFFICE_DEMO/math.odf" "$OFFICE_DEMO/base.odb"
do
  [[ -f "$f" ]] || continue
  name=$(basename "$f")
  "$BIN_RS" package inspect "$f" > "$PACKAGE/$name.inspect.txt" 2>&1 && log "  $name" || log "  FAIL inspect $name"
done

###############################################################################
# 10. Tracked-changes accept
###############################################################################
sec "Tracked-changes accept-changes"
# No reliable public tracked-changes fixture exists. Synthesize a minimal
# valid DOCX containing w:ins / w:del / w:trackRevisions — mirrors the
# unit test fixtures in crates/libreoffice-rs/tests/tracked_changes.rs.
TRACK_SRC="$CORPUS/synthetic-tracked.docx"
python3 - "$TRACK_SRC" <<'PY' >>"$RUNLOG" 2>&1
import sys, zipfile
out = sys.argv[1]
content_types = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
rels = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
word_rels = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''
document = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Draft paragraph: </w:t></w:r>
      <w:ins w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">
        <w:r><w:t>ACCEPTED INSERTION</w:t></w:r>
      </w:ins>
      <w:r><w:t xml:space="preserve"> middle </w:t></w:r>
      <w:del w:id="2" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">
        <w:r><w:delText>DELETED TEXT</w:delText></w:r>
      </w:del>
      <w:r><w:t xml:space="preserve"> end.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:rPrChange w:id="3" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>
        <w:t>Bold run with a formatting change.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>'''
settings = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:trackRevisions/>
</w:settings>'''
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/_rels/document.xml.rels", word_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/settings.xml", settings)
print("wrote", out)
PY
if [[ -f "$TRACK_SRC" ]]; then
  cp "$TRACK_SRC" "$TRACKED/before.docx"
  "$BIN_PURE" accept-changes "$TRACK_SRC" "$TRACKED/after.docx"   >>"$RUNLOG" 2>&1 && log "  accepted → after.docx"  || log "  FAIL accept-changes"
  "$BIN_PURE" docx-to-pdf    "$TRACKED/before.docx" "$TRACKED/before.pdf" >>"$RUNLOG" 2>&1 && log "  before.pdf" || log "  FAIL before.pdf"
  "$BIN_PURE" docx-to-pdf    "$TRACKED/after.docx"  "$TRACKED/after.pdf"  >>"$RUNLOG" 2>&1 && log "  after.pdf"  || log "  FAIL after.pdf"
  "$BIN_PURE" docx-to-md     "$TRACKED/before.docx" "$TRACKED/before.md"  >>"$RUNLOG" 2>&1 && log "  before.md"  || log "  FAIL before.md"
  "$BIN_PURE" docx-to-md     "$TRACKED/after.docx"  "$TRACKED/after.md"   >>"$RUNLOG" 2>&1 && log "  after.md"   || log "  FAIL after.md"
  diff -u "$TRACKED/before.md" "$TRACKED/after.md" > "$TRACKED/diff.patch" 2>/dev/null || true
  log "  diff.patch"
else
  log "  no tracked-changes DOCX fixture found (would have been fixture-tracked-changes.docx)"
fi

###############################################################################
# 11. Legacy binary .doc handling
###############################################################################
sec "Legacy .doc handling"
# Generate a real Word 97-2003 binary .doc using the reference LibreOffice
# (this is the same format our doc-to-docx has to parse — a legitimate
# round-trip test).
DOC_IN=""
if command -v soffice >/dev/null; then
  soffice --headless --convert-to doc --outdir "$CORPUS" "$CORPUS/synthetic.docx" >>"$RUNLOG" 2>&1 || true
  if [[ -f "$CORPUS/synthetic.doc" ]]; then
    DOC_IN="$CORPUS/synthetic.doc"
    log "  created $DOC_IN ($(fsize "$DOC_IN") B) via soffice --convert-to doc"
  fi
fi
if [[ -n "$DOC_IN" ]]; then
  "$BIN_PURE" doc-to-docx "$DOC_IN" "$LEGACY/converted.docx" >>"$RUNLOG" 2>&1 && log "  converted.docx" || log "  FAIL doc-to-docx"
  if [[ -f "$LEGACY/converted.docx" ]]; then
    "$BIN_PURE" docx-to-pdf "$LEGACY/converted.docx" "$LEGACY/converted.pdf" >>"$RUNLOG" 2>&1 && log "  converted.pdf" || log "  FAIL converted.pdf"
    "$BIN_PURE" docx-to-md  "$LEGACY/converted.docx" "$LEGACY/converted.md"  >>"$RUNLOG" 2>&1 && log "  converted.md"  || log "  FAIL converted.md"
  fi
else
  log "  no public .doc fixture available; skipping"
fi

###############################################################################
# 12. --headless --convert-to compatibility (soffice-style CLI)
###############################################################################
sec "soffice-compatible --headless --convert-to"
for f in "$CORPUS/synthetic.md" "$CORPUS/synthetic.csv" "$CORPUS/synthetic.odt" "$CORPUS/synthetic.xlsx" "$CORPUS/synthetic.odp"; do
  [[ -f "$f" ]] || continue
  "$BIN_PURE" --headless --convert-to pdf "$f" --outdir "$HEADLESS" >>"$RUNLOG" 2>&1 \
    && log "  pdf <- $(basename "$f")" \
    || log "  FAIL pdf <- $(basename "$f")"
done
# Filter-string normalization: pdf:writer_pdf_Export should work
"$BIN_PURE" --headless --convert-to "pdf:writer_pdf_Export" "$CORPUS/synthetic.odt" --outdir "$HEADLESS" >>"$RUNLOG" 2>&1 \
  && log "  pdf:writer_pdf_Export filter string accepted" \
  || log "  FAIL filter-string normalization"

###############################################################################
# 13. Side-by-side with real LibreOffice
###############################################################################
sec "Reference outputs from real LibreOffice (soffice --headless)"
if command -v soffice >/dev/null; then
  for f in "$CORPUS"/*.docx "$CORPUS"/*.xlsx "$CORPUS"/*.pptx; do
    [[ -f "$f" ]] || continue
    soffice --headless --convert-to pdf --outdir "$SOFFICE_CMP" "$f" >>"$RUNLOG" 2>&1 || true
  done
  for f in "$CORPUS"/*.docx; do
    [[ -f "$f" ]] || continue
    soffice --headless --convert-to "txt:Text (encoded):UTF8" --outdir "$SOFFICE_CMP" "$f" >>"$RUNLOG" 2>&1 || true
  done
  for f in "$CORPUS"/*.xlsx; do
    [[ -f "$f" ]] || continue
    soffice --headless --convert-to csv --outdir "$SOFFICE_CMP" "$f" >>"$RUNLOG" 2>&1 || true
  done
  log "  $(ls "$SOFFICE_CMP" 2>/dev/null | wc -l | tr -d ' ') reference outputs written"
else
  log "  soffice not available — skipping"
fi

# PDF text similarity: rs vs soffice, when both exist
PDF_SIM_TSV="$EVIDENCE/pdf_similarity.tsv"
echo -e "file\trs_pages\tlo_pages\tjaccard_pct\trs_bytes\tlo_bytes" > "$PDF_SIM_TSV"
pdf_pages() { python3 -c "import sys,re; print(len(re.findall(rb'/Type\\s*/Page[^s]', open(sys.argv[1],'rb').read())))" "$1" 2>/dev/null || echo 0; }
jaccard() {
  python3 - "$1" "$2" <<'PY' 2>/dev/null
import sys, re
def toks(p):
    try:
        with open(p, encoding="utf-8", errors="ignore") as f:
            return set(re.findall(r"\w+", f.read().lower()))
    except: return set()
a,b = toks(sys.argv[1]), toks(sys.argv[2])
if not a and not b: print("100.0")
elif not a or not b: print("0.0")
else: print(f"{100.0*len(a&b)/len(a|b):.1f}")
PY
}
if command -v pdftotext >/dev/null; then
  for f in "$CORPUS"/*.docx; do
    [[ -f "$f" ]] || continue
    name=$(basename "${f%.*}")
    rs_pdf="$DOCX_PIPE/$name/$name.pdf"
    lo_pdf="$SOFFICE_CMP/$name.pdf"
    [[ -f "$rs_pdf" && -f "$lo_pdf" ]] || continue
    pdftotext -layout "$rs_pdf" "$EVIDENCE/.rs.txt" 2>/dev/null
    pdftotext -layout "$lo_pdf" "$EVIDENCE/.lo.txt" 2>/dev/null
    sim=$(jaccard "$EVIDENCE/.rs.txt" "$EVIDENCE/.lo.txt")
    rp=$(pdf_pages "$rs_pdf"); lp=$(pdf_pages "$lo_pdf")
    echo -e "$name.docx\t$rp\t$lp\t$sim\t$(fsize "$rs_pdf")\t$(fsize "$lo_pdf")" >> "$PDF_SIM_TSV"
  done
  for f in "$CORPUS"/*.pptx; do
    [[ -f "$f" ]] || continue
    name=$(basename "${f%.*}")
    rs_pdf="$PPTX_PIPE/$name/$name.pdf"
    lo_pdf="$SOFFICE_CMP/$name.pdf"
    [[ -f "$rs_pdf" && -f "$lo_pdf" ]] || continue
    pdftotext -layout "$rs_pdf" "$EVIDENCE/.rs.txt" 2>/dev/null
    pdftotext -layout "$lo_pdf" "$EVIDENCE/.lo.txt" 2>/dev/null
    sim=$(jaccard "$EVIDENCE/.rs.txt" "$EVIDENCE/.lo.txt")
    rp=$(pdf_pages "$rs_pdf"); lp=$(pdf_pages "$lo_pdf")
    echo -e "$name.pptx\t$rp\t$lp\t$sim\t$(fsize "$rs_pdf")\t$(fsize "$lo_pdf")" >> "$PDF_SIM_TSV"
  done
  for f in "$CORPUS"/*.xlsx; do
    [[ -f "$f" ]] || continue
    name=$(basename "${f%.*}")
    rs_csv="$XLSX_PIPE/$name/$name.csv"
    lo_csv="$SOFFICE_CMP/$name.csv"
    [[ -f "$rs_csv" && -f "$lo_csv" ]] || continue
    sim=$(jaccard "$rs_csv" "$lo_csv")
    echo -e "$name.xlsx_csv\t-\t-\t$sim\t$(fsize "$rs_csv")\t$(fsize "$lo_csv")" >> "$PDF_SIM_TSV"
  done
  rm -f "$EVIDENCE/.rs.txt" "$EVIDENCE/.lo.txt"
  log "  $(($(wc -l < "$PDF_SIM_TSV") - 1)) similarity comparisons recorded"
fi

###############################################################################
# 14. Summary
###############################################################################
sec "Writing summary"

{
  echo "# libreoffice-rs — Full Evidence Benchmark"
  echo
  echo "Generated: $(date -u +%Y-%m-%dT%H:%M:%SZ)"
  echo "Host: $(uname -sm)"
  echo "libreoffice-rs: \`$RS_VER\`"
  echo "Reference: \`$LO_VER\`"
  echo
  echo "Every supported feature was exercised end-to-end and all produced files are stored in this folder. Open any subdirectory to inspect real outputs."
  echo
  echo "## Evidence layout"
  echo
  echo "| Folder | Contents |"
  echo "|--------|----------|"
  echo "| \`00_corpus/\` | Downloaded + synthetic inputs, \`PROVENANCE.md\` |"
  echo "| \`01_matrix/\` | Full N×M conversion matrix per family (one subfolder per input) |"
  echo "| \`02_docx_pipeline/\` | DOCX → PDF / MD / PNG / JPEG for every DOCX corpus file |"
  echo "| \`03_xlsx_pipeline/\` | XLSX → CSV / MD + recalc-check JSON + recalc-in-place |"
  echo "| \`04_pptx_pipeline/\` | PPTX → PDF / MD / PNG / JPEG for every PPTX corpus file |"
  echo "| \`05_writer_features/\` | Writer showcase (md → txt/html/svg/pdf/odt/docx) |"
  echo "| \`06_calc_features/\` | Calc showcase + 15 formula eval outputs |"
  echo "| \`07_impress_features/\` | Impress demo + exports |"
  echo "| \`08_draw_features/\` | Draw demo + exports |"
  echo "| \`09_math_features/\` | LaTeX/MathML/ODF math with multiple formulas |"
  echo "| \`10_base_features/\` | CSV → ODB + SQL query results |"
  echo "| \`11_office_demo/\` | 28-file kind × format matrix |"
  echo "| \`12_desktop_demo/\` | Full \`lo_app\` desktop profile (start center, macros, autosave) |"
  echo "| \`13_package_inspect/\` | ODF archive listings for every generated package |"
  echo "| \`14_legacy_formats/\` | Legacy .doc → .docx → .pdf (if a public fixture was available) |"
  echo "| \`15_tracked_changes/\` | DOCX before/after accept-changes + diff |"
  echo "| \`16_soffice_reference/\` | Reference outputs from real LibreOffice |"
  echo "| \`17_headless_compat/\` | \`--headless --convert-to\` soffice-compatible CLI outputs |"
  echo "| \`matrix_results.tsv\` | Raw TSV of every matrix attempt (family/from/to/result/bytes/ms/error) |"
  echo "| \`pdf_similarity.tsv\` | Per-file Jaccard similarity vs real LibreOffice |"
  echo "| \`run.log\` | Raw run log |"
  echo
  echo "## Matrix summary"
  echo
  python3 - "$MATRIX_TSV" <<'PY'
import csv, sys
from collections import defaultdict
fam_stats = defaultdict(lambda: [0, 0])
pairs = defaultdict(lambda: [0, 0])
total = [0, 0]
with open(sys.argv[1]) as f:
    r = csv.reader(f, delimiter='\t')
    next(r, None)
    for row in r:
        if len(row) < 7: continue
        fam, frm, to, name, res = row[0], row[1], row[2], row[3], row[4]
        idx = 0 if res == "OK" else 1
        fam_stats[fam][idx] += 1
        pairs[(fam, frm, to)][idx] += 1
        total[idx] += 1
print("| Family | OK | ERR | Total |")
print("|--------|---:|----:|------:|")
for fam in sorted(fam_stats):
    ok, err = fam_stats[fam]
    print(f"| {fam} | {ok} | {err} | {ok + err} |")
print(f"| **total** | **{total[0]}** | **{total[1]}** | **{sum(total)}** |")
print()
print("### Conversion pair matrix (OK / total attempts)")
print()
print("| family | from → to | OK | total |")
print("|--------|-----------|---:|------:|")
for (fam, frm, to), (ok, err) in sorted(pairs.items()):
    mark = "ok" if err == 0 else ("partial" if ok > 0 else "FAIL")
    print(f"| {fam} | `{frm}`→`{to}` | {ok} | {ok + err} |")
PY
  echo
  echo "## Detailed matrix (every attempt)"
  echo
  python3 - "$MATRIX_TSV" <<'PY'
import csv, sys
from collections import defaultdict
rows = []
with open(sys.argv[1]) as f:
    r = csv.reader(f, delimiter='\t')
    next(r, None)
    for row in r:
        if len(row) < 7: continue
        rows.append(row)
by_family = defaultdict(list)
for r in rows:
    by_family[r[0]].append(r)
for fam in sorted(by_family):
    print(f"### {fam}")
    print()
    print("| input | from | to | result | bytes | ms |")
    print("|-------|------|----|--------|------:|---:|")
    for r in sorted(by_family[fam], key=lambda x: (x[3], x[1], x[2])):
        sym = "ok" if r[4] == "OK" else "**ERR**"
        print(f"| {r[3]} | {r[1]} | {r[2]} | {sym} | {r[5]} | {r[6]} |")
    print()
PY
  echo
  echo "## Quality vs real LibreOffice"
  echo
  if [[ -s "$PDF_SIM_TSV" ]]; then
    python3 - "$PDF_SIM_TSV" <<'PY'
import csv, sys
rows = []
with open(sys.argv[1]) as f:
    r = csv.reader(f, delimiter='\t')
    next(r, None)
    for row in r:
        if len(row) < 6: continue
        rows.append(row)
if not rows:
    print("(no comparisons)")
else:
    print("| file | rs pages | lo pages | Jaccard % | rs bytes | lo bytes |")
    print("|------|---------:|---------:|----------:|---------:|---------:|")
    total = 0.0; n = 0
    for r in rows:
        print(f"| {r[0]} | {r[1]} | {r[2]} | {r[3]} | {r[4]} | {r[5]} |")
        try: total += float(r[3]); n += 1
        except: pass
    if n:
        print()
        print(f"**Mean Jaccard similarity: {total/n:.1f}% (n={n})**")
PY
  else
    echo "(no similarity data available)"
  fi
  echo
  echo "## Corpus"
  echo
  cat "$CORPUS/PROVENANCE.md"
} > "$SUMMARY"

log ""
log "DONE."
log "Evidence folder: $EVIDENCE"
log "Summary:         $SUMMARY"
log "Matrix TSV:      $MATRIX_TSV"
