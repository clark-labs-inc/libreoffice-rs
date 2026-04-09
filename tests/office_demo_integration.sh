#!/usr/bin/env bash
#
# Integration test for the new `office-demo` CLI command and the new
# DOCX / XLSX / PPTX / HTML / SVG / PDF exporters added when porting
# parity features into the workspace.
#
# It generates one document of every kind in every supported format and
# then asks real LibreOffice to convert each native format (ODT, DOCX,
# ODS, XLSX, ODP, PPTX, ODG, ODF, ODB) to PDF. The test passes if every
# conversion succeeds and the resulting PDF is non-empty.
#
# Requirements: LibreOffice (`soffice`) on PATH, the workspace built.
set -uo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
SOFFICE="soffice"
BINARY="cargo run --quiet --release -p libreoffice-pure --"
WORK_DIR="$(mktemp -d)"
PASS=0
FAIL=0
ERRORS=""

cleanup() { rm -rf "$WORK_DIR"; }
trap cleanup EXIT

cd "$SCRIPT_DIR"

pass() { PASS=$((PASS + 1)); echo "  PASS: $1"; }
fail() {
    FAIL=$((FAIL + 1))
    ERRORS="${ERRORS}\n  FAIL: $1 — $2"
    echo "  FAIL: $1 — $2"
}

if ! command -v "$SOFFICE" >/dev/null 2>&1; then
    echo "soffice not found on PATH; skipping office_demo_integration"
    exit 0
fi

echo "Building release CLI..."
cargo build --release --quiet -p libreoffice-pure 2>&1 || {
    echo "release build failed"
    exit 1
}

DEMO_DIR="$WORK_DIR/demo"
echo
echo "--- Generating office-demo into $DEMO_DIR ---"
if $BINARY office-demo "$DEMO_DIR" >/dev/null 2>&1; then
    pass "office-demo generated"
else
    fail "office-demo" "CLI returned non-zero"
    echo "============================================"
    echo " Results: $PASS passed, $FAIL failed"
    echo "============================================"
    exit 1
fi

# Sanity-check that every expected output exists and is non-empty.
EXPECTED=(
    writer.txt writer.html writer.svg writer.pdf writer.odt writer.docx
    calc.csv calc.html calc.svg calc.pdf calc.ods calc.xlsx
    impress.html impress.svg impress.pdf impress.odp impress.pptx
    draw.svg draw.pdf draw.odg
    math.mathml math.svg math.pdf math.odf
    base.html base.svg base.pdf base.odb
)
echo
echo "--- Verifying generated files exist ---"
for entry in "${EXPECTED[@]}"; do
    if [[ -s "$DEMO_DIR/$entry" ]]; then
        pass "$entry exists"
    else
        fail "$entry exists" "missing or empty"
    fi
done

# Round-trip every native LO format through real LibreOffice.
echo
echo "--- Round-tripping native LO formats through soffice ---"
ROUNDTRIP_DIR="$WORK_DIR/rt"
mkdir -p "$ROUNDTRIP_DIR"
for f in writer.odt writer.docx calc.ods calc.xlsx impress.odp impress.pptx draw.odg math.odf base.odb; do
    if $SOFFICE --headless --convert-to pdf --outdir "$ROUNDTRIP_DIR" "$DEMO_DIR/$f" >/dev/null 2>&1; then
        out="$ROUNDTRIP_DIR/${f%.*}.pdf"
        if [[ -s "$out" ]]; then
            pass "$f → soffice pdf"
        else
            fail "$f → soffice pdf" "empty output"
        fi
    else
        fail "$f → soffice pdf" "soffice conversion failed"
    fi
done

# DOCX text round-trip — make sure heading + body + table content survives.
echo
echo "--- DOCX text round-trip content check ---"
TXT_DIR="$WORK_DIR/txt"
mkdir -p "$TXT_DIR"
if $SOFFICE --headless --convert-to txt --outdir "$TXT_DIR" "$DEMO_DIR/writer.docx" >/dev/null 2>&1; then
    if grep -q "Office Demo" "$TXT_DIR/writer.txt" \
        && grep -q "bold" "$TXT_DIR/writer.txt" \
        && grep -q "https://example.com" "$TXT_DIR/writer.txt" \
        && grep -q "col a" "$TXT_DIR/writer.txt"; then
        pass "DOCX → txt preserves heading, bold, link and table"
    else
        fail "DOCX text round-trip" "missing expected content"
    fi
else
    fail "DOCX text round-trip" "soffice conversion failed"
fi

# XLSX → CSV content check.
echo
echo "--- XLSX CSV round-trip content check ---"
CSV_DIR="$WORK_DIR/csv"
mkdir -p "$CSV_DIR"
if $SOFFICE --headless --calc --convert-to csv --outdir "$CSV_DIR" "$DEMO_DIR/calc.xlsx" >/dev/null 2>&1; then
    if grep -q "region,units,price" "$CSV_DIR/calc.csv" \
        && grep -q "North,12" "$CSV_DIR/calc.csv"; then
        pass "XLSX → csv preserves header and rows"
    else
        fail "XLSX csv round-trip" "missing expected rows"
    fi
else
    fail "XLSX csv round-trip" "soffice conversion failed"
fi

echo
echo "============================================"
echo " Results: $PASS passed, $FAIL failed (of $((PASS + FAIL)))"
echo "============================================"
if [[ $FAIL -gt 0 ]]; then
    echo -e "$ERRORS"
    exit 1
fi
