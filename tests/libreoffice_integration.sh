#!/usr/bin/env bash
#
# Integration tests: generate documents with libreoffice-rs and validate
# them by opening/converting with real LibreOffice.
#
# Requirements: LibreOffice installed, cargo built.
#
set -uo pipefail

SOFFICE="soffice"
BINARY="cargo run --quiet --"
SCRIPT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
WORK_DIR="$(mktemp -d)"
PASS=0
FAIL=0
ERRORS=""

cleanup() {
    rm -rf "$WORK_DIR"
}
trap cleanup EXIT

cd "$SCRIPT_DIR"

pass() {
    PASS=$((PASS + 1))
    echo "  PASS: $1"
}

fail() {
    FAIL=$((FAIL + 1))
    ERRORS="${ERRORS}\n  FAIL: $1 — $2"
    echo "  FAIL: $1 — $2"
}

# Check that a file exists and is non-empty
check_file() {
    local path="$1" label="$2"
    if [[ -f "$path" && -s "$path" ]]; then
        return 0
    else
        return 1
    fi
}

# Convert an ODF file to PDF using LibreOffice and check the result
convert_to_pdf() {
    local input="$1" label="$2"
    local outdir="$WORK_DIR/pdf_out"
    mkdir -p "$outdir"
    if $SOFFICE --headless --convert-to pdf --outdir "$outdir" "$input" 2>&1; then
        local basename
        basename="$(basename "${input%.*}").pdf"
        if check_file "$outdir/$basename" "$label"; then
            local size
            size=$(stat -f%z "$outdir/$basename" 2>/dev/null || stat -c%s "$outdir/$basename" 2>/dev/null)
            pass "$label → PDF ($size bytes)"
            return 0
        else
            fail "$label → PDF" "output PDF missing or empty"
            return 1
        fi
    else
        fail "$label → PDF" "soffice conversion failed"
        return 1
    fi
}

# Convert an ODF file to a different format and check
convert_to_format() {
    local input="$1" format="$2" label="$3"
    local outdir="$WORK_DIR/convert_out"
    mkdir -p "$outdir"
    if $SOFFICE --headless --convert-to "$format" --outdir "$outdir" "$input" 2>&1; then
        local basename ext
        ext="${format%%:*}"
        basename="$(basename "${input%.*}").$ext"
        if check_file "$outdir/$basename" "$label"; then
            pass "$label → $ext"
            return 0
        else
            fail "$label → $ext" "output file missing or empty"
            return 1
        fi
    else
        fail "$label → $ext" "soffice conversion failed"
        return 1
    fi
}

# Macro-check: open in LibreOffice and export to PDF (proves the file can be loaded)
macro_check() {
    local input="$1" label="$2"
    local outdir="$WORK_DIR/macro_out"
    mkdir -p "$outdir"

    # Use macro to open, get page/sheet count info, and close
    if $SOFFICE --headless --calc --convert-to pdf --outdir "$outdir" "$input" 2>/dev/null; then
        return 0
    fi
    return 1
}

echo "============================================"
echo " libreoffice-rs Integration Tests"
echo " LibreOffice: $($SOFFICE --version 2>/dev/null || echo 'unknown')"
echo " Working dir: $WORK_DIR"
echo "============================================"
echo ""

########################################################################
# 1. WRITER (ODT)
########################################################################
echo "--- Writer (ODT) ---"

# 1a. Create a simple ODT
ODT_SIMPLE="$WORK_DIR/simple.odt"
if $BINARY writer new "$ODT_SIMPLE" --title "Test Doc" --text "Hello, LibreOffice!" 2>&1; then
    if check_file "$ODT_SIMPLE" "writer new"; then
        pass "writer new → created $(stat -f%z "$ODT_SIMPLE") bytes"
        convert_to_pdf "$ODT_SIMPLE" "simple.odt"
        convert_to_format "$ODT_SIMPLE" "docx" "simple.odt"
        convert_to_format "$ODT_SIMPLE" "txt:Text" "simple.odt (txt)"
    else
        fail "writer new" "output file missing"
    fi
else
    fail "writer new" "command failed"
fi

# 1b. Markdown to ODT
ODT_MD="$WORK_DIR/markdown.odt"
if $BINARY writer markdown-to-odt examples/sample.md "$ODT_MD" --title "Markdown Test" 2>&1; then
    if check_file "$ODT_MD" "markdown-to-odt"; then
        pass "markdown-to-odt → created $(stat -f%z "$ODT_MD") bytes"
        convert_to_pdf "$ODT_MD" "markdown.odt"
    else
        fail "markdown-to-odt" "output file missing"
    fi
else
    fail "markdown-to-odt" "command failed"
fi

# 1c. Rich markdown with more content
RICH_MD="$WORK_DIR/rich.md"
cat > "$RICH_MD" <<'MDEOF'
# Chapter One

This is a paragraph with **bold**, *italic*, and `code` formatting.

## Section 1.1

Here is a [link](https://example.com) in text.

- First item
- Second item
- Third item

## Section 1.2

| Name  | Score | Grade |
|-------|-------|-------|
| Alice | 95    | A     |
| Bob   | 82    | B     |
| Carol | 78    | C     |

# Chapter Two

Another paragraph to verify multi-heading documents render correctly.
MDEOF

ODT_RICH="$WORK_DIR/rich.odt"
if $BINARY writer markdown-to-odt "$RICH_MD" "$ODT_RICH" --title "Rich Document" 2>&1; then
    if check_file "$ODT_RICH" "rich markdown"; then
        pass "rich markdown → created $(stat -f%z "$ODT_RICH") bytes"
        convert_to_pdf "$ODT_RICH" "rich.odt"
    else
        fail "rich markdown" "output file missing"
    fi
else
    fail "rich markdown" "command failed"
fi

echo ""

########################################################################
# 2. CALC (ODS)
########################################################################
echo "--- Calc (ODS) ---"

# 2a. Simple CSV → ODS
ODS_SIMPLE="$WORK_DIR/simple.ods"
if $BINARY calc csv-to-ods examples/sample.csv "$ODS_SIMPLE" --sheet "People" --title "People DB" 2>&1; then
    if check_file "$ODS_SIMPLE" "csv-to-ods"; then
        pass "csv-to-ods → created $(stat -f%z "$ODS_SIMPLE") bytes"
        convert_to_pdf "$ODS_SIMPLE" "simple.ods"
        convert_to_format "$ODS_SIMPLE" "xlsx" "simple.ods"
        convert_to_format "$ODS_SIMPLE" "csv" "simple.ods (csv)"
    else
        fail "csv-to-ods" "output file missing"
    fi
else
    fail "csv-to-ods" "command failed"
fi

# 2b. CSV with formulas
FORMULA_CSV="$WORK_DIR/formulas.csv"
cat > "$FORMULA_CSV" <<'CSVEOF'
A,B,Sum,Product
10,20,=A2+B2,=A2*B2
5,15,=A3+B3,=A3*B3
1,2,=A4+B4,=A4*B4
CSVEOF

ODS_FORMULAS="$WORK_DIR/formulas.ods"
if $BINARY calc csv-to-ods "$FORMULA_CSV" "$ODS_FORMULAS" --sheet "Formulas" --title "Formula Test" 2>&1; then
    if check_file "$ODS_FORMULAS" "formulas.ods"; then
        pass "formula csv → ods created $(stat -f%z "$ODS_FORMULAS") bytes"
        convert_to_pdf "$ODS_FORMULAS" "formulas.ods"
    else
        fail "formulas csv" "output file missing"
    fi
else
    fail "formulas csv" "command failed"
fi

# 2c. Larger dataset
LARGE_CSV="$WORK_DIR/large.csv"
{
    echo "ID,Value,Category,Score"
    for i in $(seq 1 100); do
        echo "$i,$((RANDOM % 1000)),Cat$((i % 5)),$((RANDOM % 100))"
    done
} > "$LARGE_CSV"

ODS_LARGE="$WORK_DIR/large.ods"
if $BINARY calc csv-to-ods "$LARGE_CSV" "$ODS_LARGE" --sheet "Data" --title "Large Dataset" 2>&1; then
    if check_file "$ODS_LARGE" "large.ods"; then
        pass "large dataset → ods created $(stat -f%z "$ODS_LARGE") bytes"
        convert_to_pdf "$ODS_LARGE" "large.ods"
    else
        fail "large csv" "output file missing"
    fi
else
    fail "large csv" "command failed"
fi

echo ""

########################################################################
# 3. IMPRESS (ODP)
########################################################################
echo "--- Impress (ODP) ---"

ODP_DEMO="$WORK_DIR/demo.odp"
if $BINARY impress demo "$ODP_DEMO" --title "Test Presentation" 2>&1; then
    if check_file "$ODP_DEMO" "impress demo"; then
        pass "impress demo → created $(stat -f%z "$ODP_DEMO") bytes"
        convert_to_pdf "$ODP_DEMO" "demo.odp"
        convert_to_format "$ODP_DEMO" "pptx" "demo.odp"
    else
        fail "impress demo" "output file missing"
    fi
else
    fail "impress demo" "command failed"
fi

echo ""

########################################################################
# 4. DRAW (ODG)
########################################################################
echo "--- Draw (ODG) ---"

ODG_DEMO="$WORK_DIR/demo.odg"
if $BINARY draw demo "$ODG_DEMO" --title "Test Diagram" 2>&1; then
    if check_file "$ODG_DEMO" "draw demo"; then
        pass "draw demo → created $(stat -f%z "$ODG_DEMO") bytes"
        convert_to_pdf "$ODG_DEMO" "demo.odg"
        convert_to_format "$ODG_DEMO" "svg" "demo.odg"
    else
        fail "draw demo" "output file missing"
    fi
else
    fail "draw demo" "command failed"
fi

echo ""

########################################################################
# 5. MATH (ODF formula)
########################################################################
echo "--- Math (ODF) ---"

ODF_MATH="$WORK_DIR/formula.odf"
if $BINARY math latex-to-odf examples/formula.txt "$ODF_MATH" --title "Test Formula" 2>&1; then
    if check_file "$ODF_MATH" "math latex-to-odf"; then
        pass "latex-to-odf → created $(stat -f%z "$ODF_MATH") bytes"
        convert_to_pdf "$ODF_MATH" "formula.odf"
    else
        fail "latex-to-odf" "output file missing"
    fi
else
    fail "latex-to-odf" "command failed"
fi

echo ""

########################################################################
# 6. BASE (ODB)
########################################################################
echo "--- Base (ODB) ---"

ODB_DEMO="$WORK_DIR/people.odb"
if $BINARY base csv-to-odb examples/sample.csv people "$ODB_DEMO" --title "People DB" 2>&1; then
    if check_file "$ODB_DEMO" "base csv-to-odb"; then
        pass "csv-to-odb → created $(stat -f%z "$ODB_DEMO") bytes"
        # ODB conversion to PDF often not supported, just verify the file is valid ZIP
        if unzip -t "$ODB_DEMO" > /dev/null 2>&1; then
            pass "people.odb is valid ZIP archive"
        else
            fail "people.odb" "not a valid ZIP archive"
        fi
    else
        fail "csv-to-odb" "output file missing"
    fi
else
    fail "csv-to-odb" "command failed"
fi

echo ""

########################################################################
# 7. PACKAGE INSPECT (verify internal structure)
########################################################################
echo "--- Package Inspect ---"

for f in "$ODT_SIMPLE" "$ODS_SIMPLE" "$ODP_DEMO" "$ODG_DEMO"; do
    if [ -f "$f" ]; then
        label="$(basename "$f")"
        output=$($BINARY package inspect "$f" 2>&1)
        # Check for required ODF entries
        has_content=false
        has_meta=false
        has_manifest=false
        has_mimetype=false
        if echo "$output" | grep -q "content.xml"; then has_content=true; fi
        if echo "$output" | grep -q "meta.xml"; then has_meta=true; fi
        if echo "$output" | grep -q "manifest.xml"; then has_manifest=true; fi
        if echo "$output" | grep -q "mimetype"; then has_mimetype=true; fi

        if $has_content && $has_meta && $has_manifest && $has_mimetype; then
            pass "$label has all required ODF entries"
        else
            fail "$label structure" "missing: content=$has_content meta=$has_meta manifest=$has_manifest mimetype=$has_mimetype"
        fi
    fi
done

echo ""

########################################################################
# 8. CONTENT VALIDATION (extract and check XML)
########################################################################
echo "--- XML Content Validation ---"

# Check content.xml of the simple ODT for expected text
EXTRACT_DIR="$WORK_DIR/extract"
mkdir -p "$EXTRACT_DIR"
if unzip -o "$ODT_SIMPLE" content.xml -d "$EXTRACT_DIR" > /dev/null 2>&1; then
    if grep -q "Hello, LibreOffice!" "$EXTRACT_DIR/content.xml"; then
        pass "simple.odt content.xml contains expected text"
    else
        fail "simple.odt content" "expected text not found in content.xml"
    fi

    if grep -q "office:document-content" "$EXTRACT_DIR/content.xml"; then
        pass "simple.odt has valid ODF root element"
    else
        fail "simple.odt XML" "missing office:document-content root"
    fi
else
    fail "simple.odt unzip" "could not extract content.xml"
fi

# Check ODS content
if unzip -o "$ODS_SIMPLE" content.xml -d "$EXTRACT_DIR" > /dev/null 2>&1; then
    if grep -q "Alice" "$EXTRACT_DIR/content.xml" && grep -q "Bob" "$EXTRACT_DIR/content.xml"; then
        pass "simple.ods content.xml contains CSV data"
    else
        fail "simple.ods content" "CSV data not found in content.xml"
    fi
else
    fail "simple.ods unzip" "could not extract content.xml"
fi

echo ""

########################################################################
# 9. ROUND-TRIP: LibreOffice re-export to ODF
########################################################################
echo "--- Round-Trip (re-export through LibreOffice) ---"

ROUNDTRIP_DIR="$WORK_DIR/roundtrip"
mkdir -p "$ROUNDTRIP_DIR"

# Open ODT in LibreOffice and re-export as ODT
if $SOFFICE --headless --convert-to odt --outdir "$ROUNDTRIP_DIR" "$ODT_RICH" 2>&1; then
    RT_FILE="$ROUNDTRIP_DIR/rich.odt"
    if check_file "$RT_FILE" "roundtrip odt"; then
        # Extract and compare
        RT_EXTRACT="$WORK_DIR/rt_extract"
        mkdir -p "$RT_EXTRACT"
        if unzip -o "$RT_FILE" content.xml -d "$RT_EXTRACT" > /dev/null 2>&1; then
            if grep -q "Chapter One" "$RT_EXTRACT/content.xml" && grep -q "Alice" "$RT_EXTRACT/content.xml"; then
                pass "ODT round-trip preserves content"
            else
                fail "ODT round-trip" "content not preserved after LibreOffice re-export"
            fi
        else
            fail "ODT round-trip extract" "could not extract re-exported file"
        fi
    else
        fail "ODT round-trip" "re-exported file missing"
    fi
else
    fail "ODT round-trip" "LibreOffice re-export failed"
fi

# ODS round-trip
if $SOFFICE --headless --convert-to ods --outdir "$ROUNDTRIP_DIR" "$ODS_SIMPLE" 2>&1; then
    RT_ODS="$ROUNDTRIP_DIR/simple.ods"
    if check_file "$RT_ODS" "roundtrip ods"; then
        pass "ODS round-trip produces valid file"
    else
        fail "ODS round-trip" "re-exported file missing"
    fi
else
    fail "ODS round-trip" "LibreOffice re-export failed"
fi

echo ""

########################################################################
# SUMMARY
########################################################################
TOTAL=$((PASS + FAIL))
echo "============================================"
echo " Results: $PASS passed, $FAIL failed (of $TOTAL)"
if [[ $FAIL -gt 0 ]]; then
    echo ""
    echo " Failures:"
    echo -e "$ERRORS"
fi
echo "============================================"

exit $FAIL
