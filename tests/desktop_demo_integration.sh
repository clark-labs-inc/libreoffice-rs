#!/usr/bin/env bash
#
# Integration test for the new `desktop-demo` CLI command and the lo_app
# desktop application surface.
#
# Generates a complete profile dir, then asserts:
#  1. every expected artifact exists (start-center.html, per-window shells,
#     exports, autosave snapshots, preferences.ini, recent.tsv,
#     recovery.tsv, workspace.txt)
#  2. start-center.html lists at least one template
#  3. each per-window HTML shell embeds an SVG preview tile
#  4. each native ODF artifact opens cleanly in real LibreOffice
#  5. the Writer ODT round-trips through soffice and the appended desktop
#     text survives
#  6. the Calc ODS round-trips and the desktop-recorded `Forecast` formula
#     evaluates to the expected value (= 600).
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
    echo "soffice not found on PATH; skipping desktop_demo_integration"
    exit 0
fi

echo "Building release CLI..."
cargo build --release --quiet -p libreoffice-pure 2>&1 || {
    echo "release build failed"
    exit 1
}

PROFILE="$WORK_DIR/profile"
echo
echo "--- Running desktop-demo into $PROFILE ---"
if $BINARY desktop-demo "$PROFILE" >/dev/null 2>&1; then
    pass "desktop-demo CLI returned success"
else
    fail "desktop-demo" "CLI returned non-zero"
    echo "============================================"
    echo " Results: $PASS passed, $FAIL failed"
    echo "============================================"
    exit 1
fi

# 1. Expected files exist.
EXPECTED=(
    start-center.html
    workspace.txt
    preferences.ini
    recent.tsv
    recovery.tsv
    shells/writer.html shells/calc.html shells/impress.html
    shells/draw.html  shells/math.html  shells/base.html
    exports/project-report.odt exports/budget-spreadsheet.ods
    exports/blank-presentation.odp exports/blank-drawing.odg
    exports/blank-database.odb exports/formula.mathml
)
echo
echo "--- Verifying generated profile artifacts ---"
for entry in "${EXPECTED[@]}"; do
    if [[ -s "$PROFILE/$entry" ]]; then
        pass "$entry exists"
    else
        fail "$entry exists" "missing or empty"
    fi
done

# 2. start-center.html lists at least one of the seeded templates.
if grep -q "Blank Writer Document" "$PROFILE/start-center.html"; then
    pass "start-center.html lists templates"
else
    fail "start-center.html lists templates" "Blank Writer Document missing"
fi

# 3. Each window shell HTML embeds an inline SVG preview tile.
for shell in shells/writer.html shells/calc.html shells/impress.html shells/draw.html shells/math.html shells/base.html; do
    if grep -q "<svg" "$PROFILE/$shell"; then
        pass "$shell embeds SVG preview"
    else
        fail "$shell embeds SVG preview" "no <svg> tag"
    fi
done

# 4. Each native ODF format opens in real LibreOffice and converts to PDF.
echo
echo "--- Round-tripping native ODF artifacts through soffice ---"
RT_DIR="$WORK_DIR/rt"
mkdir -p "$RT_DIR"
for f in exports/project-report.odt exports/budget-spreadsheet.ods exports/blank-presentation.odp exports/blank-drawing.odg exports/blank-database.odb; do
    if $SOFFICE --headless --convert-to pdf --outdir "$RT_DIR" "$PROFILE/$f" >/dev/null 2>&1; then
        base="$(basename "${f%.*}")"
        if [[ -s "$RT_DIR/${base}.pdf" ]]; then
            pass "$f → soffice pdf"
        else
            fail "$f → soffice pdf" "empty output"
        fi
    else
        fail "$f → soffice pdf" "soffice conversion failed"
    fi
done

# Recovery snapshots also need to be valid LO documents.
echo
echo "--- Validating autosave snapshots ---"
for f in recovery/win-1-project-report.odt recovery/win-2-budget-spreadsheet.ods recovery/win-3-blank-presentation.odp recovery/win-4-blank-drawing.odg recovery/win-6-blank-database.odb; do
    if [[ -s "$PROFILE/$f" ]] && $SOFFICE --headless --convert-to pdf --outdir "$RT_DIR" "$PROFILE/$f" >/dev/null 2>&1; then
        base="$(basename "${f%.*}")"
        if [[ -s "$RT_DIR/${base}.pdf" ]]; then
            pass "$f → soffice pdf"
        else
            fail "$f → soffice pdf" "empty output"
        fi
    else
        fail "$f → soffice pdf" "missing or unreadable"
    fi
done

# 5. Writer round-trip — assert the appended desktop text survives.
echo
echo "--- Writer ODT content round-trip ---"
TXT_DIR="$WORK_DIR/txt"
mkdir -p "$TXT_DIR"
if $SOFFICE --headless --convert-to txt --outdir "$TXT_DIR" "$PROFILE/exports/project-report.odt" >/dev/null 2>&1; then
    if grep -q "This desktop shell runs fully in Rust" "$TXT_DIR/project-report.txt"; then
        pass "Writer ODT preserves desktop-appended text"
    else
        fail "Writer ODT preserves desktop-appended text" "missing"
    fi
    if grep -q "Executive Summary" "$TXT_DIR/project-report.txt"; then
        pass "Writer ODT preserves template heading"
    else
        fail "Writer ODT preserves template heading" "missing"
    fi
else
    fail "Writer ODT round-trip" "soffice conversion failed"
fi

# 6. Calc round-trip — assert the desktop-set Forecast formula evaluates.
echo
echo "--- Calc ODS formula round-trip ---"
CSV_DIR="$WORK_DIR/csv"
mkdir -p "$CSV_DIR"
if $SOFFICE --headless --calc --convert-to csv --outdir "$CSV_DIR" "$PROFILE/exports/budget-spreadsheet.ods" >/dev/null 2>&1; then
    if grep -q "^Forecast,600" "$CSV_DIR/budget-spreadsheet.csv"; then
        pass "Calc ODS evaluates desktop-recorded Forecast formula = 600"
    else
        fail "Calc ODS Forecast formula" "csv did not contain 'Forecast,600'"
    fi
else
    fail "Calc ODS csv round-trip" "soffice conversion failed"
fi

echo
echo "============================================"
echo " Results: $PASS passed, $FAIL failed (of $((PASS + FAIL)))"
echo "============================================"
if [[ $FAIL -gt 0 ]]; then
    echo -e "$ERRORS"
    exit 1
fi
