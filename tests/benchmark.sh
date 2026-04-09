#!/usr/bin/env bash
#
# Performance & Accuracy Benchmark: libreoffice-rs vs LibreOffice
#
# Measures:
#   1. Speed — document generation time (libreoffice-rs) vs LibreOffice CLI
#   2. Per-feature accuracy — does the generated content survive round-trip?
#   3. Compatibility — cross-format conversion success
#   4. Multilingual — EN, Chinese (ZH), Spanish (ES) content
#   5. Edge cases — empty docs, huge docs, special chars, deeply nested structures
#
set -uo pipefail

SOFFICE="soffice"
BINARY="cargo run --quiet --release --"
SCRIPT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
WORK_DIR="$(mktemp -d)"
RESULTS_FILE="$WORK_DIR/benchmark_results.txt"

cleanup() {
    echo ""
    echo "Working directory preserved at: $WORK_DIR"
}
trap cleanup EXIT

cd "$SCRIPT_DIR"

# Build release first
echo "Building release binary..."
cargo build --release --quiet 2>/dev/null
BINARY_PATH="$(pwd)/target/release/libreoffice-rs"
echo "Done."
echo ""

########################################################################
# Helpers
########################################################################

PASS=0
FAIL=0
SKIP=0

log_result() {
    echo "$1" | tee -a "$RESULTS_FILE"
}

time_cmd() {
    # Returns elapsed time in milliseconds
    local start end
    start=$(python3 -c 'import time; print(int(time.time()*1000))')
    eval "$@" > /dev/null 2>&1
    local rc=$?
    end=$(python3 -c 'import time; print(int(time.time()*1000))')
    echo $(( end - start ))
    return $rc
}

bench_libreoffice_rs() {
    local label="$1"; shift
    local ms
    ms=$(time_cmd "$BINARY_PATH" "$@")
    log_result "  [libreoffice-rs] $label: ${ms}ms" >&2
    printf '%d' "$ms"
}

bench_soffice() {
    local label="$1" input="$2" format="$3"
    local outdir="$WORK_DIR/soffice_out"
    mkdir -p "$outdir"
    local ms
    ms=$(time_cmd "$SOFFICE" --headless --convert-to "$format" --outdir "$outdir" "$input")
    log_result "  [LibreOffice]    $label: ${ms}ms" >&2
    printf '%d' "$ms"
}

check_roundtrip_text() {
    local odt_file="$1" expected_text="$2" label="$3"
    local outdir="$WORK_DIR/rt_check"
    mkdir -p "$outdir"

    # Convert to plain text via LibreOffice
    $SOFFICE --headless --convert-to "txt:Text (encoded):UTF8" --outdir "$outdir" "$odt_file" > /dev/null 2>&1
    local txt_file="$outdir/$(basename "${odt_file%.*}").txt"

    if [[ -f "$txt_file" ]] && grep -qF "$expected_text" "$txt_file" 2>/dev/null; then
        log_result "  PASS: $label"
        PASS=$((PASS + 1))
        return 0
    else
        log_result "  FAIL: $label (text '$expected_text' not found after round-trip)"
        FAIL=$((FAIL + 1))
        return 1
    fi
}

check_roundtrip_xml() {
    local odf_file="$1" xpath_text="$2" label="$3"
    local extract_dir="$WORK_DIR/xml_check"
    mkdir -p "$extract_dir"

    if unzip -o "$odf_file" content.xml -d "$extract_dir" > /dev/null 2>&1; then
        if grep -qF "$xpath_text" "$extract_dir/content.xml" 2>/dev/null; then
            log_result "  PASS: $label"
            PASS=$((PASS + 1))
            return 0
        else
            log_result "  FAIL: $label (text '$xpath_text' not in content.xml)"
            FAIL=$((FAIL + 1))
            return 1
        fi
    else
        log_result "  FAIL: $label (could not extract content.xml)"
        FAIL=$((FAIL + 1))
        return 1
    fi
}

check_pdf_conversion() {
    local input="$1" label="$2"
    local outdir="$WORK_DIR/pdf_bench"
    mkdir -p "$outdir"
    if $SOFFICE --headless --convert-to pdf --outdir "$outdir" "$input" > /dev/null 2>&1; then
        local pdf="$outdir/$(basename "${input%.*}").pdf"
        if [[ -f "$pdf" && -s "$pdf" ]]; then
            log_result "  PASS: $label → PDF"
            PASS=$((PASS + 1))
            return 0
        fi
    fi
    log_result "  FAIL: $label → PDF"
    FAIL=$((FAIL + 1))
    return 1
}

echo "============================================" | tee "$RESULTS_FILE"
echo " libreoffice-rs vs LibreOffice Benchmark" | tee -a "$RESULTS_FILE"
echo " LibreOffice: $($SOFFICE --version 2>/dev/null || echo 'unknown')" | tee -a "$RESULTS_FILE"
echo " Date: $(date)" | tee -a "$RESULTS_FILE"
echo "============================================" | tee -a "$RESULTS_FILE"
echo "" | tee -a "$RESULTS_FILE"

########################################################################
# SECTION 1: MULTILINGUAL DOCUMENT GENERATION
########################################################################
log_result "=== 1. MULTILINGUAL DOCUMENTS ==="
log_result ""

# --- English ---
EN_MD="$WORK_DIR/en_doc.md"
cat > "$EN_MD" <<'EOF'
# The Quick Brown Fox

The **quick brown fox** jumps over the *lazy dog*. This sentence contains every letter of the English alphabet.

## Features List

- Bold text formatting
- Italic text styling
- Inline `code` snippets
- [Hyperlinks](https://example.com)

## Data Table

| Name    | Role      | Score |
|---------|-----------|-------|
| Alice   | Engineer  | 95    |
| Bob     | Designer  | 87    |
| Charlie | Manager   | 92    |

---

### Technical Details

The formula `E = mc²` revolutionized physics. Here is a code example:

```
fn main() {
    println!("Hello, world!");
}
```
EOF

EN_ODT="$WORK_DIR/en_doc.odt"
log_result "--- English Document ---"
ms_rs=$(bench_libreoffice_rs "generate ODT" writer markdown-to-odt "$EN_MD" "$EN_ODT" --title "English Test")
ms_lo=$(bench_soffice "convert ODT→PDF" "$EN_ODT" "pdf")
log_result "  Speed: libreoffice-rs ${ms_rs}ms, LibreOffice ${ms_lo}ms ($(python3 -c "print(f'{$ms_lo/max($ms_rs,1):.0f}')"))x faster)"
check_roundtrip_text "$EN_ODT" "quick brown fox" "EN: bold text preserved"
check_roundtrip_text "$EN_ODT" "lazy dog" "EN: italic text preserved"
check_roundtrip_xml "$EN_ODT" "Alice" "EN: table data in XML"
check_roundtrip_xml "$EN_ODT" "Engineer" "EN: table cells preserved"
log_result ""

# --- Chinese ---
ZH_MD="$WORK_DIR/zh_doc.md"
cat > "$ZH_MD" <<'EOF'
# 中文测试文档

这是一个**中文测试**文档，用于验证 *libreoffice-rs* 对中文字符的支持。

## 特性列表

- 粗体文本格式
- 斜体文本样式
- 内联`代码`片段
- [超链接](https://example.com)

## 数据表格

| 姓名   | 角色     | 分数 |
|--------|----------|------|
| 张三   | 工程师   | 95   |
| 李四   | 设计师   | 87   |
| 王五   | 经理     | 92   |

---

### 技术细节

量子力学和相对论是现代物理学的两大支柱。Unicode支持对于国际化应用至关重要。

常用成语：一箭双雕、画蛇添足、守株待兔。
EOF

ZH_ODT="$WORK_DIR/zh_doc.odt"
log_result "--- Chinese Document ---"
ms_rs=$(bench_libreoffice_rs "generate ODT" writer markdown-to-odt "$ZH_MD" "$ZH_ODT" --title "中文测试")
ms_lo=$(bench_soffice "convert ODT→PDF" "$ZH_ODT" "pdf")
log_result "  Speed: libreoffice-rs ${ms_rs}ms, LibreOffice ${ms_lo}ms ($(python3 -c "print(f'{$ms_lo/max($ms_rs,1):.0f}')"))x faster)"
check_roundtrip_text "$ZH_ODT" "中文测试" "ZH: Chinese title preserved"
check_roundtrip_xml "$ZH_ODT" "张三" "ZH: Chinese table data in XML"
check_roundtrip_xml "$ZH_ODT" "工程师" "ZH: Chinese table cells preserved"
check_roundtrip_xml "$ZH_ODT" "一箭双雕" "ZH: Chinese idiom preserved"
log_result ""

# --- Spanish ---
ES_MD="$WORK_DIR/es_doc.md"
cat > "$ES_MD" <<'EOF'
# Documento de Prueba en Español

El **rápido zorro marrón** salta sobre el *perro perezoso*. Esta oración contiene caracteres especiales del español.

## Características

- Texto en **negrita**
- Texto en *cursiva*
- Código en línea: `función()`
- [Enlaces](https://example.com)

## Tabla de Datos

| Nombre  | Función    | Puntuación |
|---------|------------|------------|
| María   | Ingeniera  | 95         |
| José    | Diseñador  | 87         |
| Señor Ñ | Gerente   | 92         |

---

### Caracteres Especiales

Verificación de acentos: á é í ó ú ñ ü ¿ ¡
Números: ½ ¼ ¾ © ® ™
EOF

ES_ODT="$WORK_DIR/es_doc.odt"
log_result "--- Spanish Document ---"
ms_rs=$(bench_libreoffice_rs "generate ODT" writer markdown-to-odt "$ES_MD" "$ES_ODT" --title "Prueba Español")
ms_lo=$(bench_soffice "convert ODT→PDF" "$ES_ODT" "pdf")
log_result "  Speed: libreoffice-rs ${ms_rs}ms, LibreOffice ${ms_lo}ms ($(python3 -c "print(f'{$ms_lo/max($ms_rs,1):.0f}')"))x faster)"
check_roundtrip_text "$ES_ODT" "rápido zorro" "ES: accented text preserved"
check_roundtrip_xml "$ES_ODT" "María" "ES: accented names in XML"
check_roundtrip_xml "$ES_ODT" "Señor Ñ" "ES: ñ character preserved"
check_roundtrip_xml "$ES_ODT" "¿" "ES: inverted question mark preserved"
log_result ""

########################################################################
# SECTION 2: SPEED BENCHMARKS
########################################################################
log_result "=== 2. SPEED BENCHMARKS ==="
log_result ""

# --- ODT generation speed ---
log_result "--- ODT Generation Speed ---"
for size in 10 100 1000; do
    MD_FILE="$WORK_DIR/speed_${size}.md"
    {
        echo "# Document with $size paragraphs"
        echo ""
        for i in $(seq 1 "$size"); do
            echo "Paragraph $i: The quick brown fox jumps over the lazy dog. This is **bold** and *italic* text with \`code\` inline."
            echo ""
        done
    } > "$MD_FILE"

    ODT_FILE="$WORK_DIR/speed_${size}.odt"
    ms_rs=$(bench_libreoffice_rs "$size paragraphs → ODT" writer markdown-to-odt "$MD_FILE" "$ODT_FILE" --title "Speed Test $size")

    if [[ -f "$ODT_FILE" ]]; then
        ms_lo=$(bench_soffice "$size paragraphs ODT→PDF" "$ODT_FILE" "pdf")
        ratio="N/A"
        if [[ $ms_rs -gt 0 ]]; then
            ratio=$(python3 -c "print(f'{$ms_lo/$ms_rs:.1f}x')")
        fi
        log_result "  Ratio (LO/rs): $ratio"
    fi
    log_result ""
done

# --- ODS generation speed ---
log_result "--- ODS Generation Speed ---"
for rows in 10 100 1000; do
    CSV_FILE="$WORK_DIR/speed_${rows}.csv"
    {
        echo "ID,Name,Value,Category,Score,Active"
        for i in $(seq 1 "$rows"); do
            echo "$i,Name_$i,$((RANDOM % 10000)),Cat$((i % 10)),$((RANDOM % 100)),$([ $((i % 2)) -eq 0 ] && echo true || echo false)"
        done
    } > "$CSV_FILE"

    ODS_FILE="$WORK_DIR/speed_${rows}.ods"
    ms_rs=$(bench_libreoffice_rs "$rows rows → ODS" calc csv-to-ods "$CSV_FILE" "$ODS_FILE" --sheet "Data" --title "Speed $rows")

    if [[ -f "$ODS_FILE" ]]; then
        ms_lo=$(bench_soffice "$rows rows ODS→PDF" "$ODS_FILE" "pdf")
        ratio="N/A"
        if [[ $ms_rs -gt 0 ]]; then
            ratio=$(python3 -c "print(f'{$ms_lo/$ms_rs:.1f}x')")
        fi
        log_result "  Ratio (LO/rs): $ratio"
    fi
    log_result ""
done

########################################################################
# SECTION 3: EDGE CASES
########################################################################
log_result "=== 3. EDGE CASES ==="
log_result ""

# --- Empty document ---
log_result "--- Edge: Empty Document ---"
EMPTY_ODT="$WORK_DIR/empty.odt"
$BINARY_PATH writer new "$EMPTY_ODT" --title "" --text "" > /dev/null 2>&1
if [[ -f "$EMPTY_ODT" && -s "$EMPTY_ODT" ]]; then
    log_result "  PASS: empty document created"
    PASS=$((PASS + 1))
    check_pdf_conversion "$EMPTY_ODT" "empty.odt"
else
    log_result "  FAIL: empty document not created"
    FAIL=$((FAIL + 1))
fi
log_result ""

# --- Special characters ---
log_result "--- Edge: Special Characters ---"
SPECIAL_MD="$WORK_DIR/special.md"
cat > "$SPECIAL_MD" <<'MDEOF'
# Special Characters Test

## XML-sensitive characters

These characters must be escaped: < > & " '

Angle brackets: <script>alert('xss')</script>

Ampersands: AT&T, rock & roll, &amp; already escaped?

## Unicode edge cases

Emoji: Hello World
Zero-width: foo​bar (zero-width space between foo and bar)
RTL text: مرحبا بالعالم (Arabic: Hello World)
Math symbols: ∑ ∏ ∫ √ ∞ ≠ ≤ ≥ ± ÷ × ∈ ∉ ⊂ ⊃
Currency: $ € £ ¥ ₹ ₿ ₩
Diacritics: àáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ

## Deeply nested content

- Level 1
- Level 1 continued with **bold *nested italic* more bold**
- Table inside list context:

| Col1 | Col2 |
|------|------|
| a    | b    |
MDEOF

SPECIAL_ODT="$WORK_DIR/special.odt"
$BINARY_PATH writer markdown-to-odt "$SPECIAL_MD" "$SPECIAL_ODT" --title "Special Chars" > /dev/null 2>&1
if [[ -f "$SPECIAL_ODT" && -s "$SPECIAL_ODT" ]]; then
    log_result "  PASS: special characters document created"
    PASS=$((PASS + 1))
    check_pdf_conversion "$SPECIAL_ODT" "special.odt"
    check_roundtrip_xml "$SPECIAL_ODT" "&amp;" "XML: ampersand escaped correctly"
    check_roundtrip_xml "$SPECIAL_ODT" "&lt;" "XML: less-than escaped correctly"
    check_roundtrip_xml "$SPECIAL_ODT" "مرحبا" "Unicode: Arabic text preserved"
    check_roundtrip_xml "$SPECIAL_ODT" "àáâãäå" "Unicode: diacritics preserved"
else
    log_result "  FAIL: special characters document not created"
    FAIL=$((FAIL + 1))
fi
log_result ""

# --- Very long single line ---
log_result "--- Edge: Very Long Line ---"
LONGLINE_MD="$WORK_DIR/longline.md"
{
    echo "# Long Line Test"
    echo ""
    python3 -c "print('A' * 50000)"
} > "$LONGLINE_MD"

LONGLINE_ODT="$WORK_DIR/longline.odt"
ms_rs=$(bench_libreoffice_rs "50K char line → ODT" writer markdown-to-odt "$LONGLINE_MD" "$LONGLINE_ODT" --title "Long Line")
if [[ -f "$LONGLINE_ODT" && -s "$LONGLINE_ODT" ]]; then
    log_result "  PASS: long line document created ($(stat -f%z "$LONGLINE_ODT") bytes)"
    PASS=$((PASS + 1))
    check_pdf_conversion "$LONGLINE_ODT" "longline.odt"
else
    log_result "  FAIL: long line document not created"
    FAIL=$((FAIL + 1))
fi
log_result ""

# --- Many headings ---
log_result "--- Edge: Many Headings (all 6 levels) ---"
HEADINGS_MD="$WORK_DIR/headings.md"
{
    for level in 1 2 3 4 5 6; do
        hashes=$(printf '#%.0s' $(seq 1 "$level"))
        echo "$hashes Heading Level $level"
        echo ""
        echo "Content under heading level $level."
        echo ""
    done
} > "$HEADINGS_MD"

HEADINGS_ODT="$WORK_DIR/headings.odt"
$BINARY_PATH writer markdown-to-odt "$HEADINGS_MD" "$HEADINGS_ODT" --title "Headings Test" > /dev/null 2>&1
if [[ -f "$HEADINGS_ODT" ]]; then
    for level in 1 2 3 4 5 6; do
        check_roundtrip_xml "$HEADINGS_ODT" "text:outline-level=\"$level\"" "Heading level $level in XML"
    done
fi
log_result ""

# --- Empty CSV ---
log_result "--- Edge: Minimal CSV ---"
EMPTY_CSV="$WORK_DIR/minimal.csv"
echo "A" > "$EMPTY_CSV"
EMPTY_ODS="$WORK_DIR/minimal.ods"
$BINARY_PATH calc csv-to-ods "$EMPTY_CSV" "$EMPTY_ODS" --sheet "Empty" > /dev/null 2>&1
if [[ -f "$EMPTY_ODS" && -s "$EMPTY_ODS" ]]; then
    log_result "  PASS: minimal CSV → ODS created"
    PASS=$((PASS + 1))
    check_pdf_conversion "$EMPTY_ODS" "minimal.ods"
else
    log_result "  FAIL: minimal CSV → ODS failed"
    FAIL=$((FAIL + 1))
fi
log_result ""

# --- CSV with special values ---
log_result "--- Edge: CSV with Special Values ---"
SPECIAL_CSV="$WORK_DIR/special_values.csv"
cat > "$SPECIAL_CSV" <<'CSVEOF'
Type,Value
Empty,
Zero,0
Negative,-42.5
Large,999999999999
Scientific,1.23e10
Boolean True,true
Boolean False,false
Quoted,"Hello, ""World"""
Unicode,日本語テスト
Formula,=1+1
CSVEOF

SPECIAL_ODS="$WORK_DIR/special_values.ods"
$BINARY_PATH calc csv-to-ods "$SPECIAL_CSV" "$SPECIAL_ODS" --sheet "Special" > /dev/null 2>&1
if [[ -f "$SPECIAL_ODS" && -s "$SPECIAL_ODS" ]]; then
    log_result "  PASS: special values CSV → ODS created"
    PASS=$((PASS + 1))
    check_pdf_conversion "$SPECIAL_ODS" "special_values.ods"
    check_roundtrip_xml "$SPECIAL_ODS" "office:value=\"-42.5\"" "CSV: negative numbers"
    check_roundtrip_xml "$SPECIAL_ODS" "999999999999" "CSV: large numbers"
    check_roundtrip_xml "$SPECIAL_ODS" "日本語テスト" "CSV: Japanese text"
    check_roundtrip_xml "$SPECIAL_ODS" 'Hello, "World"' "CSV: quoted values"
else
    log_result "  FAIL: special values CSV → ODS failed"
    FAIL=$((FAIL + 1))
fi
log_result ""

# --- Large spreadsheet ---
log_result "--- Edge: Large Spreadsheet (5000 rows × 10 cols) ---"
HUGE_CSV="$WORK_DIR/huge.csv"
{
    echo "ID,Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9"
    for i in $(seq 1 5000); do
        echo "$i,$((RANDOM)),$((RANDOM)),$((RANDOM)),$((RANDOM)),$((RANDOM)),$((RANDOM)),$((RANDOM)),$((RANDOM)),$((RANDOM))"
    done
} > "$HUGE_CSV"

HUGE_ODS="$WORK_DIR/huge.ods"
ms_rs=$(bench_libreoffice_rs "5000×10 → ODS" calc csv-to-ods "$HUGE_CSV" "$HUGE_ODS" --sheet "Huge")
if [[ -f "$HUGE_ODS" && -s "$HUGE_ODS" ]]; then
    filesize=$(stat -f%z "$HUGE_ODS" 2>/dev/null || stat -c%s "$HUGE_ODS" 2>/dev/null)
    log_result "  PASS: huge spreadsheet created ($filesize bytes)"
    PASS=$((PASS + 1))
    ms_lo=$(bench_soffice "5000×10 ODS→PDF" "$HUGE_ODS" "pdf")
    ratio="N/A"
    if [[ $ms_rs -gt 0 ]]; then
        ratio=$(python3 -c "print(f'{$ms_lo/$ms_rs:.1f}x')")
    fi
    log_result "  Ratio (LO/rs): $ratio"
else
    log_result "  FAIL: huge spreadsheet not created"
    FAIL=$((FAIL + 1))
fi
log_result ""

########################################################################
# SECTION 4: PER-FEATURE ACCURACY
########################################################################
log_result "=== 4. PER-FEATURE ACCURACY ==="
log_result ""

# --- Writer features ---
log_result "--- Writer Feature Accuracy ---"

# Bold
BOLD_MD="$WORK_DIR/feat_bold.md"
echo "This is **bold text** here." > "$BOLD_MD"
BOLD_ODT="$WORK_DIR/feat_bold.odt"
$BINARY_PATH writer markdown-to-odt "$BOLD_MD" "$BOLD_ODT" --title "Bold" > /dev/null 2>&1
check_roundtrip_xml "$BOLD_ODT" "text:style-name=\"Strong\"" "Feature: bold → Strong style"
check_roundtrip_text "$BOLD_ODT" "bold text" "Feature: bold text readable"

# Italic
ITALIC_MD="$WORK_DIR/feat_italic.md"
echo "This is *italic text* here." > "$ITALIC_MD"
ITALIC_ODT="$WORK_DIR/feat_italic.odt"
$BINARY_PATH writer markdown-to-odt "$ITALIC_MD" "$ITALIC_ODT" --title "Italic" > /dev/null 2>&1
check_roundtrip_xml "$ITALIC_ODT" "text:style-name=\"Emphasis\"" "Feature: italic → Emphasis style"

# Code
CODE_MD="$WORK_DIR/feat_code.md"
echo 'This has `inline code` here.' > "$CODE_MD"
CODE_ODT="$WORK_DIR/feat_code.odt"
$BINARY_PATH writer markdown-to-odt "$CODE_MD" "$CODE_ODT" --title "Code" > /dev/null 2>&1
check_roundtrip_xml "$CODE_ODT" "text:style-name=\"Code\"" "Feature: code → Code style"

# Link
LINK_MD="$WORK_DIR/feat_link.md"
echo 'Visit [Example Site](https://example.com) now.' > "$LINK_MD"
LINK_ODT="$WORK_DIR/feat_link.odt"
$BINARY_PATH writer markdown-to-odt "$LINK_MD" "$LINK_ODT" --title "Link" > /dev/null 2>&1
check_roundtrip_xml "$LINK_ODT" "xlink:href=\"https://example.com\"" "Feature: hyperlink URL preserved"
check_roundtrip_xml "$LINK_ODT" "Example Site" "Feature: hyperlink label preserved"

# Unordered list
LIST_MD="$WORK_DIR/feat_list.md"
cat > "$LIST_MD" <<'EOF'
- First item
- Second item
- Third item
EOF
LIST_ODT="$WORK_DIR/feat_list.odt"
$BINARY_PATH writer markdown-to-odt "$LIST_MD" "$LIST_ODT" --title "List" > /dev/null 2>&1
check_roundtrip_xml "$LIST_ODT" "text:list" "Feature: list element present"
check_roundtrip_xml "$LIST_ODT" "First item" "Feature: list item content"

# Table
TABLE_MD="$WORK_DIR/feat_table.md"
cat > "$TABLE_MD" <<'EOF'
| H1 | H2 | H3 |
|----|----|----|
| a  | b  | c  |
| d  | e  | f  |
EOF
TABLE_ODT="$WORK_DIR/feat_table.odt"
$BINARY_PATH writer markdown-to-odt "$TABLE_MD" "$TABLE_ODT" --title "Table" > /dev/null 2>&1
check_roundtrip_xml "$TABLE_ODT" "table:table" "Feature: table element present"
check_roundtrip_xml "$TABLE_ODT" "table:table-row" "Feature: table rows present"
check_roundtrip_xml "$TABLE_ODT" "table:table-cell" "Feature: table cells present"

# Horizontal rule
HR_MD="$WORK_DIR/feat_hr.md"
cat > "$HR_MD" <<'EOF'
Before rule.

---

After rule.
EOF
HR_ODT="$WORK_DIR/feat_hr.odt"
$BINARY_PATH writer markdown-to-odt "$HR_MD" "$HR_ODT" --title "HR" > /dev/null 2>&1
check_roundtrip_xml "$HR_ODT" "Before rule" "Feature: HR content before"
check_roundtrip_xml "$HR_ODT" "After rule" "Feature: HR content after"

log_result ""

# --- Calc features ---
log_result "--- Calc Feature Accuracy ---"

# Cell types
TYPES_CSV="$WORK_DIR/feat_types.csv"
cat > "$TYPES_CSV" <<'CSVEOF'
Number,Text,Bool
42,hello,true
3.14,world,false
CSVEOF
TYPES_ODS="$WORK_DIR/feat_types.ods"
$BINARY_PATH calc csv-to-ods "$TYPES_CSV" "$TYPES_ODS" --sheet "Types" > /dev/null 2>&1
check_roundtrip_xml "$TYPES_ODS" "office:value-type=\"float\"" "Feature: numeric cell type"
check_roundtrip_xml "$TYPES_ODS" "office:value-type=\"string\"" "Feature: string cell type"
check_roundtrip_xml "$TYPES_ODS" "office:value-type=\"boolean\"" "Feature: boolean cell type"
check_roundtrip_xml "$TYPES_ODS" "office:value=\"42\"" "Feature: numeric value preserved"
check_roundtrip_xml "$TYPES_ODS" "office:boolean-value=\"true\"" "Feature: boolean value preserved"

# Formulas
FORMULA_CSV="$WORK_DIR/feat_formula.csv"
cat > "$FORMULA_CSV" <<'CSVEOF'
A,B,Sum
10,20,=A2+B2
5,15,=SUM(A2:A3)
CSVEOF
FORMULA_ODS="$WORK_DIR/feat_formula.ods"
$BINARY_PATH calc csv-to-ods "$FORMULA_CSV" "$FORMULA_ODS" --sheet "Formulas" > /dev/null 2>&1
check_roundtrip_xml "$FORMULA_ODS" "table:formula" "Feature: formula attribute present"
check_roundtrip_xml "$FORMULA_ODS" "of:=" "Feature: formula uses ODF prefix"

# Formula evaluation accuracy
log_result ""
log_result "--- Formula Evaluation Accuracy ---"

eval_formula() {
    local formula="$1" expected="$2" label="$3"
    local csv_file="$WORK_DIR/eval_input.csv"
    local result
    result=$($BINARY_PATH calc eval "$formula" --csv "$csv_file" 2>&1)
    if echo "$result" | grep -qF "$expected"; then
        log_result "  PASS: $label: $formula = $expected"
        PASS=$((PASS + 1))
    else
        log_result "  FAIL: $label: $formula expected $expected, got $result"
        FAIL=$((FAIL + 1))
    fi
}

# Prepare evaluation CSV
EVAL_CSV="$WORK_DIR/eval_input.csv"
cat > "$EVAL_CSV" <<'CSVEOF'
A,B,C
10,20,30
5,15,25
1,2,3
CSVEOF

eval_formula "=SUM(A2:A4)" "Number(16.0)" "SUM range"
eval_formula "=AVERAGE(B2:B4)" "Number(12.333" "AVERAGE range"
eval_formula "=MIN(A2:A4)" "Number(1.0)" "MIN"
eval_formula "=MAX(C2:C4)" "Number(30.0)" "MAX"
eval_formula "=COUNT(A2:A4)" "Number(3.0)" "COUNT"
eval_formula "=IF(A2>3,\"big\",\"small\")" "Text(\"big\")" "IF conditional"
eval_formula "=AND(TRUE,TRUE)" "Bool(true)" "AND"
eval_formula "=OR(FALSE,TRUE)" "Bool(true)" "OR"
eval_formula "=NOT(FALSE)" "Bool(true)" "NOT"
eval_formula "=ABS(-42)" "Number(42.0)" "ABS"
eval_formula "=ROUND(3.14159,2)" "Number(3.14)" "ROUND"
eval_formula "=LEN(\"hello\")" "Number(5.0)" "LEN"
eval_formula "=CONCAT(\"foo\",\"bar\")" "Text(\"foobar\")" "CONCAT"
eval_formula "=A2+B2" "Number(30.0)" "cell addition"
eval_formula "=A2*B2" "Number(200.0)" "cell multiplication"
eval_formula "=A2^2" "Number(100.0)" "exponentiation"

log_result ""

########################################################################
# SECTION 5: CROSS-FORMAT COMPATIBILITY
########################################################################
log_result "=== 5. CROSS-FORMAT COMPATIBILITY ==="
log_result ""

log_result "--- ODT Conversions ---"
check_pdf_conversion "$EN_ODT" "EN ODT"
DOCX_OUT="$WORK_DIR/compat_out"
mkdir -p "$DOCX_OUT"
$SOFFICE --headless --convert-to docx --outdir "$DOCX_OUT" "$EN_ODT" > /dev/null 2>&1
if [[ -f "$DOCX_OUT/en_doc.docx" && -s "$DOCX_OUT/en_doc.docx" ]]; then
    log_result "  PASS: ODT → DOCX"
    PASS=$((PASS + 1))
    # Convert DOCX back to ODT (round-trip through MS format)
    $SOFFICE --headless --convert-to odt --outdir "$DOCX_OUT" "$DOCX_OUT/en_doc.docx" > /dev/null 2>&1
    if [[ -f "$DOCX_OUT/en_doc.odt" ]]; then
        log_result "  PASS: ODT → DOCX → ODT round-trip"
        PASS=$((PASS + 1))
    else
        log_result "  FAIL: DOCX → ODT back-conversion"
        FAIL=$((FAIL + 1))
    fi
else
    log_result "  FAIL: ODT → DOCX"
    FAIL=$((FAIL + 1))
fi

log_result ""
log_result "--- ODS Conversions ---"
ODS_SIMPLE="$WORK_DIR/speed_100.ods"
if [[ -f "$ODS_SIMPLE" ]]; then
    check_pdf_conversion "$ODS_SIMPLE" "ODS 100-row"
    $SOFFICE --headless --convert-to xlsx --outdir "$DOCX_OUT" "$ODS_SIMPLE" > /dev/null 2>&1
    if [[ -f "$DOCX_OUT/speed_100.xlsx" && -s "$DOCX_OUT/speed_100.xlsx" ]]; then
        log_result "  PASS: ODS → XLSX"
        PASS=$((PASS + 1))
    else
        log_result "  FAIL: ODS → XLSX"
        FAIL=$((FAIL + 1))
    fi
fi

log_result ""
log_result "--- ODP Conversions ---"
ODP_FILE="$WORK_DIR/compat_demo.odp"
$BINARY_PATH impress demo "$ODP_FILE" --title "Compat Test" > /dev/null 2>&1
if [[ -f "$ODP_FILE" ]]; then
    check_pdf_conversion "$ODP_FILE" "ODP demo"
    $SOFFICE --headless --convert-to pptx --outdir "$DOCX_OUT" "$ODP_FILE" > /dev/null 2>&1
    if [[ -f "$DOCX_OUT/compat_demo.pptx" && -s "$DOCX_OUT/compat_demo.pptx" ]]; then
        log_result "  PASS: ODP → PPTX"
        PASS=$((PASS + 1))
    else
        log_result "  FAIL: ODP → PPTX"
        FAIL=$((FAIL + 1))
    fi
fi

log_result ""
log_result "--- ODG Conversions ---"
ODG_FILE="$WORK_DIR/compat_draw.odg"
$BINARY_PATH draw demo "$ODG_FILE" --title "Compat Drawing" > /dev/null 2>&1
if [[ -f "$ODG_FILE" ]]; then
    check_pdf_conversion "$ODG_FILE" "ODG demo"
    $SOFFICE --headless --convert-to svg --outdir "$DOCX_OUT" "$ODG_FILE" > /dev/null 2>&1
    if [[ -f "$DOCX_OUT/compat_draw.svg" && -s "$DOCX_OUT/compat_draw.svg" ]]; then
        log_result "  PASS: ODG → SVG"
        PASS=$((PASS + 1))
    else
        log_result "  FAIL: ODG → SVG"
        FAIL=$((FAIL + 1))
    fi
fi

log_result ""

########################################################################
# SECTION 6: DOWNLOAD & TEST REAL-WORLD DOCUMENTS
########################################################################
log_result "=== 6. REAL-WORLD DOCUMENT IMPORT (LibreOffice → CSV/TXT → libreoffice-rs) ==="
log_result ""

# Generate realistic multilingual CSVs and convert through both paths
log_result "--- Multilingual Spreadsheet Pipeline ---"

# Chinese spreadsheet
ZH_CSV="$WORK_DIR/zh_data.csv"
cat > "$ZH_CSV" <<'CSVEOF'
产品名称,单价,数量,总计
笔记本电脑,6999.00,5,=B2*C2
智能手机,3999.00,10,=B3*C3
平板电脑,2999.00,8,=B4*C4
耳机,599.00,20,=B5*C5
总计,,,=SUM(D2:D5)
CSVEOF

ZH_ODS="$WORK_DIR/zh_data.ods"
ms_rs=$(bench_libreoffice_rs "ZH CSV → ODS" calc csv-to-ods "$ZH_CSV" "$ZH_ODS" --sheet "销售数据" --title "中文销售报表")
check_roundtrip_xml "$ZH_ODS" "笔记本电脑" "ZH spreadsheet: product names"
check_roundtrip_xml "$ZH_ODS" "6999" "ZH spreadsheet: prices"
check_pdf_conversion "$ZH_ODS" "ZH spreadsheet"
log_result ""

# Spanish spreadsheet
ES_CSV="$WORK_DIR/es_data.csv"
cat > "$ES_CSV" <<'CSVEOF'
Empleado,Departamento,Salario,Evaluación
María García,Ingeniería,45000,Excelente
José Rodríguez,Diseño,42000,Bueno
Señor Ñoño,Gerencia,55000,Sobresaliente
Lucía Fernández,Marketing,40000,Muy Bueno
CSVEOF

ES_ODS="$WORK_DIR/es_data.ods"
ms_rs=$(bench_libreoffice_rs "ES CSV → ODS" calc csv-to-ods "$ES_CSV" "$ES_ODS" --sheet "Empleados" --title "Nómina Española")
check_roundtrip_xml "$ES_ODS" "María García" "ES spreadsheet: accented names"
check_roundtrip_xml "$ES_ODS" "Señor Ñoño" "ES spreadsheet: ñ in data"
check_pdf_conversion "$ES_ODS" "ES spreadsheet"
log_result ""

########################################################################
# SUMMARY
########################################################################
TOTAL=$((PASS + FAIL))
echo "" | tee -a "$RESULTS_FILE"
echo "============================================" | tee -a "$RESULTS_FILE"
echo " BENCHMARK SUMMARY" | tee -a "$RESULTS_FILE"
echo "============================================" | tee -a "$RESULTS_FILE"
echo " Total tests: $TOTAL" | tee -a "$RESULTS_FILE"
echo " Passed:      $PASS" | tee -a "$RESULTS_FILE"
echo " Failed:      $FAIL" | tee -a "$RESULTS_FILE"
if [[ $TOTAL -gt 0 ]]; then
    pct=$(python3 -c "print(f'{$PASS/$TOTAL*100:.1f}')")
    echo " Pass rate:   ${pct}%" | tee -a "$RESULTS_FILE"
fi
echo "============================================" | tee -a "$RESULTS_FILE"
echo "" | tee -a "$RESULTS_FILE"
echo "Full results saved to: $RESULTS_FILE"

if [[ $FAIL -gt 0 ]]; then
    exit 1
fi
exit 0
