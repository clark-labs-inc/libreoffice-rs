pub mod html;
pub mod import;
pub mod markdown;
pub mod pdf;
pub mod svg;
pub mod xlsx;

pub use html::to_html;
pub use import::{from_ods_bytes, from_xlsx_bytes, load_bytes};
pub use markdown::to_markdown;
pub use pdf::to_pdf;
pub use svg::render_svg;
pub use xlsx::to_xlsx;

use std::collections::BTreeSet;
use std::path::Path;

use lo_core::{Cell, CellAddr, CellValue, LoError, Result, Sheet, Workbook};

#[derive(Clone, Debug, PartialEq)]
pub enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    Ref(CellAddr),
    Range(CellAddr, CellAddr),
    Func {
        name: String,
        args: Vec<Expr>,
    },
    Unary {
        op: UnaryOp,
        expr: Box<Expr>,
    },
    Binary {
        op: BinaryOp,
        lhs: Box<Expr>,
        rhs: Box<Expr>,
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UnaryOp {
    Plus,
    Minus,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Lte,
    Gt,
    Gte,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    Blank,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
}

impl Value {
    fn as_number(&self) -> Result<f64> {
        match self {
            Self::Blank => Ok(0.0),
            Self::Number(value) => Ok(*value),
            Self::Text(text) => text
                .parse::<f64>()
                .map_err(|_| LoError::Eval(format!("cannot convert text to number: {text}"))),
            Self::Bool(value) => Ok(if *value { 1.0 } else { 0.0 }),
            Self::Error(message) => Err(LoError::Eval(message.clone())),
        }
    }

    fn as_bool(&self) -> Result<bool> {
        match self {
            Self::Blank => Ok(false),
            Self::Number(value) => Ok(*value != 0.0),
            Self::Text(text) => Ok(!text.is_empty()),
            Self::Bool(value) => Ok(*value),
            Self::Error(message) => Err(LoError::Eval(message.clone())),
        }
    }

    fn as_text(&self) -> String {
        match self {
            Self::Blank => String::new(),
            Self::Number(value) => value.to_string(),
            Self::Text(text) => text.clone(),
            Self::Bool(value) => value.to_string(),
            Self::Error(message) => format!("#ERR {message}"),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Number(f64),
    Text(String),
    Ident(String),
    Cell(CellAddr),
    Comma,
    Colon,
    LParen,
    RParen,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Amp,
    Eq,
    Ne,
    Lt,
    Lte,
    Gt,
    Gte,
}

pub fn save_ods(path: impl AsRef<Path>, workbook: &Workbook) -> Result<()> {
    lo_odf::save_spreadsheet_document(path, workbook)
}

/// Render the workbook into bytes for the requested format.
///
/// Supported format strings (case-insensitive): `csv`, `md`, `html`, `svg`, `pdf`,
/// `ods`, `xlsx`. CSV always returns the first sheet.
pub fn save_as(workbook: &Workbook, format: &str) -> Result<Vec<u8>> {
    match format.to_ascii_lowercase().as_str() {
        "csv" => Ok(workbook
            .sheets
            .first()
            .map(sheet_to_csv)
            .unwrap_or_default()
            .into_bytes()),
        "md" | "markdown" => Ok(to_markdown(workbook).into_bytes()),
        "html" => Ok(to_html(workbook).into_bytes()),
        "svg" => {
            let size = lo_core::Size::new(
                lo_core::units::Length::pt(1024.0),
                lo_core::units::Length::pt(768.0),
            );
            Ok(render_svg(workbook, size).into_bytes())
        }
        "pdf" => Ok(to_pdf(workbook)),
        "ods" => {
            let tmp = std::env::temp_dir().join(format!(
                "lo_calc_{}.ods",
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_nanos())
                    .unwrap_or(0)
            ));
            lo_odf::save_spreadsheet_document(&tmp, workbook)?;
            let bytes = std::fs::read(&tmp)?;
            let _ = std::fs::remove_file(&tmp);
            Ok(bytes)
        }
        "xlsx" => to_xlsx(workbook),
        other => Err(LoError::Unsupported(format!(
            "calc format not supported: {other}"
        ))),
    }
}

pub fn parse_formula(input: &str) -> Result<Expr> {
    let stripped = input.strip_prefix('=').unwrap_or(input).trim();
    let tokens = tokenize(stripped)?;
    let mut parser = Parser { tokens, index: 0 };
    let expr = parser.parse_expression()?;
    if parser.index != parser.tokens.len() {
        return Err(LoError::Parse(
            "unexpected tokens after end of formula".to_string(),
        ));
    }
    Ok(expr)
}

pub fn evaluate_formula(formula: &str, sheet: &Sheet) -> Result<Value> {
    let expr = parse_formula(formula)?;
    let mut stack = BTreeSet::new();
    eval_expr(&expr, sheet, &mut stack)
}

pub fn workbook_from_csv(
    title: impl Into<String>,
    sheet_name: &str,
    csv: &str,
) -> Result<Workbook> {
    workbook_from_csv_opts(title, sheet_name, csv, false)
}

pub fn workbook_from_csv_opts(
    title: impl Into<String>,
    sheet_name: &str,
    csv: &str,
    has_header: bool,
) -> Result<Workbook> {
    let rows = parse_csv(csv)?;
    let mut workbook = Workbook::new(title);
    workbook.sheets.clear();
    let mut sheet = Sheet::new(sheet_name);
    sheet.has_header = has_header;
    for (row_idx, row) in rows.iter().enumerate() {
        for (col_idx, value) in row.iter().enumerate() {
            // Header row cells are always strings, regardless of shape.
            let cell_value = if has_header && row_idx == 0 {
                if value.is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::Text(value.clone())
                }
            } else {
                infer_cell_value(value)
            };
            if !matches!(cell_value, CellValue::Empty) {
                sheet.set(CellAddr::new(row_idx as u32, col_idx as u32), cell_value);
            }
        }
    }
    workbook.sheets.push(sheet);
    Ok(workbook)
}

pub fn sheet_to_csv(sheet: &Sheet) -> String {
    let (max_row, max_col) = sheet.max_extent();
    let mut out = String::new();
    for row in 0..=max_row {
        let mut cols = Vec::new();
        for col in 0..=max_col {
            let value = sheet
                .get(CellAddr::new(row, col))
                .map(|cell| render_cell_value(&cell.value))
                .unwrap_or_default();
            cols.push(escape_csv(&value));
        }
        out.push_str(&cols.join(","));
        if row != max_row {
            out.push('\n');
        }
    }
    out
}

pub fn evaluate_sheet_formulas(sheet: &Sheet) -> Result<Vec<(CellAddr, Value)>> {
    let mut out = Vec::new();
    for (addr, cell) in &sheet.cells {
        if let CellValue::Formula(formula) = &cell.value {
            out.push((*addr, evaluate_formula(formula, sheet)?));
        }
    }
    Ok(out)
}

fn render_cell_value(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number(value) => value.to_string(),
        CellValue::Text(value) => value.clone(),
        CellValue::Bool(value) => value.to_string(),
        CellValue::Formula(value) => value.clone(),
        CellValue::Error(value) => format!("#ERR {value}"),
    }
}

fn infer_cell_value(value: &str) -> CellValue {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        CellValue::Empty
    } else if trimmed.starts_with('=') {
        CellValue::Formula(trimmed.to_string())
    } else if trimmed.eq_ignore_ascii_case("true") {
        CellValue::Bool(true)
    } else if trimmed.eq_ignore_ascii_case("false") {
        CellValue::Bool(false)
    } else if let Ok(number) = trimmed.parse::<f64>() {
        CellValue::Number(number)
    } else {
        CellValue::Text(trimmed.to_string())
    }
}

fn tokenize(input: &str) -> Result<Vec<Token>> {
    let chars: Vec<char> = input.chars().collect();
    let mut index = 0usize;
    let mut tokens = Vec::new();

    while index < chars.len() {
        let ch = chars[index];
        if ch.is_whitespace() {
            index += 1;
            continue;
        }

        match ch {
            '(' => {
                tokens.push(Token::LParen);
                index += 1;
            }
            ')' => {
                tokens.push(Token::RParen);
                index += 1;
            }
            ',' => {
                tokens.push(Token::Comma);
                index += 1;
            }
            ':' => {
                tokens.push(Token::Colon);
                index += 1;
            }
            '+' => {
                tokens.push(Token::Plus);
                index += 1;
            }
            '-' => {
                tokens.push(Token::Minus);
                index += 1;
            }
            '*' => {
                tokens.push(Token::Star);
                index += 1;
            }
            '/' => {
                tokens.push(Token::Slash);
                index += 1;
            }
            '^' => {
                tokens.push(Token::Caret);
                index += 1;
            }
            '&' => {
                tokens.push(Token::Amp);
                index += 1;
            }
            '=' => {
                tokens.push(Token::Eq);
                index += 1;
            }
            '<' => {
                if chars.get(index + 1) == Some(&'=') {
                    tokens.push(Token::Lte);
                    index += 2;
                } else if chars.get(index + 1) == Some(&'>') {
                    tokens.push(Token::Ne);
                    index += 2;
                } else {
                    tokens.push(Token::Lt);
                    index += 1;
                }
            }
            '>' => {
                if chars.get(index + 1) == Some(&'=') {
                    tokens.push(Token::Gte);
                    index += 2;
                } else {
                    tokens.push(Token::Gt);
                    index += 1;
                }
            }
            '"' => {
                index += 1;
                let start = index;
                while index < chars.len() && chars[index] != '"' {
                    index += 1;
                }
                if index >= chars.len() {
                    return Err(LoError::Parse("unterminated string literal".to_string()));
                }
                let text: String = chars[start..index].iter().collect();
                tokens.push(Token::Text(text));
                index += 1;
            }
            _ if ch.is_ascii_digit() || ch == '.' => {
                let start = index;
                index += 1;
                while index < chars.len() && (chars[index].is_ascii_digit() || chars[index] == '.')
                {
                    index += 1;
                }
                let value: String = chars[start..index].iter().collect();
                let number = value
                    .parse::<f64>()
                    .map_err(|_| LoError::Parse(format!("invalid number literal: {value}")))?;
                tokens.push(Token::Number(number));
            }
            _ if ch.is_ascii_alphabetic() || ch == '_' => {
                let start = index;
                index += 1;
                while index < chars.len()
                    && (chars[index].is_ascii_alphanumeric()
                        || chars[index] == '_'
                        || chars[index] == '.')
                {
                    index += 1;
                }
                let ident: String = chars[start..index].iter().collect();
                if let Ok(cell) = CellAddr::from_a1(&ident) {
                    tokens.push(Token::Cell(cell));
                } else {
                    tokens.push(Token::Ident(ident.to_ascii_uppercase()));
                }
            }
            _ => {
                return Err(LoError::Parse(format!(
                    "unexpected character in formula: {ch}"
                )));
            }
        }
    }

    Ok(tokens)
}

struct Parser {
    tokens: Vec<Token>,
    index: usize,
}

impl Parser {
    fn parse_expression(&mut self) -> Result<Expr> {
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Result<Expr> {
        let mut expr = self.parse_concat()?;
        while let Some(op) = self.match_comparison() {
            let rhs = self.parse_concat()?;
            expr = Expr::Binary {
                op,
                lhs: Box::new(expr),
                rhs: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_concat(&mut self) -> Result<Expr> {
        let mut expr = self.parse_add_sub()?;
        while self.consume_if(|token| matches!(token, Token::Amp)) {
            let rhs = self.parse_add_sub()?;
            expr = Expr::Binary {
                op: BinaryOp::Concat,
                lhs: Box::new(expr),
                rhs: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_add_sub(&mut self) -> Result<Expr> {
        let mut expr = self.parse_mul_div()?;
        loop {
            let op = if self.consume_if(|token| matches!(token, Token::Plus)) {
                Some(BinaryOp::Add)
            } else if self.consume_if(|token| matches!(token, Token::Minus)) {
                Some(BinaryOp::Sub)
            } else {
                None
            };

            if let Some(op) = op {
                let rhs = self.parse_mul_div()?;
                expr = Expr::Binary {
                    op,
                    lhs: Box::new(expr),
                    rhs: Box::new(rhs),
                };
            } else {
                break;
            }
        }
        Ok(expr)
    }

    fn parse_mul_div(&mut self) -> Result<Expr> {
        let mut expr = self.parse_power()?;
        loop {
            let op = if self.consume_if(|token| matches!(token, Token::Star)) {
                Some(BinaryOp::Mul)
            } else if self.consume_if(|token| matches!(token, Token::Slash)) {
                Some(BinaryOp::Div)
            } else {
                None
            };

            if let Some(op) = op {
                let rhs = self.parse_power()?;
                expr = Expr::Binary {
                    op,
                    lhs: Box::new(expr),
                    rhs: Box::new(rhs),
                };
            } else {
                break;
            }
        }
        Ok(expr)
    }

    fn parse_power(&mut self) -> Result<Expr> {
        let mut expr = self.parse_unary()?;
        while self.consume_if(|token| matches!(token, Token::Caret)) {
            let rhs = self.parse_unary()?;
            expr = Expr::Binary {
                op: BinaryOp::Pow,
                lhs: Box::new(expr),
                rhs: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_unary(&mut self) -> Result<Expr> {
        if self.consume_if(|token| matches!(token, Token::Plus)) {
            return Ok(Expr::Unary {
                op: UnaryOp::Plus,
                expr: Box::new(self.parse_unary()?),
            });
        }
        if self.consume_if(|token| matches!(token, Token::Minus)) {
            return Ok(Expr::Unary {
                op: UnaryOp::Minus,
                expr: Box::new(self.parse_unary()?),
            });
        }
        self.parse_primary()
    }

    fn parse_primary(&mut self) -> Result<Expr> {
        let token = self
            .tokens
            .get(self.index)
            .cloned()
            .ok_or_else(|| LoError::Parse("unexpected end of formula".to_string()))?;
        self.index += 1;
        let mut expr = match token {
            Token::Number(value) => Expr::Number(value),
            Token::Text(text) => Expr::Text(text),
            Token::Cell(cell) => Expr::Ref(cell),
            Token::Ident(name) => {
                if name == "TRUE" {
                    Expr::Bool(true)
                } else if name == "FALSE" {
                    Expr::Bool(false)
                } else if self.consume_if(|token| matches!(token, Token::LParen)) {
                    let mut args = Vec::new();
                    if !self.consume_if(|token| matches!(token, Token::RParen)) {
                        loop {
                            args.push(self.parse_expression()?);
                            if self.consume_if(|token| matches!(token, Token::Comma)) {
                                continue;
                            }
                            self.expect(|token| matches!(token, Token::RParen), ")")?;
                            break;
                        }
                    }
                    Expr::Func { name, args }
                } else {
                    return Err(LoError::Parse(format!("unexpected identifier: {name}")));
                }
            }
            Token::LParen => {
                let expr = self.parse_expression()?;
                self.expect(|token| matches!(token, Token::RParen), ")")?;
                expr
            }
            other => {
                return Err(LoError::Parse(format!(
                    "unexpected token in primary: {other:?}"
                )));
            }
        };

        if self.consume_if(|token| matches!(token, Token::Colon)) {
            let end = match self
                .tokens
                .get(self.index)
                .cloned()
                .ok_or_else(|| LoError::Parse("expected range end".to_string()))?
            {
                Token::Cell(cell) => {
                    self.index += 1;
                    cell
                }
                _ => {
                    return Err(LoError::Parse(
                        "expected cell reference after ':'".to_string(),
                    ))
                }
            };
            match expr {
                Expr::Ref(start) => expr = Expr::Range(start, end),
                _ => {
                    return Err(LoError::Parse(
                        "range start must be a cell reference".to_string(),
                    ))
                }
            }
        }

        Ok(expr)
    }

    fn match_comparison(&mut self) -> Option<BinaryOp> {
        if self.consume_if(|token| matches!(token, Token::Eq)) {
            Some(BinaryOp::Eq)
        } else if self.consume_if(|token| matches!(token, Token::Ne)) {
            Some(BinaryOp::Ne)
        } else if self.consume_if(|token| matches!(token, Token::Lt)) {
            Some(BinaryOp::Lt)
        } else if self.consume_if(|token| matches!(token, Token::Lte)) {
            Some(BinaryOp::Lte)
        } else if self.consume_if(|token| matches!(token, Token::Gt)) {
            Some(BinaryOp::Gt)
        } else if self.consume_if(|token| matches!(token, Token::Gte)) {
            Some(BinaryOp::Gte)
        } else {
            None
        }
    }

    fn consume_if(&mut self, predicate: impl Fn(&Token) -> bool) -> bool {
        if self.tokens.get(self.index).map(predicate).unwrap_or(false) {
            self.index += 1;
            true
        } else {
            false
        }
    }

    fn expect(&mut self, predicate: impl Fn(&Token) -> bool, expected: &str) -> Result<()> {
        if self.consume_if(predicate) {
            Ok(())
        } else {
            Err(LoError::Parse(format!("expected {expected}")))
        }
    }
}

fn eval_expr(expr: &Expr, sheet: &Sheet, stack: &mut BTreeSet<CellAddr>) -> Result<Value> {
    match expr {
        Expr::Number(value) => Ok(Value::Number(*value)),
        Expr::Text(text) => Ok(Value::Text(text.clone())),
        Expr::Bool(value) => Ok(Value::Bool(*value)),
        Expr::Ref(addr) => evaluate_cell(sheet, *addr, stack),
        Expr::Range(_, _) => Err(LoError::Eval(
            "range cannot be used as a scalar value here".to_string(),
        )),
        Expr::Func { name, args } => evaluate_function(name, args, sheet, stack),
        Expr::Unary { op, expr } => {
            let value = eval_expr(expr, sheet, stack)?.as_number()?;
            match op {
                UnaryOp::Plus => Ok(Value::Number(value)),
                UnaryOp::Minus => Ok(Value::Number(-value)),
            }
        }
        Expr::Binary { op, lhs, rhs } => {
            match op {
                BinaryOp::Concat => {
                    let left = eval_expr(lhs, sheet, stack)?.as_text();
                    let right = eval_expr(rhs, sheet, stack)?.as_text();
                    return Ok(Value::Text(format!("{left}{right}")));
                }
                BinaryOp::Eq
                | BinaryOp::Ne
                | BinaryOp::Lt
                | BinaryOp::Lte
                | BinaryOp::Gt
                | BinaryOp::Gte => {
                    let left = eval_expr(lhs, sheet, stack)?;
                    let right = eval_expr(rhs, sheet, stack)?;
                    let result = compare_values(op, &left, &right)?;
                    return Ok(Value::Bool(result));
                }
                _ => {}
            }
            let left = eval_expr(lhs, sheet, stack)?.as_number()?;
            let right = eval_expr(rhs, sheet, stack)?.as_number()?;
            let value = match op {
                BinaryOp::Add => left + right,
                BinaryOp::Sub => left - right,
                BinaryOp::Mul => left * right,
                BinaryOp::Div => {
                    if right == 0.0 {
                        return Err(LoError::Eval("division by zero".to_string()));
                    }
                    left / right
                }
                BinaryOp::Pow => left.powf(right),
                BinaryOp::Concat
                | BinaryOp::Eq
                | BinaryOp::Ne
                | BinaryOp::Lt
                | BinaryOp::Lte
                | BinaryOp::Gt
                | BinaryOp::Gte => unreachable!(),
            };
            Ok(Value::Number(value))
        }
    }
}

fn compare_values(op: &BinaryOp, lhs: &Value, rhs: &Value) -> Result<bool> {
    let comparison = match (lhs, rhs) {
        (Value::Text(left), Value::Text(right)) => match left.cmp(right) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        },
        _ => {
            let left = lhs.as_number()?;
            let right = rhs.as_number()?;
            if left < right {
                -1
            } else if left > right {
                1
            } else {
                0
            }
        }
    };
    Ok(match op {
        BinaryOp::Eq => comparison == 0,
        BinaryOp::Ne => comparison != 0,
        BinaryOp::Lt => comparison < 0,
        BinaryOp::Lte => comparison <= 0,
        BinaryOp::Gt => comparison > 0,
        BinaryOp::Gte => comparison >= 0,
        _ => return Err(LoError::Eval("invalid comparison operator".to_string())),
    })
}

fn evaluate_cell(sheet: &Sheet, addr: CellAddr, stack: &mut BTreeSet<CellAddr>) -> Result<Value> {
    if !stack.insert(addr) {
        return Err(LoError::Eval(format!(
            "circular reference at {}",
            addr.to_a1()
        )));
    }
    let result = match sheet.get(addr) {
        None => Value::Blank,
        Some(Cell { value, .. }) => match value {
            CellValue::Empty => Value::Blank,
            CellValue::Number(value) => Value::Number(*value),
            CellValue::Text(value) => Value::Text(value.clone()),
            CellValue::Bool(value) => Value::Bool(*value),
            CellValue::Formula(formula) => {
                let expr = parse_formula(formula)?;
                eval_expr(&expr, sheet, stack)?
            }
            CellValue::Error(value) => Value::Error(value.clone()),
        },
    };
    stack.remove(&addr);
    Ok(result)
}

fn values_for_expr(
    expr: &Expr,
    sheet: &Sheet,
    stack: &mut BTreeSet<CellAddr>,
) -> Result<Vec<Value>> {
    match expr {
        Expr::Range(start, end) => {
            let range = lo_core::Range::new(*start, *end);
            let mut values = Vec::new();
            for addr in range.iter() {
                values.push(evaluate_cell(sheet, addr, stack)?);
            }
            Ok(values)
        }
        _ => Ok(vec![eval_expr(expr, sheet, stack)?]),
    }
}

fn numeric_values(
    args: &[Expr],
    sheet: &Sheet,
    stack: &mut BTreeSet<CellAddr>,
) -> Result<Vec<f64>> {
    // Aggregates over ranges follow the Calc/Excel convention: text and
    // blank cells are silently skipped instead of raising an error. This
    // matches real spreadsheets and lets `SUM(A1:A4)` work on a column
    // whose first row is a header label.
    let mut values = Vec::new();
    for arg in args {
        let is_range = matches!(arg, Expr::Range(_, _));
        for value in values_for_expr(arg, sheet, stack)? {
            match value {
                Value::Blank => continue,
                Value::Text(_) if is_range => continue,
                other => values.push(other.as_number()?),
            }
        }
    }
    Ok(values)
}

fn evaluate_function(
    name: &str,
    args: &[Expr],
    sheet: &Sheet,
    stack: &mut BTreeSet<CellAddr>,
) -> Result<Value> {
    match name {
        "SUM" => Ok(Value::Number(
            numeric_values(args, sheet, stack)?.into_iter().sum(),
        )),
        "AVERAGE" | "AVG" => {
            let values = numeric_values(args, sheet, stack)?;
            if values.is_empty() {
                return Ok(Value::Blank);
            }
            Ok(Value::Number(
                values.iter().sum::<f64>() / values.len() as f64,
            ))
        }
        "MIN" => {
            let values = numeric_values(args, sheet, stack)?;
            values
                .into_iter()
                .reduce(f64::min)
                .map(Value::Number)
                .ok_or_else(|| LoError::Eval("MIN requires at least one value".to_string()))
        }
        "MAX" => {
            let values = numeric_values(args, sheet, stack)?;
            values
                .into_iter()
                .reduce(f64::max)
                .map(Value::Number)
                .ok_or_else(|| LoError::Eval("MAX requires at least one value".to_string()))
        }
        "COUNT" => Ok(Value::Number(
            numeric_values(args, sheet, stack)?.len() as f64
        )),
        "IF" => {
            if !(2..=3).contains(&args.len()) {
                return Err(LoError::Eval("IF expects 2 or 3 arguments".to_string()));
            }
            let condition = eval_expr(&args[0], sheet, stack)?.as_bool()?;
            if condition {
                eval_expr(&args[1], sheet, stack)
            } else if args.len() == 3 {
                eval_expr(&args[2], sheet, stack)
            } else {
                Ok(Value::Blank)
            }
        }
        "AND" => {
            for arg in args {
                if !eval_expr(arg, sheet, stack)?.as_bool()? {
                    return Ok(Value::Bool(false));
                }
            }
            Ok(Value::Bool(true))
        }
        "OR" => {
            for arg in args {
                if eval_expr(arg, sheet, stack)?.as_bool()? {
                    return Ok(Value::Bool(true));
                }
            }
            Ok(Value::Bool(false))
        }
        "NOT" => {
            if args.len() != 1 {
                return Err(LoError::Eval("NOT expects 1 argument".to_string()));
            }
            Ok(Value::Bool(!eval_expr(&args[0], sheet, stack)?.as_bool()?))
        }
        "CONCAT" => {
            let mut out = String::new();
            for arg in args {
                for value in values_for_expr(arg, sheet, stack)? {
                    out.push_str(&value.as_text());
                }
            }
            Ok(Value::Text(out))
        }
        "LEN" => {
            if args.len() != 1 {
                return Err(LoError::Eval("LEN expects 1 argument".to_string()));
            }
            Ok(Value::Number(
                eval_expr(&args[0], sheet, stack)?.as_text().chars().count() as f64,
            ))
        }
        "ABS" => {
            if args.len() != 1 {
                return Err(LoError::Eval("ABS expects 1 argument".to_string()));
            }
            Ok(Value::Number(
                eval_expr(&args[0], sheet, stack)?.as_number()?.abs(),
            ))
        }
        "ROUND" => {
            if !(1..=2).contains(&args.len()) {
                return Err(LoError::Eval("ROUND expects 1 or 2 arguments".to_string()));
            }
            let value = eval_expr(&args[0], sheet, stack)?.as_number()?;
            let digits = if args.len() == 2 {
                eval_expr(&args[1], sheet, stack)?.as_number()? as i32
            } else {
                0
            };
            let scale = 10f64.powi(digits);
            Ok(Value::Number((value * scale).round() / scale))
        }
        other => Err(LoError::Unsupported(format!("unknown function: {other}"))),
    }
}

pub fn parse_csv(input: &str) -> Result<Vec<Vec<String>>> {
    let mut rows = Vec::new();
    let mut row = Vec::new();
    let mut field = String::new();
    let mut chars = input.chars().peekable();
    let mut in_quotes = false;

    while let Some(ch) = chars.next() {
        match ch {
            '"' => {
                if in_quotes && chars.peek() == Some(&'"') {
                    field.push('"');
                    chars.next();
                } else {
                    in_quotes = !in_quotes;
                }
            }
            ',' if !in_quotes => {
                row.push(std::mem::take(&mut field));
            }
            '\n' if !in_quotes => {
                row.push(std::mem::take(&mut field));
                rows.push(std::mem::take(&mut row));
            }
            '\r' if !in_quotes => {
                if chars.peek() == Some(&'\n') {
                    continue;
                }
            }
            _ => field.push(ch),
        }
    }

    if in_quotes {
        return Err(LoError::Parse("unterminated CSV quoted field".to_string()));
    }

    if !field.is_empty() || !row.is_empty() {
        row.push(field);
        rows.push(row);
    }

    Ok(rows)
}

fn escape_csv(value: &str) -> String {
    if value.contains(',') || value.contains('"') || value.contains('\n') {
        format!("\"{}\"", value.replace('"', "\"\""))
    } else {
        value.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::{
        evaluate_formula, parse_csv, parse_formula, sheet_to_csv, workbook_from_csv, Value,
    };
    use lo_core::{CellAddr, CellValue, Sheet};

    #[test]
    fn parser_handles_functions_and_ranges() {
        let expr = parse_formula("=SUM(A1:A3)").expect("parse formula");
        let dump = format!("{expr:?}");
        assert!(dump.contains("SUM"));
        assert!(dump.contains("Range"));
    }

    #[test]
    fn evaluator_handles_sum_and_if() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set(CellAddr::new(0, 0), CellValue::Number(1.0));
        sheet.set(CellAddr::new(1, 0), CellValue::Number(2.0));
        sheet.set(CellAddr::new(2, 0), CellValue::Number(3.0));
        let sum = evaluate_formula("=SUM(A1:A3)", &sheet).expect("evaluate sum");
        assert_eq!(sum, Value::Number(6.0));
        let flag = evaluate_formula("=IF(SUM(A1:A3)>5, TRUE, FALSE)", &sheet).expect("evaluate if");
        assert_eq!(flag, Value::Bool(true));
    }

    #[test]
    fn xlsx_export_is_a_zip_archive() {
        let csv = "name,age\nAlice,30\nBob,22";
        let workbook = workbook_from_csv("People", "People", csv).expect("workbook from csv");
        let bytes = super::to_xlsx(&workbook).expect("xlsx");
        assert!(bytes.starts_with(b"PK"));
    }

    #[test]
    fn html_export_includes_table() {
        let workbook = workbook_from_csv("X", "X", "a,b\n1,2").expect("workbook");
        let html = super::to_html(&workbook);
        assert!(html.contains("<table"));
        assert!(html.contains(">1<"));
    }

    #[test]
    fn pdf_export_starts_with_header() {
        let workbook = workbook_from_csv("X", "X", "a,b\n1,2").expect("workbook");
        let pdf = super::to_pdf(&workbook);
        assert!(pdf.starts_with(b"%PDF-1.4"));
    }

    #[test]
    fn save_as_dispatches_by_format() {
        let workbook = workbook_from_csv("X", "X", "a,b\n1,2").expect("workbook");
        for fmt in ["csv", "html", "svg", "pdf", "ods", "xlsx"] {
            let bytes = super::save_as(&workbook, fmt).unwrap_or_else(|e| panic!("{fmt}: {e}"));
            assert!(!bytes.is_empty(), "{fmt} produced empty output");
        }
        assert!(super::save_as(&workbook, "zzz").is_err());
    }

    #[test]
    fn csv_roundtrip_preserves_values() {
        let csv = "name,age\nAlice,30\nBob,22";
        let workbook = workbook_from_csv("People", "People", csv).expect("workbook from csv");
        let out = sheet_to_csv(&workbook.sheets[0]);
        assert!(out.contains("Alice"));
        assert!(out.contains("30"));
        let rows = parse_csv(&out).expect("parse csv");
        assert_eq!(rows.len(), 3);
    }
}
