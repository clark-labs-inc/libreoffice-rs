use std::cell::RefCell;
use std::collections::{BTreeMap, BTreeSet};

use lo_core::{CellAddr, CellValue, LoError, Result, Sheet, Workbook};

#[derive(Clone, Debug, PartialEq)]
pub enum EvalValue {
    Blank,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
}

#[derive(Clone, Debug, PartialEq)]
enum EvalNode {
    Scalar(EvalValue),
    Range(Vec<EvalValue>),
}

impl EvalNode {
    fn into_scalar(self) -> EvalValue {
        match self {
            Self::Scalar(value) => value,
            Self::Range(values) => values.into_iter().next().unwrap_or(EvalValue::Blank),
        }
    }

    fn as_values(&self) -> Vec<EvalValue> {
        match self {
            Self::Scalar(value) => vec![value.clone()],
            Self::Range(values) => values.clone(),
        }
    }
}

impl EvalValue {
    pub fn as_number(&self) -> Result<f64> {
        match self {
            Self::Blank => Ok(0.0),
            Self::Number(value) => Ok(*value),
            Self::Text(text) => text
                .trim()
                .parse::<f64>()
                .map_err(|_| LoError::Eval(format!("cannot convert text to number: {text}"))),
            Self::Bool(value) => Ok(if *value { 1.0 } else { 0.0 }),
            Self::Error(message) => Err(LoError::Eval(message.clone())),
        }
    }

    pub fn as_bool(&self) -> Result<bool> {
        match self {
            Self::Blank => Ok(false),
            Self::Number(value) => Ok(*value != 0.0),
            Self::Text(text) => Ok(!text.is_empty()),
            Self::Bool(value) => Ok(*value),
            Self::Error(message) => Err(LoError::Eval(message.clone())),
        }
    }

    pub fn as_text(&self) -> String {
        match self {
            Self::Blank => String::new(),
            Self::Number(value) => {
                if value.fract() == 0.0 && value.is_finite() {
                    format!("{}", *value as i64)
                } else {
                    value.to_string()
                }
            }
            Self::Text(text) => text.clone(),
            Self::Bool(value) => {
                if *value {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            Self::Error(message) => message.clone(),
        }
    }

    #[allow(dead_code)]
    pub fn is_error(&self) -> bool {
        matches!(self, Self::Error(_))
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct CellRef {
    col: u32,
    row: u32,
    col_abs: bool,
    row_abs: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ColRef {
    col: u32,
    abs: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct RowRef {
    row: u32,
    abs: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum RefKind {
    Cell(CellRef),
    Range(CellRef, CellRef),
    ColRange(ColRef, ColRef),
    RowRange(RowRef, RowRef),
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Reference {
    sheet: Option<String>,
    kind: RefKind,
}

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Number(f64),
    String(String),
    Bool(bool),
    Ref(Reference),
    Ident(String),
    Comma,
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

#[derive(Clone, Debug, PartialEq)]
enum Expr {
    Number(f64),
    String(String),
    Bool(bool),
    Ref(Reference),
    Func { name: String, args: Vec<Expr> },
    Unary { op: UnaryOp, expr: Box<Expr> },
    Binary { op: BinaryOp, lhs: Box<Expr>, rhs: Box<Expr> },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum UnaryOp {
    Plus,
    Minus,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum BinaryOp {
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
        loop {
            let op = match self.peek() {
                Some(Token::Eq) => BinaryOp::Eq,
                Some(Token::Ne) => BinaryOp::Ne,
                Some(Token::Lt) => BinaryOp::Lt,
                Some(Token::Lte) => BinaryOp::Lte,
                Some(Token::Gt) => BinaryOp::Gt,
                Some(Token::Gte) => BinaryOp::Gte,
                _ => break,
            };
            self.index += 1;
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
        while matches!(self.peek(), Some(Token::Amp)) {
            self.index += 1;
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
            let op = match self.peek() {
                Some(Token::Plus) => BinaryOp::Add,
                Some(Token::Minus) => BinaryOp::Sub,
                _ => break,
            };
            self.index += 1;
            let rhs = self.parse_mul_div()?;
            expr = Expr::Binary {
                op,
                lhs: Box::new(expr),
                rhs: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_mul_div(&mut self) -> Result<Expr> {
        let mut expr = self.parse_power()?;
        loop {
            let op = match self.peek() {
                Some(Token::Star) => BinaryOp::Mul,
                Some(Token::Slash) => BinaryOp::Div,
                _ => break,
            };
            self.index += 1;
            let rhs = self.parse_power()?;
            expr = Expr::Binary {
                op,
                lhs: Box::new(expr),
                rhs: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_power(&mut self) -> Result<Expr> {
        let mut expr = self.parse_unary()?;
        while matches!(self.peek(), Some(Token::Caret)) {
            self.index += 1;
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
        match self.peek() {
            Some(Token::Plus) => {
                self.index += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::Plus,
                    expr: Box::new(self.parse_unary()?),
                })
            }
            Some(Token::Minus) => {
                self.index += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::Minus,
                    expr: Box::new(self.parse_unary()?),
                })
            }
            _ => self.parse_primary(),
        }
    }

    fn parse_primary(&mut self) -> Result<Expr> {
        match self.next() {
            Some(Token::Number(value)) => Ok(Expr::Number(value)),
            Some(Token::String(value)) => Ok(Expr::String(value)),
            Some(Token::Bool(value)) => Ok(Expr::Bool(value)),
            Some(Token::Ref(reference)) => Ok(Expr::Ref(reference)),
            Some(Token::Ident(name)) => {
                if !matches!(self.peek(), Some(Token::LParen)) {
                    return Ok(Expr::String(name));
                }
                self.index += 1; // (
                let mut args = Vec::new();
                if !matches!(self.peek(), Some(Token::RParen)) {
                    loop {
                        args.push(self.parse_expression()?);
                        match self.peek() {
                            Some(Token::Comma) => self.index += 1,
                            _ => break,
                        }
                    }
                }
                match self.next() {
                    Some(Token::RParen) => Ok(Expr::Func { name, args }),
                    _ => Err(LoError::Parse("expected ')' after function arguments".to_string())),
                }
            }
            Some(Token::LParen) => {
                let expr = self.parse_expression()?;
                match self.next() {
                    Some(Token::RParen) => Ok(expr),
                    _ => Err(LoError::Parse("expected ')'".to_string())),
                }
            }
            other => Err(LoError::Parse(format!("unexpected token in formula: {other:?}"))),
        }
    }

    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.index)
    }

    fn next(&mut self) -> Option<Token> {
        let token = self.tokens.get(self.index).cloned();
        if token.is_some() {
            self.index += 1;
        }
        token
    }
}

pub struct WorkbookEvaluator<'a> {
    workbook: &'a Workbook,
    sheets: BTreeMap<String, usize>,
    cache: RefCell<BTreeMap<(usize, CellAddr), EvalValue>>,
    stack: RefCell<BTreeSet<(usize, CellAddr)>>,
}

impl<'a> WorkbookEvaluator<'a> {
    pub fn new(workbook: &'a Workbook) -> Self {
        let mut sheets = BTreeMap::new();
        for (index, sheet) in workbook.sheets.iter().enumerate() {
            sheets.insert(sheet.name.to_ascii_lowercase(), index);
        }
        Self {
            workbook,
            sheets,
            cache: RefCell::new(BTreeMap::new()),
            stack: RefCell::new(BTreeSet::new()),
        }
    }

    pub fn evaluate_formula(&self, sheet_index: usize, formula: &str) -> Result<EvalValue> {
        let tokens = tokenize(formula)?;
        let mut parser = Parser { tokens, index: 0 };
        let expr = parser.parse_expression()?;
        if parser.index != parser.tokens.len() {
            return Err(LoError::Parse("unexpected trailing tokens in formula".to_string()));
        }
        Ok(self.eval_expr(sheet_index, &expr)?.into_scalar())
    }

    fn eval_expr(&self, sheet_index: usize, expr: &Expr) -> Result<EvalNode> {
        match expr {
            Expr::Number(value) => Ok(EvalNode::Scalar(EvalValue::Number(*value))),
            Expr::String(value) => Ok(EvalNode::Scalar(EvalValue::Text(value.clone()))),
            Expr::Bool(value) => Ok(EvalNode::Scalar(EvalValue::Bool(*value))),
            Expr::Ref(reference) => self.eval_reference(sheet_index, reference),
            Expr::Func { name, args } => self.eval_function(sheet_index, name, args),
            Expr::Unary { op, expr } => {
                let value = self.eval_expr(sheet_index, expr)?.into_scalar();
                match op {
                    UnaryOp::Plus => Ok(EvalNode::Scalar(EvalValue::Number(value.as_number()?))),
                    UnaryOp::Minus => Ok(EvalNode::Scalar(EvalValue::Number(-value.as_number()?))),
                }
            }
            Expr::Binary { op, lhs, rhs } => self.eval_binary(sheet_index, *op, lhs, rhs),
        }
    }

    fn eval_binary(
        &self,
        sheet_index: usize,
        op: BinaryOp,
        lhs: &Expr,
        rhs: &Expr,
    ) -> Result<EvalNode> {
        let lhs = self.eval_expr(sheet_index, lhs)?.into_scalar();
        let rhs = self.eval_expr(sheet_index, rhs)?.into_scalar();
        let value = match op {
            BinaryOp::Add => EvalValue::Number(lhs.as_number()? + rhs.as_number()?),
            BinaryOp::Sub => EvalValue::Number(lhs.as_number()? - rhs.as_number()?),
            BinaryOp::Mul => EvalValue::Number(lhs.as_number()? * rhs.as_number()?),
            BinaryOp::Div => {
                let divisor = rhs.as_number()?;
                if divisor == 0.0 {
                    EvalValue::Error("#DIV/0!".to_string())
                } else {
                    EvalValue::Number(lhs.as_number()? / divisor)
                }
            }
            BinaryOp::Pow => EvalValue::Number(lhs.as_number()?.powf(rhs.as_number()?)),
            BinaryOp::Concat => EvalValue::Text(format!("{}{}", lhs.as_text(), rhs.as_text())),
            BinaryOp::Eq => EvalValue::Bool(values_equal(&lhs, &rhs)),
            BinaryOp::Ne => EvalValue::Bool(!values_equal(&lhs, &rhs)),
            BinaryOp::Lt => EvalValue::Bool(lhs.as_number()? < rhs.as_number()?),
            BinaryOp::Lte => EvalValue::Bool(lhs.as_number()? <= rhs.as_number()?),
            BinaryOp::Gt => EvalValue::Bool(lhs.as_number()? > rhs.as_number()?),
            BinaryOp::Gte => EvalValue::Bool(lhs.as_number()? >= rhs.as_number()?),
        };
        Ok(EvalNode::Scalar(value))
    }

    fn eval_function(&self, sheet_index: usize, name: &str, args: &[Expr]) -> Result<EvalNode> {
        let upper = name.to_ascii_uppercase();
        let values = args
            .iter()
            .map(|expr| self.eval_expr(sheet_index, expr))
            .collect::<Result<Vec<_>>>()?;
        let flat = flatten_values(&values);
        let scalar = |index: usize| -> EvalValue {
            values
                .get(index)
                .cloned()
                .unwrap_or(EvalNode::Scalar(EvalValue::Blank))
                .into_scalar()
        };
        let out = match upper.as_str() {
            "SUM" => EvalValue::Number(sum_values(&flat)?),
            "AVERAGE" => {
                let count = flat.iter().filter(|v| matches!(v, EvalValue::Number(_))).count();
                if count == 0 {
                    EvalValue::Error("#DIV/0!".to_string())
                } else {
                    EvalValue::Number(sum_values(&flat)? / count as f64)
                }
            }
            "MIN" => EvalValue::Number(min_or_max(&flat, true)?),
            "MAX" => EvalValue::Number(min_or_max(&flat, false)?),
            "COUNT" => EvalValue::Number(
                flat.iter()
                    .filter(|value| matches!(value, EvalValue::Number(_)))
                    .count() as f64,
            ),
            "COUNTA" => EvalValue::Number(
                flat.iter()
                    .filter(|value| !matches!(value, EvalValue::Blank))
                    .count() as f64,
            ),
            "IF" => {
                let cond = scalar(0).as_bool()?;
                if cond {
                    scalar(1)
                } else {
                    scalar(2)
                }
            }
            "AND" => EvalValue::Bool(flat.iter().all(|value| value.as_bool().unwrap_or(false))),
            "OR" => EvalValue::Bool(flat.iter().any(|value| value.as_bool().unwrap_or(false))),
            "NOT" => EvalValue::Bool(!scalar(0).as_bool()?),
            "ABS" => EvalValue::Number(scalar(0).as_number()?.abs()),
            "INT" => EvalValue::Number(scalar(0).as_number()?.floor()),
            "ROUND" => {
                let number = scalar(0).as_number()?;
                let digits = scalar(1).as_number().unwrap_or(0.0) as i32;
                let factor = 10_f64.powi(digits);
                EvalValue::Number((number * factor).round() / factor)
            }
            "LEN" => EvalValue::Number(scalar(0).as_text().chars().count() as f64),
            "CONCAT" | "CONCATENATE" => EvalValue::Text(
                flat.iter()
                    .map(EvalValue::as_text)
                    .collect::<Vec<_>>()
                    .join(""),
            ),
            "LEFT" => {
                let text = scalar(0).as_text();
                let count = scalar(1).as_number().unwrap_or(1.0).max(0.0) as usize;
                EvalValue::Text(text.chars().take(count).collect())
            }
            "RIGHT" => {
                let text = scalar(0).as_text();
                let count = scalar(1).as_number().unwrap_or(1.0).max(0.0) as usize;
                let len = text.chars().count();
                EvalValue::Text(text.chars().skip(len.saturating_sub(count)).collect())
            }
            "UPPER" => EvalValue::Text(scalar(0).as_text().to_ascii_uppercase()),
            "LOWER" => EvalValue::Text(scalar(0).as_text().to_ascii_lowercase()),
            "TRIM" => EvalValue::Text(scalar(0).as_text().trim().to_string()),
            "POWER" => EvalValue::Number(scalar(0).as_number()?.powf(scalar(1).as_number()?)),
            other => {
                return Err(LoError::Unsupported(format!(
                    "xlsx formula function not supported yet: {other}"
                )))
            }
        };
        Ok(EvalNode::Scalar(out))
    }

    fn eval_reference(&self, sheet_index: usize, reference: &Reference) -> Result<EvalNode> {
        let target_sheet_index = reference
            .sheet
            .as_ref()
            .and_then(|name| self.sheets.get(&name.to_ascii_lowercase()).copied())
            .unwrap_or(sheet_index);
        let Some(sheet) = self.workbook.sheets.get(target_sheet_index) else {
            return Ok(EvalNode::Scalar(EvalValue::Error("#REF!".to_string())));
        };
        match &reference.kind {
            RefKind::Cell(cell) => {
                let addr = CellAddr::new(cell.row, cell.col);
                Ok(EvalNode::Scalar(self.cell_value(target_sheet_index, sheet, addr)?))
            }
            RefKind::Range(start, end) => Ok(EvalNode::Range(self.collect_rect(sheet_index, target_sheet_index, sheet, *start, *end)?)),
            RefKind::ColRange(start, end) => Ok(EvalNode::Range(self.collect_columns(target_sheet_index, sheet, *start, *end)?)),
            RefKind::RowRange(start, end) => Ok(EvalNode::Range(self.collect_rows(target_sheet_index, sheet, *start, *end)?)),
        }
    }

    fn collect_rect(
        &self,
        _current_sheet_index: usize,
        target_sheet_index: usize,
        sheet: &Sheet,
        start: CellRef,
        end: CellRef,
    ) -> Result<Vec<EvalValue>> {
        let row_start = start.row.min(end.row);
        let row_end = start.row.max(end.row);
        let col_start = start.col.min(end.col);
        let col_end = start.col.max(end.col);
        let start = CellAddr::new(row_start, col_start);
        let end = CellAddr::new(row_end, col_end);
        let mut out = Vec::new();
        for row in start.row..=end.row {
            for col in start.col..=end.col {
                out.push(self.cell_value(target_sheet_index, sheet, CellAddr::new(row, col))?);
            }
        }
        Ok(out)
    }

    fn collect_columns(
        &self,
        target_sheet_index: usize,
        sheet: &Sheet,
        start: ColRef,
        end: ColRef,
    ) -> Result<Vec<EvalValue>> {
        let (_, max_col) = sheet.max_extent();
        let max_row = sheet.max_extent().0.max(0);
        let col_start = start.col.min(end.col);
        let col_end = start.col.max(end.col).min(max_col.max(col_start));
        let mut out = Vec::new();
        for row in 0..=max_row {
            for col in col_start..=col_end {
                out.push(self.cell_value(target_sheet_index, sheet, CellAddr::new(row, col))?);
            }
        }
        Ok(out)
    }

    fn collect_rows(
        &self,
        target_sheet_index: usize,
        sheet: &Sheet,
        start: RowRef,
        end: RowRef,
    ) -> Result<Vec<EvalValue>> {
        let (max_row, max_col) = sheet.max_extent();
        let row_start = start.row.min(end.row);
        let row_end = start.row.max(end.row).min(max_row.max(row_start));
        let mut out = Vec::new();
        for row in row_start..=row_end {
            for col in 0..=max_col {
                out.push(self.cell_value(target_sheet_index, sheet, CellAddr::new(row, col))?);
            }
        }
        Ok(out)
    }

    fn cell_value(&self, sheet_index: usize, sheet: &Sheet, addr: CellAddr) -> Result<EvalValue> {
        if let Some(cached) = self.cache.borrow().get(&(sheet_index, addr)) {
            return Ok(cached.clone());
        }
        if self.stack.borrow().contains(&(sheet_index, addr)) {
            return Ok(EvalValue::Error("#CYCLE!".to_string()));
        }
        let value = match sheet.get(addr).map(|cell| &cell.value) {
            Some(CellValue::Empty) | None => EvalValue::Blank,
            Some(CellValue::Number(value)) => EvalValue::Number(*value),
            Some(CellValue::Text(value)) => EvalValue::Text(value.clone()),
            Some(CellValue::Bool(value)) => EvalValue::Bool(*value),
            Some(CellValue::Error(value)) => EvalValue::Error(value.clone()),
            Some(CellValue::Formula(formula)) => {
                self.stack.borrow_mut().insert((sheet_index, addr));
                let result = self.evaluate_formula(sheet_index, formula);
                self.stack.borrow_mut().remove(&(sheet_index, addr));
                match result {
                    Ok(value) => value,
                    Err(_) => EvalValue::Error("#VALUE!".to_string()),
                }
            }
        };
        self.cache.borrow_mut().insert((sheet_index, addr), value.clone());
        Ok(value)
    }
}

fn values_equal(lhs: &EvalValue, rhs: &EvalValue) -> bool {
    match (lhs, rhs) {
        (EvalValue::Blank, EvalValue::Blank) => true,
        (EvalValue::Number(a), EvalValue::Number(b)) => (*a - *b).abs() < f64::EPSILON,
        (EvalValue::Text(a), EvalValue::Text(b)) => a == b,
        (EvalValue::Bool(a), EvalValue::Bool(b)) => a == b,
        (EvalValue::Error(a), EvalValue::Error(b)) => a == b,
        _ => lhs.as_text() == rhs.as_text(),
    }
}

fn flatten_values(nodes: &[EvalNode]) -> Vec<EvalValue> {
    let mut out = Vec::new();
    for node in nodes {
        out.extend(node.as_values());
    }
    out
}

fn sum_values(values: &[EvalValue]) -> Result<f64> {
    let mut total = 0.0;
    for value in values {
        match value {
            EvalValue::Blank => {}
            EvalValue::Number(number) => total += *number,
            EvalValue::Bool(flag) => total += if *flag { 1.0 } else { 0.0 },
            EvalValue::Text(text) => {
                if let Ok(number) = text.parse::<f64>() {
                    total += number;
                }
            }
            EvalValue::Error(message) => return Err(LoError::Eval(message.clone())),
        }
    }
    Ok(total)
}

fn min_or_max(values: &[EvalValue], min: bool) -> Result<f64> {
    let mut iter = values.iter().filter_map(|value| value.as_number().ok());
    let Some(mut best) = iter.next() else {
        return Ok(0.0);
    };
    for value in iter {
        if (min && value < best) || (!min && value > best) {
            best = value;
        }
    }
    Ok(best)
}

fn tokenize(input: &str) -> Result<Vec<Token>> {
    let input = input.strip_prefix('=').unwrap_or(input).trim();
    let chars: Vec<char> = input.chars().collect();
    let mut tokens = Vec::new();
    let mut index = 0usize;
    while index < chars.len() {
        let ch = chars[index];
        if ch.is_whitespace() {
            index += 1;
            continue;
        }
        if let Some((len, reference)) = parse_reference(&chars, index) {
            tokens.push(Token::Ref(reference));
            index += len;
            continue;
        }
        if ch.is_ascii_digit() || (ch == '.' && chars.get(index + 1).map(|c| c.is_ascii_digit()).unwrap_or(false)) {
            let start = index;
            index += 1;
            while index < chars.len() && (chars[index].is_ascii_digit() || chars[index] == '.') {
                index += 1;
            }
            let number: String = chars[start..index].iter().collect();
            tokens.push(Token::Number(number.parse::<f64>().map_err(|_| {
                LoError::Parse(format!("invalid number in formula: {number}"))
            })?));
            continue;
        }
        if ch == '"' {
            index += 1;
            let mut text = String::new();
            while index < chars.len() {
                if chars[index] == '"' {
                    if chars.get(index + 1) == Some(&'"') {
                        text.push('"');
                        index += 2;
                        continue;
                    }
                    break;
                }
                text.push(chars[index]);
                index += 1;
            }
            if index >= chars.len() || chars[index] != '"' {
                return Err(LoError::Parse("unterminated string literal in formula".to_string()));
            }
            index += 1;
            tokens.push(Token::String(text));
            continue;
        }
        if is_ident_start(ch) {
            let start = index;
            index += 1;
            while index < chars.len() && is_ident_continue(chars[index]) {
                index += 1;
            }
            let ident: String = chars[start..index].iter().collect();
            match ident.to_ascii_uppercase().as_str() {
                "TRUE" => tokens.push(Token::Bool(true)),
                "FALSE" => tokens.push(Token::Bool(false)),
                _ => tokens.push(Token::Ident(ident)),
            }
            continue;
        }
        let token = match ch {
            ',' => Token::Comma,
            '(' => Token::LParen,
            ')' => Token::RParen,
            '+' => Token::Plus,
            '-' => Token::Minus,
            '*' => Token::Star,
            '/' => Token::Slash,
            '^' => Token::Caret,
            '&' => Token::Amp,
            '=' => Token::Eq,
            '<' if chars.get(index + 1) == Some(&'>') => {
                index += 2;
                tokens.push(Token::Ne);
                continue;
            }
            '<' if chars.get(index + 1) == Some(&'=') => {
                index += 2;
                tokens.push(Token::Lte);
                continue;
            }
            '>' if chars.get(index + 1) == Some(&'=') => {
                index += 2;
                tokens.push(Token::Gte);
                continue;
            }
            '<' => Token::Lt,
            '>' => Token::Gt,
            _ => {
                return Err(LoError::Parse(format!(
                    "unsupported character in formula: {ch}"
                )))
            }
        };
        index += 1;
        tokens.push(token);
    }
    Ok(tokens)
}

fn is_ident_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_'
}

fn is_ident_continue(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.')
}

fn parse_reference(chars: &[char], start: usize) -> Option<(usize, Reference)> {
    let mut index = start;
    let sheet = parse_sheet_qualifier(chars, &mut index);
    let (first, len) = parse_ref_part(chars, index, true)?;
    index += len;
    let mut second = None;
    if chars.get(index) == Some(&':') {
        index += 1;
        let (end_part, end_len) = parse_ref_part(chars, index, false)?;
        index += end_len;
        second = Some(end_part);
    }
    if is_reference_trailer(chars.get(index).copied()) {
        return None;
    }
    let kind = match (first, second) {
        (RefPart::Cell(a), Some(RefPart::Cell(b))) => RefKind::Range(a, b),
        (RefPart::Col(a), Some(RefPart::Col(b))) => RefKind::ColRange(a, b),
        (RefPart::Row(a), Some(RefPart::Row(b))) => RefKind::RowRange(a, b),
        (RefPart::Cell(cell), None) => RefKind::Cell(cell),
        (RefPart::Col(_), None) | (RefPart::Row(_), None) => return None,
        _ => return None,
    };
    Some((index - start, Reference { sheet, kind }))
}

fn is_reference_trailer(ch: Option<char>) -> bool {
    matches!(ch, Some(c) if c.is_ascii_alphanumeric() || matches!(c, '_' | '.'))
}

fn parse_sheet_qualifier(chars: &[char], index: &mut usize) -> Option<String> {
    let original = *index;
    if chars.get(*index) == Some(&'\'') {
        *index += 1;
        let mut name = String::new();
        while *index < chars.len() {
            if chars[*index] == '\'' {
                if chars.get(*index + 1) == Some(&'\'') {
                    name.push('\'');
                    *index += 2;
                    continue;
                }
                *index += 1;
                break;
            }
            name.push(chars[*index]);
            *index += 1;
        }
        if chars.get(*index) == Some(&'!') {
            *index += 1;
            return Some(name);
        }
        *index = original;
        return None;
    }
    let mut probe = *index;
    let mut name = String::new();
    while probe < chars.len() && (chars[probe].is_ascii_alphanumeric() || matches!(chars[probe], '_' | '.')) {
        name.push(chars[probe]);
        probe += 1;
    }
    if !name.is_empty() && chars.get(probe) == Some(&'!') {
        *index = probe + 1;
        return Some(name);
    }
    None
}

enum RefPart {
    Cell(CellRef),
    Col(ColRef),
    Row(RowRef),
}

fn parse_ref_part(chars: &[char], start: usize, require_range: bool) -> Option<(RefPart, usize)> {
    let mut index = start;
    let col_abs = if chars.get(index) == Some(&'$') {
        index += 1;
        true
    } else {
        false
    };
    let letters_start = index;
    while index < chars.len() && chars[index].is_ascii_alphabetic() {
        index += 1;
    }
    let letters_end = index;
    let row_abs = if chars.get(index) == Some(&'$') {
        index += 1;
        true
    } else {
        false
    };
    let digits_start = index;
    while index < chars.len() && chars[index].is_ascii_digit() {
        index += 1;
    }
    let digits_end = index;

    if letters_start != letters_end && digits_start != digits_end {
        let letters: String = chars[letters_start..letters_end].iter().collect();
        let digits: String = chars[digits_start..digits_end].iter().collect();
        let col = column_letters_to_index(&letters)?;
        let row = digits.parse::<u32>().ok()?.checked_sub(1)?;
        return Some((
            RefPart::Cell(CellRef {
                col,
                row,
                col_abs,
                row_abs,
            }),
            index - start,
        ));
    }

    if letters_start != letters_end {
        if require_range && chars.get(index) != Some(&':') {
            return None;
        }
        let letters: String = chars[letters_start..letters_end].iter().collect();
        let col = column_letters_to_index(&letters)?;
        return Some((RefPart::Col(ColRef { col, abs: col_abs }), index - start));
    }

    if digits_start != digits_end {
        if require_range && chars.get(index) != Some(&':') {
            return None;
        }
        let digits: String = chars[digits_start..digits_end].iter().collect();
        let row = digits.parse::<u32>().ok()?.checked_sub(1)?;
        return Some((RefPart::Row(RowRef { row, abs: row_abs }), index - start));
    }

    None
}

fn column_letters_to_index(letters: &str) -> Option<u32> {
    let mut col = 0u32;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        col = col.checked_mul(26)? + ((ch.to_ascii_uppercase() as u8 - b'A') as u32 + 1);
    }
    col.checked_sub(1)
}

fn column_index_to_letters(mut col: u32) -> String {
    col += 1;
    let mut letters = String::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        letters.insert(0, (b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    letters
}

pub fn translate_shared_formula(formula: &str, from: CellAddr, to: CellAddr) -> String {
    let row_delta = to.row as i64 - from.row as i64;
    let col_delta = to.col as i64 - from.col as i64;
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::new();
    let mut index = 0usize;
    let mut in_string = false;
    while index < chars.len() {
        if chars[index] == '"' {
            in_string = !in_string;
            out.push(chars[index]);
            index += 1;
            continue;
        }
        if !in_string {
            if let Some((len, reference)) = parse_reference(&chars, index) {
                out.push_str(&render_reference(&shift_reference(reference, row_delta, col_delta)));
                index += len;
                continue;
            }
        }
        out.push(chars[index]);
        index += 1;
    }
    out
}

fn shift_reference(reference: Reference, row_delta: i64, col_delta: i64) -> Reference {
    let kind = match reference.kind {
        RefKind::Cell(cell) => RefKind::Cell(shift_cell(cell, row_delta, col_delta)),
        RefKind::Range(a, b) => RefKind::Range(
            shift_cell(a, row_delta, col_delta),
            shift_cell(b, row_delta, col_delta),
        ),
        RefKind::ColRange(a, b) => RefKind::ColRange(
            shift_col(a, col_delta),
            shift_col(b, col_delta),
        ),
        RefKind::RowRange(a, b) => RefKind::RowRange(
            shift_row(a, row_delta),
            shift_row(b, row_delta),
        ),
    };
    Reference {
        sheet: reference.sheet,
        kind,
    }
}

fn shift_cell(cell: CellRef, row_delta: i64, col_delta: i64) -> CellRef {
    CellRef {
        col: if cell.col_abs {
            cell.col
        } else {
            ((cell.col as i64 + col_delta).max(0)) as u32
        },
        row: if cell.row_abs {
            cell.row
        } else {
            ((cell.row as i64 + row_delta).max(0)) as u32
        },
        col_abs: cell.col_abs,
        row_abs: cell.row_abs,
    }
}

fn shift_col(col: ColRef, col_delta: i64) -> ColRef {
    ColRef {
        col: if col.abs {
            col.col
        } else {
            ((col.col as i64 + col_delta).max(0)) as u32
        },
        abs: col.abs,
    }
}

fn shift_row(row: RowRef, row_delta: i64) -> RowRef {
    RowRef {
        row: if row.abs {
            row.row
        } else {
            ((row.row as i64 + row_delta).max(0)) as u32
        },
        abs: row.abs,
    }
}

fn render_reference(reference: &Reference) -> String {
    let mut out = String::new();
    if let Some(sheet) = &reference.sheet {
        if sheet.chars().all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.')) {
            out.push_str(sheet);
        } else {
            out.push('\'');
            out.push_str(&sheet.replace('\'', "''"));
            out.push('\'');
        }
        out.push('!');
    }
    match &reference.kind {
        RefKind::Cell(cell) => out.push_str(&render_cell(*cell)),
        RefKind::Range(a, b) => {
            out.push_str(&render_cell(*a));
            out.push(':');
            out.push_str(&render_cell(*b));
        }
        RefKind::ColRange(a, b) => {
            out.push_str(&render_col(*a));
            out.push(':');
            out.push_str(&render_col(*b));
        }
        RefKind::RowRange(a, b) => {
            out.push_str(&render_row(*a));
            out.push(':');
            out.push_str(&render_row(*b));
        }
    }
    out
}

fn render_cell(cell: CellRef) -> String {
    format!(
        "{}{}{}{}",
        if cell.col_abs { "$" } else { "" },
        column_index_to_letters(cell.col),
        if cell.row_abs { "$" } else { "" },
        cell.row + 1
    )
}

fn render_col(col: ColRef) -> String {
    format!(
        "{}{}",
        if col.abs { "$" } else { "" },
        column_index_to_letters(col.col)
    )
}

fn render_row(row: RowRef) -> String {
    format!("{}{}", if row.abs { "$" } else { "" }, row.row + 1)
}

#[cfg(test)]
mod tests {
    use super::*;
    use lo_core::{Cell, CellStyle, Metadata};

    #[test]
    fn shared_formula_translation_handles_sheet_and_whole_column() {
        let from = CellAddr::new(0, 0);
        let to = CellAddr::new(4, 2);
        let formula = "SUM(Sheet2!A:A)+Sheet2!B1";
        let shifted = translate_shared_formula(formula, from, to);
        assert!(shifted.contains("Sheet2!C:C"));
        assert!(shifted.contains("Sheet2!D5"));
    }

    #[test]
    fn evaluator_reads_cross_sheet_cells() {
        let mut workbook = Workbook {
            meta: Metadata::default(),
            sheets: vec![Sheet::new("Main"), Sheet::new("Data")],
            named_ranges: Vec::new(),
        };
        workbook.sheets[1].cells.insert(
            CellAddr::new(0, 0),
            Cell {
                value: CellValue::Number(41.0),
                style: Some(CellStyle::default()),
                comment: None,
            },
        );
        let eval = WorkbookEvaluator::new(&workbook);
        let value = eval.evaluate_formula(0, "=Data!A1+1").unwrap();
        assert_eq!(value, EvalValue::Number(42.0));
    }
}
