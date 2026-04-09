use crate::meta::Metadata;

#[derive(Clone, Debug, PartialEq)]
pub enum FormulaNode {
    Number(String),
    Identifier(String),
    Symbol(String),
    Operator {
        op: String,
        lhs: Box<FormulaNode>,
        rhs: Box<FormulaNode>,
    },
    Fraction {
        numerator: Box<FormulaNode>,
        denominator: Box<FormulaNode>,
    },
    Superscript {
        base: Box<FormulaNode>,
        exponent: Box<FormulaNode>,
    },
    Subscript {
        base: Box<FormulaNode>,
        subscript: Box<FormulaNode>,
    },
    Group(Vec<FormulaNode>),
}

#[derive(Clone, Debug, PartialEq)]
pub struct FormulaDocument {
    pub meta: Metadata,
    pub root: FormulaNode,
}

impl FormulaDocument {
    pub fn new(title: impl Into<String>, root: FormulaNode) -> Self {
        Self {
            meta: Metadata::titled(title),
            root,
        }
    }
}
