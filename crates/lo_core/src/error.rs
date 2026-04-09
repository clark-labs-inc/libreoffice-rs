use std::error::Error;
use std::fmt::{Display, Formatter};

pub type Result<T> = std::result::Result<T, LoError>;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LoError {
    Io(String),
    Parse(String),
    InvalidInput(String),
    Unsupported(String),
    Eval(String),
}

impl Display for LoError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(message) => write!(f, "I/O error: {message}"),
            Self::Parse(message) => write!(f, "parse error: {message}"),
            Self::InvalidInput(message) => write!(f, "invalid input: {message}"),
            Self::Unsupported(message) => write!(f, "unsupported: {message}"),
            Self::Eval(message) => write!(f, "evaluation error: {message}"),
        }
    }
}

impl Error for LoError {}

impl From<std::io::Error> for LoError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value.to_string())
    }
}
