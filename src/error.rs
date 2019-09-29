use std::fmt;

#[derive(Debug, PartialEq)]
pub enum ParseError {
    InvalidFormat,
    Overflow,
}

impl fmt::Display for ParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:?}", self)
    }
}
