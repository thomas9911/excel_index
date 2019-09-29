use crate::error::ParseError;
use crate::helper::AlphaNumber;

use std::cmp::Ordering;

lazy_static! {
    static ref RE: Regex = Regex::new("^(?P<x>[[:upper:]]+)(?P<y>[[:digit:]]+)$").unwrap();
}

use regex::Regex;
use std::str::FromStr;

/// zero indexed Excel like column row indexes. (examples A3, J18)
#[derive(Debug, PartialEq, PartialOrd, Clone)]
pub struct ExcelIndex {
    pub coordinates: (u32, u32),
    pub letters: String,
}

impl ExcelIndex {
    pub fn from_tuple(x: u32, y: u32) -> ExcelIndex {
        let tmp = AlphaNumber::from(x);
        ExcelIndex {
            coordinates: (x, y),
            letters: format!("{}{}", tmp.letters, y + 1),
        }
    }

    /// creates an infinite iterator
    pub fn into_iter(self) -> ExcelIndexRange {
        ExcelIndexRange::new(self)
    }

    /// creates an iterator over the range, inclusive
    pub fn into_range(self, end: ExcelIndex) -> ExcelIndexRange {
        ExcelIndexRange::range(self, end)
    }
}

impl From<(u32, u32)> for ExcelIndex {
    fn from(tup: (u32, u32)) -> ExcelIndex {
        ExcelIndex::from_tuple(tup.0, tup.1)
    }
}

impl From<ExcelIndex> for (u32, u32) {
    fn from(ei: ExcelIndex) -> (u32, u32) {
        ei.coordinates
    }
}

impl From<ExcelIndex> for String {
    fn from(ei: ExcelIndex) -> String {
        ei.letters
    }
}

impl FromStr for ExcelIndex {
    type Err = ParseError;
    fn from_str(s: &str) -> Result<ExcelIndex, ParseError> {
        let s = s.to_uppercase();
        let caps = match RE.captures(&s) {
            Some(x) => x,
            None => return Err(ParseError::InvalidFormat),
        };
        let x: AlphaNumber = caps
            .name("x")
            .expect("Regex 'RE' is not correct")
            .as_str()
            .parse()?;
        let y: u32 = caps
            .name("y")
            .expect("Regex 'RE' is not correct")
            .as_str()
            .parse()
            .expect("Regex 'RE' is not correct");
        Ok(ExcelIndex {
            coordinates: (x.number, y - 1),
            letters: s,
        })
    }
}

impl std::fmt::Display for ExcelIndex {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.letters)
    }
}

#[derive(Debug)]
pub struct ExcelIndexRange {
    current: ExcelIndex,
    start: ExcelIndex,
    end: Option<ExcelIndex>,
}

impl ExcelIndexRange {
    pub fn new(current: ExcelIndex) -> ExcelIndexRange {
        ExcelIndexRange {
            start: current.clone(),
            current,
            end: None,
        }
    }

    pub fn range(start: ExcelIndex, end: ExcelIndex) -> ExcelIndexRange {
        ExcelIndexRange {
            current: start.clone(),
            start,
            end: Some(end),
        }
    }

    pub fn range_bounded(start: ExcelIndex, end: ExcelIndex) -> ExcelIndexRange {
        ExcelIndexRange::range(start, end)
    }

    pub fn range_unbounded(start: ExcelIndex) -> ExcelIndexRange {
        ExcelIndexRange::new(start)
    }
}

impl Iterator for ExcelIndexRange {
    type Item = ExcelIndex;
    fn next(&mut self) -> Option<Self::Item> {
        if let Some(end) = &self.end {
            if end.coordinates < self.current.coordinates {
                return None;
            } else {
                let (x, y): (u32, u32) = end.clone().into();
                let (r, t): (u32, u32) = self.current.clone().into();
                match x.cmp(&r) {
                    Ordering::Equal => match y.cmp(&t) {
                        Ordering::Less => return None,
                        Ordering::Equal => {
                            let next: ExcelIndex = (r + 1, t).into();
                            let current = self.current.clone();
                            self.current = next;
                            return Some(current);
                        }
                        Ordering::Greater => {
                            let (s, _): (u32, u32) = self.start.clone().into();
                            let next: ExcelIndex = (s, t + 1).into();
                            let current = self.current.clone();
                            self.current = next;
                            return Some(current);
                        }
                    },
                    Ordering::Greater => {
                        let next: ExcelIndex = (r + 1, t).into();
                        let current = self.current.clone();
                        self.current = next;
                        return Some(current);
                    }
                    Ordering::Less => return None,
                }
            }
        } else {
            let (x, y): (u32, u32) = self.current.clone().into();
            let next: ExcelIndex = (x + 1, y).into();
            let current = self.current.clone();
            self.current = next;
            return Some(current);
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::ExcelIndex;
    #[test]
    fn parse_easy_str() {
        let expected = ExcelIndex {
            coordinates: (0, 1),
            letters: String::from("A2"),
        };
        let ey: ExcelIndex = "A2".parse().unwrap();
        assert_eq!(expected, ey);
    }

    #[test]
    fn parse_easy_str_lowercase() {
        let expected = ExcelIndex {
            coordinates: (0, 1),
            letters: String::from("A2"),
        };
        let ey: ExcelIndex = "a2".parse().unwrap();
        assert_eq!(expected, ey);
    }

    #[test]
    fn parse_easy_tuple() {
        let expected = ExcelIndex {
            coordinates: (0, 1),
            letters: String::from("A2"),
        };
        let ey: ExcelIndex = (0, 1).into();
        assert_eq!(expected, ey);
    }

    #[test]
    fn parse_hard_str() {
        let expected = ExcelIndex {
            coordinates: (699, 1233),
            letters: String::from("ZX1234"),
        };
        let ey: ExcelIndex = "ZX1234".parse().unwrap();
        assert_eq!(expected, ey);
    }

    #[test]
    fn parse_hard_tuple() {
        let expected = ExcelIndex {
            coordinates: (699, 1233),
            letters: String::from("ZX1234"),
        };
        let ey: ExcelIndex = (699, 1233).into();
        assert_eq!(expected, ey);
    }

    #[test]
    fn into_tuple() {
        let s: ExcelIndex = "B12".parse().unwrap();
        let t: (u32, u32) = s.into();

        assert_eq!(t, (1, 11));
    }

    #[test]
    fn into_string() {
        let e: ExcelIndex = (1, 11).into();
        let s: String = e.into();

        assert_eq!(s, String::from("B12"));
    }

    #[test]
    fn invertible_test() {
        let excel_coordinate = String::from("F81");
        let e1: ExcelIndex = excel_coordinate.parse().unwrap();
        let t: (u32, u32) = e1.into();
        let e2: ExcelIndex = t.into();
        let s: String = e2.into();

        assert_eq!(excel_coordinate, s);
    }

    #[test]
    fn test_into_range() {
        let expected = vec![
            "A4", "B4", "C4", "D4", "A5", "B5", "C5", "D5", "A6", "B6", "C6", "D6",
        ];

        let e1: ExcelIndex = "A4".parse().unwrap();
        let e2: ExcelIndex = "D6".parse().unwrap();

        let result: Vec<String> = e1.into_range(e2).map(|x| x.to_string()).collect();

        assert_eq!(expected, result);
    }

    #[test]
    fn test_into_iter() {
        let expected = vec!["A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4", "I4"];

        let e1: ExcelIndex = "A4".parse().unwrap();

        let result: Vec<String> = e1.into_iter().take(9).map(|x| x.to_string()).collect();

        assert_eq!(expected, result);
    }

}
