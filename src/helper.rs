use regex::Regex;

use std::str::FromStr;

use crate::error::ParseError;

lazy_static! {
    static ref RE2: Regex = Regex::new(r"^(?P<x>[[:upper:]]+)$").unwrap();
}

static UPPER_CHARS: [char; 26] = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
    'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
];

#[derive(Debug, PartialEq, Clone)]
pub struct AlphaNumber {
    pub number: u32,
    pub letters: String,
}

impl From<u32> for AlphaNumber {
    fn from(x: u32) -> AlphaNumber {
        let mut letters = String::new();

        unpack_integer_to_chars(x + 1, &mut letters);

        let letters: String = letters.chars().rev().collect();

        AlphaNumber { number: x, letters }
    }
}

/// function that converts one-based numbers to excel like column value into the given string in reverse order
fn unpack_integer_to_chars(x: u32, txt: &mut String) {
    let mut t = x / 26;
    let mut r = x - t * 26;

    if r == 0 {
        r = 26;
        t = t - 1;
    }

    txt.push(UPPER_CHARS[r as usize - 1]);

    if t != 0 {
        unpack_integer_to_chars(t, txt);
    }
}

impl FromStr for AlphaNumber {
    type Err = ParseError;
    fn from_str(s: &str) -> Result<AlphaNumber, ParseError> {
        let s = s.to_uppercase();
        // let caps = RE2.captures(&s).unwrap();
        let caps = match RE2.captures(&s) {
            Some(x) => x,
            None => return Err(ParseError::InvalidFormat),
        };
        let x = caps.name("x").expect("Regex 'RE2' is not correct").as_str();
        let mut sum: u32 = 0;

        for (i, c) in x.chars().rev().enumerate() {
            let p = UPPER_CHARS
                .iter()
                .position(|x| x == &c)
                .expect("Regex 'RE2' is not correct");
            let q = match i {
                0 => 0,
                _ => 1,
            };
            let power = match 26u32.checked_pow(i as u32) {
                Some(x) => x,
                None => return Err(ParseError::Overflow),
            };
            sum += match ((p + q) as u32).checked_mul(power) {
                Some(x) => x,
                None => return Err(ParseError::Overflow),
            };
        }

        Ok(AlphaNumber {
            number: sum,
            letters: s,
        })
    }
}

#[cfg(test)]
mod tests {
    use crate::error::ParseError;
    use crate::helper::AlphaNumber;

    #[test]
    fn alpha_number_from_string_success() {
        let t: AlphaNumber = "A".parse().unwrap();
        assert_eq!(t.number, 0);

        let t: AlphaNumber = "B".parse().unwrap();
        assert_eq!(t.number, 1);

        let t: AlphaNumber = "Z".parse().unwrap();
        assert_eq!(t.number, 25);

        let t: AlphaNumber = "AA".parse().unwrap();
        assert_eq!(t.number, 26);

        let t: AlphaNumber = "ZZ".parse().unwrap();
        assert_eq!(t.number, 701);

        let t: AlphaNumber = "AAA".parse().unwrap();
        assert_eq!(t.number, 702);

        let t: AlphaNumber = "AAB".parse().unwrap();
        assert_eq!(t.number, 703);

        let t: AlphaNumber = "AMH".parse().unwrap();
        assert_eq!(t.number, 1021);

        let t: AlphaNumber = "b".parse().unwrap();
        assert_eq!(t.number, 1);
    }

    #[test]
    fn alpha_number_from_string_error() {
        let t: Result<AlphaNumber, _> = "123".parse();
        assert_eq!(t, Err(ParseError::InvalidFormat));
        let t: Result<AlphaNumber, _> = "?test".parse();
        assert_eq!(t, Err(ParseError::InvalidFormat));
        let t: Result<AlphaNumber, _> = "ERR8".parse();
        assert_eq!(t, Err(ParseError::InvalidFormat));
        let t: Result<AlphaNumber, _> = "ABCDEFGHIJKLMNOP".parse();
        assert_eq!(t, Err(ParseError::Overflow));
    }

    #[test]
    fn alpha_number_from_u32_success() {
        let t: AlphaNumber = 0.into();
        assert_eq!(t.letters, String::from("A"));

        let t: AlphaNumber = 1.into();
        assert_eq!(t.letters, String::from("B"));

        let t: AlphaNumber = 25.into();
        assert_eq!(t.letters, String::from("Z"));

        let t: AlphaNumber = 26.into();
        assert_eq!(t.letters, String::from("AA"));

        let t: AlphaNumber = 27.into();
        assert_eq!(t.letters, String::from("AB"));

        let t: AlphaNumber = 676.into();
        assert_eq!(t.letters, String::from("ZA"));

        let t: AlphaNumber = 701.into();
        assert_eq!(t.letters, String::from("ZZ"));

        let t: AlphaNumber = 702.into();
        assert_eq!(t.letters, String::from("AAA"));

        let t: AlphaNumber = 703.into();
        assert_eq!(t.letters, String::from("AAB"));

        let t: AlphaNumber = 1021.into();
        assert_eq!(t.letters, String::from("AMH"));
    }
}
