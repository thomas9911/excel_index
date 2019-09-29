//! Zero indexed Excel like column row indexes. (examples A3, J18)
//!
//!
//! ```
//! use excel_index::ExcelIndex;
//!
//! use std::str::FromStr;
//!
//! let A1: ExcelIndex = (0, 0).into();
//! let A6 = ExcelIndex::from((0, 5));
//! let B1 = ExcelIndex::from_tuple(1, 0);
//! let excel_index_with_tuple_two_six = ExcelIndex::from_str("C7").unwrap();
//!
//! ```
//!
//! ```
//! use excel_index::ExcelIndex;
//!
//! use std::str::FromStr;
//!
//! // creates an iterator from A1 to F4 (A1, B1, C1, D1, E1, F1, A2, B2, ..., E4, F4)
//! for cell in ExcelIndex::from((0, 0)).into_range(ExcelIndex::from_str("F4").unwrap()) {
//!     println!("{}", cell);
//! }
//! ```

#[macro_use]
extern crate lazy_static;
extern crate regex;

pub mod error;
pub mod helper;
mod indexes;

pub use indexes::{ExcelIndex, ExcelIndexRange};
