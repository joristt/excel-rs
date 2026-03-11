//! A fast, low-allocation library for writing XLSX (Excel Open XML) files.
//!
//! `excel-rs-xlsx` streams cell data directly into a ZIP-compressed XLSX
//! archive, flushing to the underlying writer in batches.  It never
//! materialises the full file in memory, which makes it suitable for large
//! datasets.
//!
//! # Quick start
//!
//! ```no_run
//! use std::fs::File;
//! use excel_rs::{WorkBook, sheet::CellType};
//!
//! let file = File::create("output.xlsx").unwrap();
//! let mut wb = WorkBook::new(file);
//!
//! let mut sheet = wb.new_worksheet("Sales".to_string()).unwrap();
//! sheet.write_row([b"Name".as_ref(), b"Revenue"].into_iter(), None).unwrap();
//! sheet.write_row([b"Alice".as_ref(), b"42000"].into_iter(), Some(&[CellType::String, CellType::Number])).unwrap();
//! sheet.close().unwrap();
//!
//! wb.finish().unwrap();
//! ```
//!
//! # Limits
//!
//! These match the OOXML / Excel specification:
//!
//! | Limit | Value |
//! |---|---|
//! | Rows per sheet | 1 048 576 |
//! | Columns per sheet | 16 384 |
//! | Sheets per workbook | 65 535 |

pub mod error;
mod format;
pub mod sheet;
pub mod workbook;

pub use error::{ExcelError, Result};
pub use workbook::WorkBook;

/// Maximum number of rows per worksheet (Excel limit).
pub const MAX_ROWS: u32 = 1048576; // 2^20

/// Maximum number of columns per worksheet (Excel limit).
pub const MAX_COLS: usize = 16384; // 2^14

/// Maximum number of worksheets per workbook (Excel limit).
pub const MAX_SHEETS: usize = u16::MAX as usize; // 2^16
