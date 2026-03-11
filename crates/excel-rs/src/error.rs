use std::fmt;

#[derive(Debug)]
pub enum ExcelError {
    /// An IO error occurred while writing to an underlying writer.
    Io(std::io::Error),
    /// An error occurred zipping the excel file.
    Zip(zip::result::ZipError),
    /// Attempted to create more sheets than the maximum allowed (65535).
    TooManySheets,
    /// Attempted to write more rows than Excel supports (1048576).
    RowLimitExceeded,
    /// Attempted to write more columns than Excel supports (16384).
    ColumnLimitExceeded,
    /// Attempted to write header after already writing rows.
    HeaderAfterRows,
    /// Attempted to create a sheet that already exists.
    SheetAlreadyExists,
}

pub type Result<T> = std::result::Result<T, ExcelError>;

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io(e) => write!(f, "IO error: {e}"),
            Self::Zip(e) => write!(f, "ZIP error: {e}"),
            Self::TooManySheets => {
                write!(f, "exceeded maximum sheet count of {}", u16::MAX)
            }
            Self::RowLimitExceeded => {
                write!(f, "exceeded Excel row limit of {}", 1048576)
            }
            Self::ColumnLimitExceeded => {
                write!(f, "exceeded Excel column limit of {}", 16384)
            }
            Self::HeaderAfterRows => {
                write!(f, "you cannot write a header after already writing rows")
            }
            Self::SheetAlreadyExists => {
                write!(f, "a sheet with that name already exists. sheets need to have unique names")
            }
        }
    }
}

impl std::error::Error for ExcelError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(e) => Some(e),
            Self::Zip(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for ExcelError {
    fn from(e: std::io::Error) -> Self {
        Self::Io(e)
    }
}

impl From<zip::result::ZipError> for ExcelError {
    fn from(e: zip::result::ZipError) -> Self {
        Self::Zip(e)
    }
}
