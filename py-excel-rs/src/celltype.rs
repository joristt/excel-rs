use excel_rs::sheet::CellType;
use pyo3::pyclass;

/// The type of data stored in an Excel cell column.
///
/// Pass a list of `CellType` values to `py_excel_rs.WorkBook.write_csv_to_sheet`
/// or `py_excel_rs.WorkBook.write_numpy_to_sheet` to control how each column is
/// formatted in the output workbook.
///
/// When no type hints are provided every column is written as a string.
///
/// Example:
///
///     from excel_rs import WorkBook, CellType
///
///     wb = WorkBook()
///     wb.write_csv_to_sheet("Sales", csv_bytes, [CellType.Number, CellType.String, CellType.Date])
///     with open("out.xlsx", "wb") as f:
///         wb.finish(f)
#[pyclass(eq, eq_int, from_py_object)]
#[derive(PartialEq, Clone)]
pub enum PyCellType {
    /// Plain text. Use this when unsure.
    String,
    /// A date or datetime value stored as an Excel serial number.
    Date,
    /// A boolean (`TRUE` / `FALSE`) cell.
    Boolean,
    /// A numeric value (integer or floating-point).
    Number,
}

impl From<PyCellType> for CellType {
    fn from(value: PyCellType) -> CellType {
        match value {
            PyCellType::String => CellType::String,
            PyCellType::Date => CellType::Date,
            PyCellType::Number => CellType::Number,
            PyCellType::Boolean => CellType::Boolean,
        }
    }
}
