mod celltype;
mod error;
mod workbook;

use pyo3::prelude::*;

use crate::{celltype::PyCellType, workbook::PyWorkBook};

/// High-performance CSV / NumPy to XLSX converter.
///
/// This module exposes two public symbols:
///
/// * `py_excel_rs.WorkBook` — the main workbook builder.
/// * `py_excel_rs.CellType` — an enum for column type hints.
#[pymodule]
fn _excel_rs<'py>(m: &Bound<'py, PyModule>) -> PyResult<()> {
    m.add_class::<PyWorkBook>()?;
    m.add_class::<PyCellType>()?;

    Ok(())
}
