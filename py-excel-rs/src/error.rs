
use excel_rs::ExcelError;
use pyo3::{exceptions::PyRuntimeError, PyErr};

pub(crate) fn to_py_err(err: ExcelError) -> PyErr {
    PyRuntimeError::new_err(err.to_string())
}