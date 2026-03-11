use std::io::Cursor;

use csv::ByteRecord;
use excel_rs::{sheet::CellType, WorkBook};
use numpy::{ndarray::ArrayView1, PyArray2, PyArrayMethods};
use pyo3::{
    Bound, Py, PyAny, PyResult, Python, exceptions::{PyRuntimeError, PyTypeError}, pyclass, pymethods, types::{PyAnyMethods, PyBool, PyBytes, PyBytesMethods, PyFloat, PyInt, PyStringMethods}
};

use crate::{celltype::PyCellType, error::to_py_err};

/// An in-memory Excel workbook builder.
///
/// ``WorkBook`` accumulates one or more worksheets and serialises the final
/// XLSX file when :meth:`finish` is called.  All data is buffered in memory
/// until that point.
///
/// Example::
///
///     from excel_rs import WorkBook, CellType
///
///     wb = WorkBook()
///
///     # write a CSV sheet
///     with open("data.csv", "rb") as f:
///         wb.write_csv_to_sheet("Raw Data", f.read())
///
///     # write a NumPy sheet with explicit types
///     import numpy as np
///     arr = np.array([["Alice", 30], ["Bob", 25]])
///     wb.write_numpy_to_sheet("People", arr)
///
///     with open("output.xlsx", "wb") as f:
///         wb.finish(f)
///
/// .. note::
///     ``WorkBook`` is **not** thread-safe.  Do not share an instance across
///     threads.
#[pyclass(unsendable)]
pub struct PyWorkBook {
    workbook: Option<excel_rs::WorkBook<Cursor<Vec<u8>>>>,
}

#[pymethods]
impl PyWorkBook {
    /// Create a new empty workbook.
    #[new]
    pub fn new() -> PyWorkBook {
        PyWorkBook {
            workbook: Some(WorkBook::new(Cursor::new(vec![]))),
        }
    }

    /// Write CSV bytes as a new worksheet.
    ///
    /// The CSV data is read directly from `buffer`` without an intermediate
    /// copy. The first row of the CSV becomes the header row in Excel.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet tab.
    ///     buffer: Raw CSV bytes (e.g. `open("file.csv", "rb").read()`).
    ///     type_hints: Optional list of `py_excel_rs.CellType` values, one per
    ///         column, controlling how each column is typed in the workbook.
    ///         Defaults to all-string when omitted.
    ///
    /// Raises:
    ///     RuntimeError: If the workbook has already been closed via `workbook.finish()`,
    ///     or if an underlying error occurred writing to the excel file.
    #[pyo3(signature = (sheet_name, buffer, type_hints=None))]
    pub fn write_csv_to_sheet(
        &mut self,
        sheet_name: String,
        buffer: &Bound<'_, PyBytes>,
        type_hints: Option<Vec<PyCellType>>,
    ) -> PyResult<()> {
        let wb = self
            .workbook
            .as_mut()
            .ok_or_else(|| PyRuntimeError::new_err("workbook already closed"))?;

        let mut sheet = wb.new_worksheet(sheet_name).map_err(to_py_err)?;

        let mut reader = csv::ReaderBuilder::new().from_reader(buffer.as_bytes());

        let headers = reader
            .byte_headers()
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

        sheet.write_row(headers.iter(), None).map_err(to_py_err)?;

        let mut record = ByteRecord::new();

        let type_hints =
            type_hints.map(|x| x.into_iter().map(CellType::from).collect::<Vec<CellType>>());

        while reader
            .read_byte_record(&mut record)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?
        {
            sheet
                .write_row(record.iter(), type_hints.as_deref())
                .map_err(to_py_err)?;
        }

        sheet.close().map_err(to_py_err)?;

        Ok(())
    }

    /// Write a NumPy 2-D array or Pandas DataFrame as a new worksheet.
    ///
    /// Supported dtypes:
    ///
    /// * `f` (float64) — written as numbers.
    /// * `i` / `u` (int64 / uint64) — written as numbers.
    /// * `b` (bool) — written as booleans.
    /// * `U` / `S` (unicode / byte string) — written as strings.
    /// * `O` (object) — each value is inspected at runtime; pass
    ///   `should_infer_types_from_first_row=True` to detect numeric, boolean,
    ///   and datetime columns automatically from the first row.
    ///
    /// Args:
    ///     sheet_name: Name of the worksheet tab.
    ///     array: A 2-D NumPy array or pandas `pd.DataFrame`.
    ///     should_infer_types_from_first_row: When `True` and the array dtype
    ///         is `object`, inspect the first row to infer per-column
    ///         `py_excel_rs.CellType` values.  Has no effect for typed dtypes.
    ///         Defaults to `False`.
    ///
    /// Raises:
    ///     TypeError: If `array`` is not 2-D or has an unsupported dtype.
    ///     RuntimeError: If the workbook has already been closed.
    #[pyo3(signature = (sheet_name, array, should_infer_types_from_first_row=None))]
    pub fn write_data_to_sheet(
        &mut self,
        sheet_name: String,
        array: &Bound<'_, PyAny>,
        should_infer_types_from_first_row: Option<bool>,
    ) -> PyResult<()> {
        // handles pandas dataframes
        let (array, resolved_headers) = if array.hasattr("columns")? {
            let cols = array
                .getattr("columns")?
                .try_iter()?
                .map(|c| c?.str()?.to_str().map(|s| s.to_string()))
                .collect::<PyResult<Vec<_>>>()?;
            let values = array.getattr("values")?;
            (values, Some(cols))
        } else {
            // this doesn't clone the array, apparently it just creates another ref to the array
            (array.clone(), None)
        };

        let ndim = array.getattr("ndim")?.extract::<usize>()?;
        if ndim != 2 {
            return Err(PyTypeError::new_err(format!(
                "expected 2d array, got {}d array",
                ndim
            )));
        }

        let dtype = array
            .getattr("dtype")?
            .getattr("kind")?
            .extract::<String>()?;

        let should_infer_types = should_infer_types_from_first_row.unwrap_or(false);
        let num_of_cols = array.getattr("shape")?.extract::<(usize, usize)>()?.1;

        match dtype.as_str() {
            "f" => self.write_typed_array::<f64>(
                sheet_name,
                &array,
                resolved_headers,
                Some(vec![CellType::Number; num_of_cols]),
            ),
            "i" | "u" => self.write_typed_array::<i64>(
                sheet_name,
                &array,
                resolved_headers,
                Some(vec![CellType::Number; num_of_cols]),
            ),
            "b" => self.write_typed_array::<bool>(
                sheet_name,
                &array,
                resolved_headers,
                Some(vec![CellType::Boolean; num_of_cols]),
            ),
            "U" | "S" => {
                self.write_typed_array::<Py<PyAny>>(sheet_name, &array, resolved_headers, None)
            }
            "O" => {
                let type_hints = if should_infer_types {
                    Some(Python::attach(|py| {
                        infer_types_from_object_array(py, &array)
                    })?)
                } else {
                    None
                };
                self.write_typed_array::<Py<PyAny>>(sheet_name, &array, resolved_headers, type_hints)
            }
            _ => Err(PyTypeError::new_err(format!(
                "unsupported dtype kind: {}",
                dtype
            ))),
        }
    }

    /// Finalise the workbook and write the XLSX bytes to `output``.
    ///
    /// `output` must be a writable binary file-like object (anything with a
    /// `write(bytes)` method, e.g. an open file, `io.BytesIO`, an HTTP
    /// response body, etc.).
    ///
    /// After this call the `WorkBook` is consumed and cannot be used again.
    ///
    /// Args:
    ///     output: A writable binary file-like object.
    ///
    /// Raises:
    ///     RuntimeError: If `workbook.finish()` has already been called, or if
    ///         serialisation fails.
    ///
    /// Example::
    ///
    ///     import io
    ///     buf = io.BytesIO()
    ///     wb.finish(buf)
    ///     xlsx_bytes = buf.getvalue()
    pub fn finish(&mut self, py: Python<'_>, output: Bound<'_, PyAny>) -> PyResult<()> {
        let wb = self
            .workbook
            .take()
            .ok_or_else(|| PyRuntimeError::new_err("workbook already closed"))?;

        let final_buffer = wb.finish().map_err(to_py_err)?;
        output.call_method1("write", (PyBytes::new(py, final_buffer.get_ref()),))?;

        Ok(())
    }
}

impl PyWorkBook {
    fn write_typed_array<T>(
        &mut self,
        sheet_name: String,
        array: &Bound<'_, PyAny>,
        headers: Option<Vec<String>>,
        type_hints: Option<Vec<CellType>>,
    ) -> PyResult<()>
    where
        T: numpy::Element + ToString + Clone,
    {
        let wb = self
            .workbook
            .as_mut()
            .ok_or_else(|| PyRuntimeError::new_err("workbook already closed"))?;

        let downcasted_arr = array.cast::<PyArray2<T>>()?.readonly();
        let ndarray = downcasted_arr.as_array();

        let mut sheet = wb.new_worksheet(sheet_name).map_err(to_py_err)?;

        if let Some(headers) = headers {
            sheet
                .write_row(headers.iter().map(|x| x.as_bytes()), None)
                .map_err(to_py_err)?;
        }

        for row in ndarray.rows() {
            // Collect per-row to avoid allocating the full array upfront.
            // Each Vec<u8> is freed once the row is written.
            let cells: Vec<Vec<u8>> = row.iter().map(|x| x.to_string().into_bytes()).collect();
            sheet
                .write_row(cells.iter().map(|x| x.as_slice()), type_hints.as_deref())
                .map_err(to_py_err)?;
        }

        sheet.close().map_err(to_py_err)?;

        Ok(())
    }
}

fn infer_types_from_object_array(py: Python, array: &Bound<'_, PyAny>) -> PyResult<Vec<CellType>> {
    let arr = array.cast::<PyArray2<Py<PyAny>>>()?.readonly();
    let casted_arr = arr.as_array();
    Ok(infer_types_from_first_row(py, casted_arr.row(0)))
}

fn infer_types_from_first_row(py: Python<'_>, row: ArrayView1<Py<PyAny>>) -> Vec<CellType> {
    let datetime_type = py
        .import("datetime")
        .and_then(|m| m.getattr("datetime"))
        .ok();

    row.iter()
        .map(|obj: &Py<PyAny>| infer_cell_type(py, obj.bind(py), datetime_type.as_ref()))
        .collect()
}

fn infer_cell_type(
    _py: Python<'_>,
    obj: &Bound<'_, PyAny>,
    datetime_type: Option<&Bound<'_, PyAny>>,
) -> CellType {
    if obj.is_instance_of::<PyBool>() {
        CellType::Boolean
    } else if obj.is_instance_of::<PyInt>() || obj.is_instance_of::<PyFloat>() {
        CellType::Number
    } else if let Some(dt) = datetime_type {
        if obj.is_instance(dt).unwrap_or(false) {
            CellType::Date
        } else {
            CellType::String
        }
    } else {
        CellType::String
    }
}
