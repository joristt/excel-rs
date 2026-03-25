use std::io::{Seek, Write};

use crate::error::{ExcelError, Result};
use zip::{write::SimpleFileOptions, ZipWriter};

const MAX_ROWS_BUFFER: u32 = 100_000;

/// A single worksheet within an XLSX workbook.
///
/// Obtained from [`WorkBook::new_worksheet`](crate::WorkBook::new_worksheet).
/// Write rows with [`write_row`](Sheet::write_row), then call
/// [`close`](Sheet::close) when finished.
///
/// # Example
///
/// ```no_run
/// # use excel_rs::{WorkBook, sheet::CellType};
/// # use std::io::Cursor;
/// # let mut wb = WorkBook::new(Cursor::new(vec![]));
/// let mut sheet = wb.new_worksheet("My Sheet".to_string()).unwrap();
/// sheet.write_row([b"Col A".as_ref(), b"Col B"].into_iter(), None).unwrap();
/// sheet.close().unwrap();
/// ```
pub struct Sheet<'a, W: Write + Seek> {
    /// Name of this sheet as it appears in the Excel tab bar.
    pub name: String,
    current_row_num: u32,
    sheet_buf: &'a mut ZipWriter<W>,
    global_shared_vec: Vec<u8>,
}

/// The data type of a cell value.
///
/// Pass a slice of `CellType` as `type_hints` to [`Sheet::write_row`] to
/// control how each column is encoded in the XLSX file.  When in doubt, use
/// [`CellType::String`].
#[derive(Clone, Debug)]
pub enum CellType {
    /// Plain text. Written as `<c t="str">`.
    String,
    /// Date/time stored as a number with a date format applied. Written as `<c t="n" s="1">`.
    Date,
    /// Boolean. Written as `<c t="b">`.
    Boolean,
    /// Numeric (integer or float). Written as `<c t="n">`.
    Number,
}

impl CellType {
    #[inline(always)]
    fn as_static_bytes(&self) -> &'static [u8] {
        match self {
            CellType::String => b"str",
            CellType::Date => b"n\" s=\"1",
            CellType::Boolean => b"b",
            CellType::Number => b"n",
        }
    }
}

fn write_escaped(out: &mut Vec<u8>, bytes: &[u8]) {
    // just learnt of SIMD instructions and this resulted in ~5% perf boost
    // i'm assuming that cells needing escapes are relatively rarer than cells containing normal text
    if memchr::memchr3(b'<', b'>', b'&', bytes).is_none()
    // && memchr::memchr2(b'&', b'"', bytes).is_none()
    {
        out.extend_from_slice(bytes);
        return;
    }

    let mut start = 0;
    for (i, &b) in bytes.iter().enumerate() {
        let escape: &[u8] = match b {
            b'<' => b"&lt;",
            b'>' => b"&gt;",
            // b'\'' => b"&apos;",
            b'&' => b"&amp;",
            // b'"' => b"&quot;",
            _ => continue,
        };
        out.extend_from_slice(&bytes[start..i]);
        out.extend_from_slice(escape);
        start = i + 1;
    }
    out.extend_from_slice(&bytes[start..]);
}

impl<'a, W: Write + Seek> Sheet<'a, W> {
    /// Returns the 1-based index of the last row written (0 before any rows are written).
    #[inline]
    pub fn current_row(&self) -> u32 {
        self.current_row_num
    }

    pub(crate) fn new(name: String, id: u16, writer: &'a mut ZipWriter<W>) -> Result<Self> {
        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .compression_level(Some(1))
            .large_file(true);

        writer.start_file(format!("xl/worksheets/sheet{}.xml", id), options)?;
        writer.write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData>")?;

        Ok(Sheet {
            sheet_buf: writer,
            name,
            current_row_num: 0,
            global_shared_vec: Vec::with_capacity(64 * 1024 * 1024),
        })
    }

    pub fn write_row<'b>(
        &mut self,
        cells: impl Iterator<Item = &'b [u8]>,
        type_hints: Option<&[CellType]>,
    ) -> Result<()> {
        self.current_row_num += 1;

        if self.current_row_num > crate::MAX_ROWS {
            return Err(ExcelError::RowLimitExceeded);
        }

        self.global_shared_vec.extend_from_slice(b"<row>");

        for (col, cell) in cells.enumerate() {
            if col >= crate::MAX_COLS {
                return Err(ExcelError::ColumnLimitExceeded);
            }

            self.global_shared_vec.extend_from_slice(b"<c t=\"");
            self.global_shared_vec.extend_from_slice(type_hints.map_or(
                CellType::String.as_static_bytes(),
                |x| {
                    x.get(col)
                        .map_or(CellType::String.as_static_bytes(), |x| x.as_static_bytes())
                },
            ));

            self.global_shared_vec.extend_from_slice(b"\"><v>");
            write_escaped(&mut self.global_shared_vec, cell);
            self.global_shared_vec.extend_from_slice(b"</v></c>");
        }

        self.global_shared_vec.extend_from_slice(b"</row>");

        if self.current_row_num.is_multiple_of(MAX_ROWS_BUFFER) {
            self.flush()?;
        }

        Ok(())
    }

    fn flush(&mut self) -> Result<()> {
        self.sheet_buf.write_all(&self.global_shared_vec)?;
        self.global_shared_vec.clear();
        Ok(())
    }

    pub fn close(&mut self) -> Result<()> {
        self.flush()?;
        self.sheet_buf.write_all(b"</sheetData></worksheet>")?;
        Ok(())
    }
}
