# excel-rs-xlsx

A fast, low-allocation Rust library for writing XLSX (Excel Open XML) files.

`excel-rs-xlsx` is the core engine behind [`py-excel-rs`](https://pypi.org/project/py-excel-rs/) and `cli-excel-rs`. It streams cell data directly into a ZIP-compressed XLSX archive, flushing to the underlying writer in batches — it never materialises the full file in memory.

## Features

- Streams directly to any `Write + Seek` target (file, `Cursor<Vec<u8>>`, etc.)
- SIMD-accelerated XML escaping via `memchr`
- Compile-time column-letter lookup table (zero runtime allocation for A1 references)
- Supports strings, numbers, booleans, and dates
- Named worksheets
- Row/column/sheet limit enforcement matching the OOXML spec

## Usage

```rust
use std::fs::File;
use excel_rs::{WorkBook, sheet::CellType};

let file = File::create("output.xlsx")?;
let mut wb = WorkBook::new(file);

let mut sheet = wb.new_worksheet("Sales".to_string())?;

// Write a header row (all strings by default)
sheet.write_row([b"Name".as_ref(), b"Revenue"].into_iter(), None)?;

// Write a data row with explicit type hints
sheet.write_row(
    [b"Alice".as_ref(), b"42000"].into_iter(),
    Some(&[CellType::String, CellType::Number]),
)?;

sheet.close()?;
wb.finish()?;
```

## Limits

| Resource | Maximum |
|---|---|
| Rows per worksheet | 1 048 576 |
| Columns per worksheet | 16 384 |
| Worksheets per workbook | 65 535 |

## License

MIT
