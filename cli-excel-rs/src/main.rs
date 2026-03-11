use std::{
    fs::File,
    io::{self, Read},
    path::PathBuf,
    process,
};

use anyhow::{bail, Context, Result};
use clap::{Arg, ArgAction, Command};
use excel_rs::WorkBook;

fn main() {
    if let Err(e) = run() {
        eprintln!("error: {e}");
        for cause in e.chain().skip(1) {
            eprintln!("  caused by: {cause}");
        }
        process::exit(1);
    }
}

fn run() -> Result<()> {
    let matches = cli().get_matches();
    match matches.subcommand() {
        Some(("csv", m)) => cmd_csv(m),
        _ => unreachable!(),
    }
}

fn cli() -> Command {
    Command::new("excel-rs")
        .about("Convert data files to XLSX format")
        .version(env!("CARGO_PKG_VERSION"))
        .subcommand_required(true)
        .arg_required_else_help(true)
        .subcommand(
            Command::new("csv")
                .about("Convert a CSV file to XLSX")
                .after_help(
                    "EXAMPLES:\n  \
                     excel-rs csv data.csv\n  \
                     excel-rs csv data.csv -o report.xlsx -s \"Sales Data\"\n  \
                     excel-rs csv data.tsv --tsv -o data.xlsx\n  \
                     cat data.csv | excel-rs csv -o output.xlsx",
                )
                .arg(
                    Arg::new("input")
                        .value_name("INPUT")
                        .help("Input CSV file (omit or use - to read from stdin)")
                        .index(1),
                )
                .arg(
                    Arg::new("output")
                        .short('o')
                        .long("output")
                        .value_name("FILE")
                        .help("Output XLSX file [default: <input>.xlsx]"),
                )
                .arg(
                    Arg::new("sheet")
                        .short('s')
                        .long("sheet")
                        .value_name("NAME")
                        .default_value("Sheet 1")
                        .help("Worksheet name"),
                )
                .arg(
                    Arg::new("delimiter")
                        .short('d')
                        .long("delimiter")
                        .value_name("CHAR")
                        .default_value(",")
                        .help("Field delimiter character")
                        .conflicts_with("tsv"),
                )
                .arg(
                    Arg::new("tsv")
                        .long("tsv")
                        .action(ArgAction::SetTrue)
                        .help("Use tab as delimiter (shorthand for -d '\\t')")
                        .conflicts_with("delimiter"),
                )
                .arg(
                    Arg::new("no-header")
                        .long("no-header")
                        .action(ArgAction::SetTrue)
                        .help("Treat first row as data; do not promote it to a header row"),
                )
                .arg(
                    Arg::new("quote")
                        .short('q')
                        .long("quote")
                        .value_name("CHAR")
                        .default_value("\"")
                        .help("Quote character used in the CSV"),
                )
                .arg(
                    Arg::new("comment")
                        .short('c')
                        .long("comment")
                        .value_name("CHAR")
                        .help("Skip lines beginning with this character"),
                ),
        )
}

fn cmd_csv(m: &clap::ArgMatches) -> Result<()> {
    let input = m.get_one::<String>("input").map(|s| s.as_str());
    let output = m.get_one::<String>("output");
    let sheet_name = m.get_one::<String>("sheet").unwrap().clone();
    let tsv = m.get_flag("tsv");
    let no_header = m.get_flag("no-header");

    let delimiter = if tsv {
        b'\t'
    } else {
        parse_single_ascii_char(m.get_one::<String>("delimiter").unwrap(), "delimiter")?
    };

    let quote = parse_single_ascii_char(m.get_one::<String>("quote").unwrap(), "quote")?;

    let comment = m
        .get_one::<String>("comment")
        .map(|s| parse_single_ascii_char(s, "comment"))
        .transpose()?;

    let output_path: PathBuf = match (input, output) {
        (_, Some(o)) => PathBuf::from(o),
        (Some(i), None) if i != "-" => PathBuf::from(i).with_extension("xlsx"),
        _ => bail!("--output is required when reading from stdin"),
    };

    let reader: Box<dyn Read> = match input {
        None | Some("-") => Box::new(io::stdin()),
        Some(path) => {
            Box::new(File::open(path).with_context(|| format!("failed to open '{path}'"))?)
        }
    };

    let mut csv_reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .quote(quote)
        .comment(comment)
        .has_headers(!no_header)
        .flexible(true)
        .from_reader(reader);

    let output_file = File::create(&output_path)
        .with_context(|| format!("failed to create '{}'", output_path.display()))?;

    let mut workbook = WorkBook::new(output_file);
    let mut worksheet = workbook
        .new_worksheet(sheet_name)
        .context("failed to create worksheet")?;

    if !no_header {
        let headers = csv_reader
            .byte_headers()
            .context("failed to read CSV headers")?
            .clone();
        worksheet
            .write_row(headers.iter(), None)
            .context("failed to write header row")?;
    }

    for result in csv_reader.byte_records() {
        let record = result.context("failed to parse CSV record")?;
        worksheet
            .write_row(record.iter(), None)
            .context("failed to write row")?;
    }

    worksheet.close().context("failed to finalise worksheet")?;
    workbook.finish().context("failed to finalise workbook")?;

    Ok(())
}

fn parse_single_ascii_char(s: &str, name: &str) -> Result<u8> {
    let bytes = s.as_bytes();
    if bytes.len() == 1 && bytes[0].is_ascii() {
        Ok(bytes[0])
    } else {
        bail!("{name} must be a single ASCII character, got {s:?}")
    }
}
