//! CLI Tool Example
//!
//! This example demonstrates how to build a command-line tool
//! using xlsxzero for converting Excel files to Markdown.

use std::fs::File;
use std::io::{self, Write};
use std::process;
use xlsxzero::{ConverterBuilder, SheetSelector, XlsxToMdError};

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() < 3 {
        eprintln!("Usage: {} <input.xlsx> <output.md> [options]", args[0]);
        eprintln!("\nOptions:");
        eprintln!("  --sheet-index <n>    Select sheet by index (0-based)");
        eprintln!("  --sheet-name <name>  Select sheet by name");
        eprintln!("  --all-sheets        Convert all sheets (default)");
        eprintln!("  --stdout            Write output to stdout instead of file");
        eprintln!("\nExamples:");
        eprintln!("  {} input.xlsx output.md", args[0]);
        eprintln!("  {} input.xlsx output.md --sheet-index 0", args[0]);
        eprintln!("  {} input.xlsx output.md --sheet-name \"Sheet1\"", args[0]);
        eprintln!("  {} input.xlsx - --stdout", args[0]);
        process::exit(1);
    }

    let input_path = &args[1];
    let output_path = &args[2];
    let use_stdout = output_path == "-" || args.contains(&"--stdout".to_string());

    // Parse options
    let mut sheet_selector = SheetSelector::All;
    let mut i = 3;
    while i < args.len() {
        match args[i].as_str() {
            "--sheet-index" => {
                if i + 1 >= args.len() {
                    eprintln!("Error: --sheet-index requires a value");
                    process::exit(1);
                }
                let index = args[i + 1].parse::<usize>().unwrap_or_else(|_| {
                    eprintln!("Error: Invalid sheet index: {}", args[i + 1]);
                    process::exit(1);
                });
                sheet_selector = SheetSelector::Index(index);
                i += 2;
            }
            "--sheet-name" => {
                if i + 1 >= args.len() {
                    eprintln!("Error: --sheet-name requires a value");
                    process::exit(1);
                }
                sheet_selector = SheetSelector::Name(args[i + 1].clone());
                i += 2;
            }
            "--all-sheets" => {
                sheet_selector = SheetSelector::All;
                i += 1;
            }
            "--stdout" => {
                // Already handled above
                i += 1;
            }
            _ => {
                eprintln!("Error: Unknown option: {}", args[i]);
                process::exit(1);
            }
        }
    }

    // Convert Excel file
    match convert_excel(input_path, output_path, &sheet_selector, use_stdout) {
        Ok(_) => {
            if !use_stdout {
                println!("Conversion completed: {} -> {}", input_path, output_path);
            }
        }
        Err(e) => {
            handle_error(e);
            process::exit(1);
        }
    }
}

fn convert_excel(
    input_path: &str,
    output_path: &str,
    sheet_selector: &SheetSelector,
    use_stdout: bool,
) -> Result<(), XlsxToMdError> {
    // Build converter with specified sheet selector
    let converter = ConverterBuilder::new()
        .with_sheet_selector(sheet_selector.clone())
        .build()?;

    // Open input file
    let input = File::open(input_path)?;

    // Handle output
    if use_stdout {
        // Write to stdout
        let stdout = io::stdout();
        let mut handle = stdout.lock();
        converter.convert(input, &mut handle)?;
        handle.flush()?;
    } else {
        // Write to file
        let output = File::create(output_path)?;
        converter.convert(input, output)?;
    }

    Ok(())
}

fn handle_error(error: XlsxToMdError) {
    match error {
        XlsxToMdError::Io(io_err) => {
            eprintln!("I/O Error: {}", io_err);
            eprintln!("Please check that the file exists and you have permission to access it.");
        }
        XlsxToMdError::Parse(parse_err) => {
            eprintln!("Parse Error: {}", parse_err);
            eprintln!("The file may not be a valid Excel file or may be corrupted.");
        }
        XlsxToMdError::Config(msg) => {
            eprintln!("Configuration Error: {}", msg);
            eprintln!("Please check your sheet selection or range specification.");
        }
        XlsxToMdError::UnsupportedFeature {
            sheet,
            cell,
            message,
        } => {
            eprintln!("Unsupported Feature:");
            eprintln!("  Sheet: {}", sheet);
            eprintln!("  Cell: {}", cell);
            eprintln!("  Details: {}", message);
        }
        XlsxToMdError::Utf8(utf8_err) => {
            eprintln!("UTF-8 Conversion Error: {}", utf8_err);
            eprintln!("The file contains invalid UTF-8 characters.");
        }
        XlsxToMdError::Zip(msg) => {
            eprintln!("ZIP Archive Error: {}", msg);
            eprintln!("The file may be corrupted or not a valid ZIP archive.");
        }
        XlsxToMdError::ParseInt(parse_int_err) => {
            eprintln!("Number Parse Error: {}", parse_int_err);
            eprintln!("Failed to parse a number in the file.");
        }
        XlsxToMdError::SecurityViolation(msg) => {
            eprintln!("Security Violation: {}", msg);
            eprintln!("The file violates security constraints (e.g., file size limit).");
        }
    }
}
