//! Basic Conversion Example
//!
//! This example demonstrates the most basic usage of xlsxzero:
//! converting an Excel file to Markdown using default settings.
//!
//! # Usage
//!
//! ```bash
//! cargo run --example basic_conversion -- input.xlsx output.md
//! ```
//!
//! If no arguments are provided, it will try to use `examples/fixtures/sample.xlsx`
//! as input and `examples/fixtures/output.md` as output.

use std::fs::File;
use xlsxzero::ConverterBuilder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Get input file path from command line arguments or use default
    let input_path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "examples/fixtures/sample.xlsx".to_string());

    // Get output file path from command line arguments or use default
    let output_path = std::env::args()
        .nth(2)
        .unwrap_or_else(|| "examples/fixtures/output.md".to_string());

    println!("Converting {} to {}...", input_path, output_path);

    // Create a converter with default settings
    let converter = ConverterBuilder::new().build()?;

    // Open input file
    let input = File::open(&input_path).map_err(|e| {
        eprintln!("Error: Could not open input file '{}'", input_path);
        eprintln!("  {}", e);
        eprintln!("\nHint: Create a sample Excel file or provide a path to an existing file.");
        e
    })?;

    // Create output file
    let output = File::create(&output_path).map_err(|e| {
        eprintln!("Error: Could not create output file '{}'", output_path);
        eprintln!("  {}", e);
        e
    })?;

    // Convert Excel to Markdown
    converter.convert(input, output)?;

    println!("Conversion completed successfully!");
    println!("Output written to: {}", output_path);

    Ok(())
}
