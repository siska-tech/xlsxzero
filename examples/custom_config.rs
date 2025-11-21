//! Custom Configuration Example
//!
//! This example demonstrates how to use custom configuration options
//! such as sheet selection, merge strategy, date format, and formula mode.

use std::fs::File;
use xlsxzero::{ConverterBuilder, DateFormat, FormulaMode, MergeStrategy, SheetSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Get input file path from command line arguments or use default
    let input_path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "examples/fixtures/sample.xlsx".to_string());

    // Get output file path from command line arguments or use default
    let output_path = std::env::args()
        .nth(2)
        .unwrap_or_else(|| "examples/fixtures/output_custom.md".to_string());

    println!(
        "Converting {} to {} with custom settings...",
        input_path, output_path
    );

    // Create a converter with custom settings
    let converter = ConverterBuilder::new()
        // Select only the first sheet (index 0)
        .with_sheet_selector(SheetSelector::Index(0))
        // Use HTML fallback for merged cells
        .with_merge_strategy(MergeStrategy::HtmlFallback)
        // Use Japanese date format
        .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()))
        // Output formula strings instead of cached values
        .with_formula_mode(FormulaMode::Formula)
        // Skip hidden elements
        .include_hidden(false)
        // Limit to range A1:E20 (0-indexed: row 0-19, col 0-4)
        .with_range((0, 0), (19, 4))
        .build()?;

    // Open input file
    let input = File::open(&input_path)?;

    // Create output file
    let output = File::create(&output_path)?;

    // Convert Excel to Markdown
    converter.convert(input, output)?;

    println!("Conversion completed successfully!");
    println!("Output written to: {}", output_path);
    println!("\nCustom settings used:");
    println!("  - Sheet: First sheet only (index 0)");
    println!("  - Merge strategy: HTML fallback");
    println!("  - Date format: Japanese (YYYY年MM月DD日)");
    println!("  - Formula mode: Formula strings");
    println!("  - Hidden elements: Excluded");
    println!("  - Range: A1:E20");

    Ok(())
}
