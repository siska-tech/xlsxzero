#![cfg_attr(not(feature = "std"), no_std)]

//! xlsxzero - Pure-Rust Excel parser and Markdown converter for RAG systems
//!
//! This crate provides functionality to parse Excel files (XLSX) and convert them
//! to structured Markdown format, optimized for RAG (Retrieval-Augmented Generation) systems.
//!
//! # Quick Start
//!
//! ```rust,no_run
//! use std::fs::File;
//! use xlsxzero::ConverterBuilder;
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     // Create a converter with default settings
//!     let converter = ConverterBuilder::new().build()?;
//!
//!     // Open input Excel file
//!     let input = File::open("example.xlsx")?;
//!
//!     // Create output Markdown file
//!     let output = File::create("output.md")?;
//!
//!     // Convert Excel to Markdown
//!     converter.convert(input, output)?;
//!
//!     Ok(())
//! }
//! ```
//!
//! For in-memory conversion, use `Cursor`:
//!
//! ```rust,no_run
//! use std::io::Cursor;
//! use xlsxzero::ConverterBuilder;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let converter = ConverterBuilder::new().build()?;
//! let excel_data: Vec<u8> = vec![]; // Your Excel file bytes
//! let mut markdown_output = Vec::new();
//! converter.convert(Cursor::new(excel_data), &mut markdown_output)?;
//! # Ok(())
//! # }
//! ```
//!
//! # Custom Configuration
//!
//! ```rust,no_run
//! use std::fs::File;
//! use xlsxzero::{ConverterBuilder, SheetSelector, MergeStrategy, DateFormat};
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     // Create a converter with custom settings
//!     let converter = ConverterBuilder::new()
//!         .with_sheet_selector(SheetSelector::Index(0))  // First sheet only
//!         .with_merge_strategy(MergeStrategy::HtmlFallback)  // HTML for merged cells
//!         .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()))  // Japanese format
//!         .build()?;
//!
//!     let input = File::open("example.xlsx")?;
//!     let output = File::create("output.md")?;
//!     converter.convert(input, output)?;
//!
//!     Ok(())
//! }
//! ```
//!
//! # Convert to String
//!
//! ```rust,no_run
//! use std::fs::File;
//! use xlsxzero::ConverterBuilder;
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     let converter = ConverterBuilder::new().build()?;
//!     let input = File::open("example.xlsx")?;
//!
//!     // Convert to String instead of writing to a file
//!     let markdown = converter.convert_to_string(input)?;
//!     println!("{}", markdown);
//!
//!     Ok(())
//! }
//! ```

mod api;
mod builder;
mod error;
mod format;
mod formatter;
mod grid;
mod output;
mod parser;
mod security;
mod types;

// 公開API
pub use api::{DateFormat, FormulaMode, MergeStrategy, OutputFormat, SheetSelector};
pub use builder::{Converter, ConverterBuilder};
pub use error::XlsxToMdError;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        // Placeholder test
        // This test always passes
    }
}
