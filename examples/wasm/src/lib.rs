//! xlsxzero WASM Example
//!
//! This module provides WebAssembly bindings for xlsxzero,
//! allowing Excel files to be converted to Markdown in the browser.

use wasm_bindgen::prelude::*;
use xlsxzero::ConverterBuilder;
use std::io::Cursor;

/// Initialize panic hook for better error messages in the browser
#[wasm_bindgen(start)]
pub fn init() {
    #[cfg(feature = "console_error_panic_hook")]
    console_error_panic_hook::set_once();
}

/// Convert Excel file (as bytes) to Markdown string
///
/// # Arguments
/// * `excel_bytes` - Excel file content as a Uint8Array from JavaScript
///
/// # Returns
/// * Success: Markdown string
/// * Error: Error message string
#[wasm_bindgen]
pub fn convert_excel_to_markdown(excel_bytes: &[u8]) -> Result<String, String> {
    // Create converter with default settings
    let converter = ConverterBuilder::new()
        .build()
        .map_err(|e| format!("Failed to create converter: {}", e))?;

    // Convert from byte slice to Cursor
    let cursor = Cursor::new(excel_bytes);

    // Convert to Markdown string
    let markdown = converter
        .convert_to_string(cursor)
        .map_err(|e| format!("Conversion error: {}", e))?;

    Ok(markdown)
}

/// Convert Excel file with custom options
///
/// # Arguments
/// * `excel_bytes` - Excel file content as a Uint8Array from JavaScript
/// * `sheet_index` - Optional sheet index (0-based), null for all sheets
/// * `merge_strategy` - Merge strategy: "data_duplication" or "html_fallback"
/// * `date_format` - Date format: "iso8601" or custom format string
///
/// # Returns
/// * Success: Markdown string
/// * Error: Error message string
#[wasm_bindgen]
pub fn convert_excel_to_markdown_custom(
    excel_bytes: &[u8],
    sheet_index: Option<usize>,
    merge_strategy: Option<String>,
    date_format: Option<String>,
) -> Result<String, String> {
    // Build converter with custom options
    let mut builder = ConverterBuilder::new();

    // Set sheet selector
    if let Some(idx) = sheet_index {
        builder = builder.with_sheet_selector(xlsxzero::SheetSelector::Index(idx));
    }

    // Set merge strategy
    if let Some(ref strategy) = merge_strategy {
        match strategy.as_str() {
            "data_duplication" => {
                builder = builder.with_merge_strategy(xlsxzero::MergeStrategy::DataDuplication)
            }
            "html_fallback" => {
                builder = builder.with_merge_strategy(xlsxzero::MergeStrategy::HtmlFallback)
            }
            _ => return Err(format!("Invalid merge strategy: {}", strategy)),
        }
    }

    // Set date format
    if let Some(ref format) = date_format {
        if format == "iso8601" {
            builder = builder.with_date_format(xlsxzero::DateFormat::Iso8601);
        } else {
            builder = builder.with_date_format(xlsxzero::DateFormat::Custom(format.clone()));
        }
    }

    let converter = builder.build().map_err(|e| format!("Failed to create converter: {}", e))?;

    // Convert from byte slice to Cursor
    let cursor = Cursor::new(excel_bytes);

    // Convert to Markdown string
    let markdown = converter
        .convert_to_string(cursor)
        .map_err(|e| format!("Conversion error: {}", e))?;

    Ok(markdown)
}

/// Get version information
#[wasm_bindgen]
pub fn get_version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

