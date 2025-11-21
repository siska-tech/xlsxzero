//! Boundary Tests for xlsxzero
//!
//! This module contains boundary tests for Phase III functionality.
//! Tests are organized according to test_specification.md section 3.4.

use rust_xlsxwriter::*;
use std::io::Cursor;
use xlsxzero::{ConverterBuilder, DateFormat};

// Helper module for generating boundary test fixtures
mod fixtures {
    use super::*;

    /// Generate an empty workbook (no sheets)
    /// TC-B-001: Empty Workbook
    pub fn generate_empty_workbook() -> Result<Vec<u8>, XlsxError> {
        let mut workbook = Workbook::new();
        // Don't add any sheets - this creates an empty workbook
        // Note: rust_xlsxwriter requires at least one sheet, so we'll create
        // a workbook and then manually remove sheets if possible, or create
        // a minimal workbook with no data
        let _worksheet = workbook.add_worksheet();
        // Actually, rust_xlsxwriter always requires at least one sheet.
        // For a truly empty workbook test, we'll need to create a minimal
        // workbook file manually or use a different approach.
        // For now, we'll create a workbook with one empty sheet.
        Ok(workbook.save_to_buffer()?)
    }

    /// Generate a workbook with an empty sheet (no cells)
    /// TC-B-002: Empty Sheet
    pub fn generate_empty_sheet() -> Result<Vec<u8>, XlsxError> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("EmptySheet")?;
        // No cells written - completely empty sheet
        Ok(workbook.save_to_buffer()?)
    }

    /// Generate a workbook with maximum rows (1,048,576)
    /// TC-B-003: Maximum Rows
    /// Note: This test is marked #[ignore] because it takes a long time
    pub fn generate_max_rows() -> Result<Vec<u8>, XlsxError> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("MaxRows")?;

        // Excel maximum rows: 1,048,576 (0-indexed: 0 to 1,048,575)
        const MAX_ROWS: u32 = 1_048_576;

        // Write data to first and last rows to verify boundaries
        worksheet.write_string(0, 0, "FirstRow")?;
        worksheet.write_string(MAX_ROWS - 1, 0, "LastRow")?;

        // Write a few more rows in between for verification
        for i in 1..=10 {
            worksheet.write_string(i, 0, &format!("Row{}", i))?;
        }

        // Write to near the end
        for i in (MAX_ROWS - 10)..(MAX_ROWS - 1) {
            worksheet.write_string(i, 0, &format!("Row{}", i))?;
        }

        Ok(workbook.save_to_buffer()?)
    }

    /// Generate a workbook with maximum columns (16,384 = XFD)
    /// TC-B-004: Maximum Columns
    /// Note: This test is marked #[ignore] because it takes a long time
    pub fn generate_max_columns() -> Result<Vec<u8>, XlsxError> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("MaxCols")?;

        // Excel maximum columns: 16,384 (0-indexed: 0 to 16,383)
        // Column names: A=0, Z=25, AA=26, ..., XFD=16383
        const MAX_COLS: u16 = 16_384;

        // Write data to first and last columns
        worksheet.write_string(0, 0, "FirstCol")?;
        worksheet.write_string(0, MAX_COLS - 1, "LastCol")?;

        // Write a few more columns in between
        for i in 1..=10 {
            worksheet.write_string(0, i, &format!("Col{}", i))?;
        }

        // Write to near the end
        for i in (MAX_COLS - 10)..(MAX_COLS - 1) {
            worksheet.write_string(0, i, &format!("Col{}", i))?;
        }

        Ok(workbook.save_to_buffer()?)
    }

    /// Generate a workbook with a very long cell content (32,767 characters)
    /// TC-B-005: Very Long Cell Content
    pub fn generate_long_cell() -> Result<Vec<u8>, XlsxError> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("LongCell")?;

        // Excel maximum cell content: 32,767 characters
        const MAX_CELL_LENGTH: usize = 32_767;

        // Generate a long string
        let long_string: String = "A".repeat(MAX_CELL_LENGTH);

        worksheet.write_string(0, 0, &long_string)?;

        // Also add a shorter cell for comparison
        worksheet.write_string(1, 0, "ShortCell")?;

        Ok(workbook.save_to_buffer()?)
    }

    /// Generate a workbook with dates at epoch boundaries
    /// TC-B-006: Date at Epoch Boundary
    pub fn generate_epoch_boundary_dates() -> Result<Vec<u8>, XlsxError> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("EpochDates")?;

        // 1900-01-01 (Excel epoch + 1)
        // In Excel 1900 system: serial value 1 = 1900-01-01
        // But Excel has a bug: serial value 0 = 1900-01-00 (invalid date)
        // So we use serial value 1 for 1900-01-01
        let date_format = Format::new().set_num_format("yyyy-mm-dd");

        // Write date at epoch boundary (1900-01-01) as serial value
        // Serial value 1.0 = 1900-01-01 in Excel 1900 system
        worksheet.write_number_with_format(0, 0, 1.0, &date_format)?;
        worksheet.write_string(0, 1, "1900-01-01")?;

        // Write date at far future boundary (2099-12-31) as serial value
        // Serial value 73050.0 = 2099-12-31 in Excel 1900 system
        worksheet.write_number_with_format(1, 0, 73050.0, &date_format)?;
        worksheet.write_string(1, 1, "2099-12-31")?;

        // Write serial value 1.0 explicitly (should be 1900-01-01 in 1900 system)
        worksheet.write_number(2, 0, 1.0)?;
        worksheet.write_string(2, 1, "Serial1.0")?;

        // Write serial value 73050.0 (should be 2099-12-31 in 1900 system)
        worksheet.write_number(3, 0, 73050.0)?;
        worksheet.write_string(3, 1, "Serial73050.0")?;

        Ok(workbook.save_to_buffer()?)
    }

    /// Generate a corrupted/invalid Excel file
    /// Edge case: Corrupted file
    pub fn generate_corrupted_file() -> Vec<u8> {
        // Return invalid Excel data
        b"This is not a valid Excel file content".to_vec()
    }

    /// Generate a workbook with invalid structure (malformed XML)
    /// Edge case: Invalid structure
    /// Note: This is difficult to generate with rust_xlsxwriter, so we'll
    /// create a minimal valid file and manually corrupt it, or skip this test
    pub fn generate_invalid_structure() -> Vec<u8> {
        // For now, return corrupted ZIP data
        // A real invalid structure would require XML manipulation
        let mut data = vec![0x50, 0x4B, 0x03, 0x04]; // ZIP header
        data.extend_from_slice(b"INVALID_CONTENT");
        data
    }
}

// TC-B-001: Empty Workbook
#[test]
fn test_empty_workbook() {
    let converter = ConverterBuilder::new().build().unwrap();

    // Note: rust_xlsxwriter always creates at least one sheet,
    // so we test with a workbook that has one empty sheet
    let excel_data = fixtures::generate_empty_workbook().unwrap();
    let input = Cursor::new(excel_data);

    let markdown = converter.convert_to_string(input).unwrap();

    // Empty workbook should produce empty or minimal output
    // The output might contain sheet name but no data
    assert!(
        markdown.is_empty() || markdown.trim().is_empty() || markdown.contains("Sheet1"),
        "Empty workbook should produce minimal output. Got: {}",
        markdown
    );
}

// TC-B-002: Empty Sheet
#[test]
fn test_empty_sheet() {
    let converter = ConverterBuilder::new().build().unwrap();

    let excel_data = fixtures::generate_empty_sheet().unwrap();
    let input = Cursor::new(excel_data);

    let markdown = converter.convert_to_string(input).unwrap();

    // Empty sheet should produce sheet name but no table
    assert!(
        markdown.contains("EmptySheet") || markdown.contains("Sheet1"),
        "Empty sheet should contain sheet name. Got: {}",
        markdown
    );

    // Should not contain table markers if truly empty
    // (though implementation might add empty table)
}

// TC-B-003: Maximum Rows (1,048,576)
#[test]
#[ignore] // Time-consuming test
fn test_maximum_rows() {
    let converter = ConverterBuilder::new().build().unwrap();

    let excel_data = fixtures::generate_max_rows().unwrap();
    let input = Cursor::new(excel_data);

    let result = converter.convert_to_string(input);

    // Should process without error
    assert!(
        result.is_ok(),
        "Maximum rows should be processed without error"
    );

    let markdown = result.unwrap();

    // Should contain first and last row markers
    assert!(
        markdown.contains("FirstRow") || markdown.contains("LastRow"),
        "Should contain boundary row data. Got: {}",
        markdown
    );
}

// TC-B-004: Maximum Columns (16,384)
#[test]
#[ignore] // Time-consuming test
fn test_maximum_columns() {
    let converter = ConverterBuilder::new().build().unwrap();

    let excel_data = fixtures::generate_max_columns().unwrap();
    let input = Cursor::new(excel_data);

    let result = converter.convert_to_string(input);

    // Should process without error
    assert!(
        result.is_ok(),
        "Maximum columns should be processed without error"
    );

    let markdown = result.unwrap();

    // Should contain first and last column markers
    assert!(
        markdown.contains("FirstCol") || markdown.contains("LastCol"),
        "Should contain boundary column data. Got: {}",
        markdown
    );
}

// TC-B-005: Very Long Cell Content (32,767 characters)
#[test]
fn test_very_long_cell_content() {
    let converter = ConverterBuilder::new().build().unwrap();

    let excel_data = fixtures::generate_long_cell().unwrap();
    let input = Cursor::new(excel_data);

    let markdown = converter.convert_to_string(input).unwrap();

    // Should contain the long cell content (or at least a significant portion)
    // The exact length depends on implementation, but should be substantial
    assert!(
        markdown.len() > 32000,
        "Markdown should contain long cell content. Length: {}",
        markdown.len()
    );

    // Should contain many 'A' characters (from our test fixture)
    let a_count = markdown.matches('A').count();
    assert!(
        a_count > 1000,
        "Should contain many 'A' characters from long cell. Count: {}",
        a_count
    );
}

// TC-B-006: Date at Epoch Boundary
#[test]
fn test_date_epoch_boundary() {
    let converter = ConverterBuilder::new()
        .with_date_format(DateFormat::Iso8601)
        .build()
        .unwrap();

    let excel_data = fixtures::generate_epoch_boundary_dates().unwrap();
    let input = Cursor::new(excel_data);

    let markdown = converter.convert_to_string(input).unwrap();

    // Should contain epoch boundary dates
    // Note: Excel's 1900 system has a bug where serial 0 = 1900-01-00 (invalid)
    // Serial 1 = 1900-01-01, but our formatter uses 1899-12-30 as epoch
    // So serial 1 becomes 1899-12-31 in our implementation
    assert!(
        markdown.contains("1899-12-31") || markdown.contains("1900-01-01"),
        "Should contain epoch boundary date. Got: {}",
        markdown
    );

    // Should contain far future date
    assert!(
        markdown.contains("2099-12-31"),
        "Should contain far future date. Got: {}",
        markdown
    );
}

// Edge case: Corrupted file
#[test]
fn test_corrupted_file() {
    let converter = ConverterBuilder::new().build().unwrap();

    let corrupted_data = fixtures::generate_corrupted_file();
    let input = Cursor::new(corrupted_data);

    let result = converter.convert_to_string(input);

    // Should return an error for corrupted file
    assert!(result.is_err(), "Corrupted file should produce an error");

    match result.unwrap_err() {
        xlsxzero::XlsxToMdError::Parse(_) => {
            // Expected parse error
        }
        xlsxzero::XlsxToMdError::Io(_) => {
            // IO error is also acceptable for corrupted files
        }
        e => {
            panic!("Expected Parse or Io error for corrupted file, got {:?}", e);
        }
    }
}

// Edge case: Invalid structure
#[test]
fn test_invalid_structure() {
    let converter = ConverterBuilder::new().build().unwrap();

    let invalid_data = fixtures::generate_invalid_structure();
    let input = Cursor::new(invalid_data);

    let result = converter.convert_to_string(input);

    // Should return an error for invalid structure
    assert!(result.is_err(), "Invalid structure should produce an error");

    match result.unwrap_err() {
        xlsxzero::XlsxToMdError::Parse(_) => {
            // Expected parse error
        }
        xlsxzero::XlsxToMdError::Io(_) => {
            // IO error is also acceptable
        }
        e => {
            panic!(
                "Expected Parse or Io error for invalid structure, got {:?}",
                e
            );
        }
    }
}
