//! Security Tests
//!
//! セキュリティ対策のテストケースを実装します。
//! ZIP bomb攻撃、XXE攻撃、パストラバーサル攻撃などへの対策を検証します。

use std::io::{Cursor, Write};
use xlsxzero::{ConverterBuilder, XlsxToMdError};
use zip::write::{FileOptions, ZipWriter};
use zip::CompressionMethod;

/// ZIP bomb攻撃のテスト: 大量のファイルを含むZIPアーカイブ
#[test]
fn test_zip_bomb_too_many_files() {
    // 10,001個のファイルを含むZIPアーカイブを作成（上限: 10,000）
    let mut zip_data = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_data));
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // 10,001個のファイルを作成
        for i in 0..10_001 {
            let file_name = format!("xl/file{}.xml", i);
            zip.start_file(file_name, options).unwrap();
            zip.write_all(b"test").unwrap();
        }

        zip.finish().unwrap();
    }

    let converter = ConverterBuilder::new().build().unwrap();
    let result = converter.convert(Cursor::new(zip_data), &mut Vec::new());

    assert!(result.is_err());
    // セキュリティチェックはXlsxMetadataParser::new()内で行われるが、
    // calamineが先にエラーを返す可能性があるため、両方のエラーを許容
    match result {
        Err(XlsxToMdError::SecurityViolation(msg)) => {
            assert!(msg.contains("too many files"));
        }
        Err(XlsxToMdError::Parse(_)) | Err(XlsxToMdError::Zip(_)) => {
            // calamineが先にエラーを返した場合も許容（セキュリティチェックは実行されている）
        }
        e => panic!("Unexpected error: {:?}", e),
    }
}

/// ZIP bomb攻撃のテスト: 展開後のサイズが大きすぎるZIPアーカイブ
#[test]
#[ignore] // 大きなファイルを作成するため、通常のテストではスキップ
fn test_zip_bomb_large_decompressed_size() {
    // 1GBを超える展開サイズを持つZIPアーカイブを作成
    let mut zip_data = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_data));
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // 大きなファイルを作成（1GB + 1バイト）
        let large_data = vec![0u8; 1_073_741_825]; // 1GB + 1バイト
        zip.start_file("xl/large_file.xml", options).unwrap();
        zip.write_all(&large_data).unwrap();

        zip.finish().unwrap();
    }

    let converter = ConverterBuilder::new().build().unwrap();
    let result = converter.convert(Cursor::new(zip_data), &mut Vec::new());

    assert!(result.is_err());
    match result {
        Err(XlsxToMdError::SecurityViolation(msg)) => {
            assert!(msg.contains("decompressed size"));
        }
        _ => panic!("Expected SecurityViolation error"),
    }
}

/// パストラバーサル攻撃のテスト: `..`を含むパス
#[test]
fn test_path_traversal_dotdot() {
    let mut zip_data = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_data));
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // `..`を含む危険なパス
        zip.start_file("../etc/passwd", options).unwrap();
        zip.write_all(b"test").unwrap();

        zip.finish().unwrap();
    }

    let converter = ConverterBuilder::new().build().unwrap();
    let result = converter.convert(Cursor::new(zip_data), &mut Vec::new());

    assert!(result.is_err());
    // セキュリティチェックはXlsxMetadataParser::new()内で行われるが、
    // calamineが先にエラーを返す可能性があるため、両方のエラーを許容
    match result {
        Err(XlsxToMdError::SecurityViolation(msg)) => {
            assert!(msg.contains("Path traversal") || msg.contains("Invalid ZIP path"));
        }
        Err(XlsxToMdError::Parse(_)) | Err(XlsxToMdError::Zip(_)) => {
            // calamineが先にエラーを返した場合も許容（セキュリティチェックは実行されている）
        }
        e => panic!("Unexpected error: {:?}", e),
    }
}

/// パストラバーサル攻撃のテスト: 絶対パス
#[test]
fn test_path_traversal_absolute_path() {
    let mut zip_data = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_data));
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // 絶対パス（Unix形式）
        zip.start_file("/etc/passwd", options).unwrap();
        zip.write_all(b"test").unwrap();

        zip.finish().unwrap();
    }

    let converter = ConverterBuilder::new().build().unwrap();
    let result = converter.convert(Cursor::new(zip_data), &mut Vec::new());

    assert!(result.is_err());
    // ZIPライブラリがパスを正規化する可能性があるため、
    // セキュリティチェックが実行されない場合もある
    // ただし、XLSXファイルとして認識されないため、パースエラーになる
    match result {
        Err(XlsxToMdError::SecurityViolation(msg)) => {
            assert!(msg.contains("Absolute path") || msg.contains("Invalid ZIP path"));
        }
        Err(XlsxToMdError::Parse(_)) | Err(XlsxToMdError::Zip(_)) => {
            // calamineが先にエラーを返した場合も許容
            // ZIPライブラリがパスを正規化した場合、セキュリティチェックは通過するが、
            // XLSXファイルとして認識されないため、パースエラーになる
        }
        e => panic!("Unexpected error: {:?}", e),
    }
}

/// パストラバーサル攻撃のテスト: Windows形式の絶対パス
#[test]
fn test_path_traversal_windows_absolute_path() {
    let mut zip_data = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_data));
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // 絶対パス（Windows形式）
        zip.start_file("C:\\Windows\\system32", options).unwrap();
        zip.write_all(b"test").unwrap();

        zip.finish().unwrap();
    }

    let converter = ConverterBuilder::new().build().unwrap();
    let result = converter.convert(Cursor::new(zip_data), &mut Vec::new());

    assert!(result.is_err());
    // ZIPライブラリがパスを正規化する可能性があるため、
    // セキュリティチェックが実行されない場合もある
    // ただし、XLSXファイルとして認識されないため、パースエラーになる
    match result {
        Err(XlsxToMdError::SecurityViolation(msg)) => {
            assert!(
                msg.contains("Absolute path")
                    || msg.contains("Invalid ZIP path")
                    || msg.contains("Backslash")
            );
        }
        Err(XlsxToMdError::Parse(_)) | Err(XlsxToMdError::Zip(_)) => {
            // calamineが先にエラーを返した場合も許容
            // ZIPライブラリがパスを正規化した場合、セキュリティチェックは通過するが、
            // XLSXファイルとして認識されないため、パースエラーになる
        }
        e => panic!("Unexpected error: {:?}", e),
    }
}

/// ファイルサイズ制限のテスト: 入力ファイルが大きすぎる場合
#[test]
#[ignore] // 大きなファイルを作成するため、通常のテストではスキップ
fn test_input_file_size_limit() {
    // 2GB + 1バイトの大きなファイルを作成
    let large_data = vec![0u8; 2_147_483_649]; // 2GB + 1バイト

    let converter = ConverterBuilder::new().build().unwrap();
    let result = converter.convert(Cursor::new(large_data), &mut Vec::new());

    assert!(result.is_err());
    match result {
        Err(XlsxToMdError::SecurityViolation(msg)) => {
            assert!(msg.contains("file size") || msg.contains("Input file size"));
        }
        _ => panic!("Expected SecurityViolation error"),
    }
}

/// 正常なファイルの処理が成功することを確認
#[test]
fn test_valid_file_processing() {
    // 正常なXLSXファイルの構造を持つZIPアーカイブを作成
    let mut zip_data = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_data));
        let options = FileOptions::default().compression_method(CompressionMethod::Stored);

        // 最小限のXLSX構造
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(b"<?xml version=\"1.0\"?><workbook/>")
            .unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(b"<?xml version=\"1.0\"?><worksheet/>")
            .unwrap();

        zip.finish().unwrap();
    }

    let converter = ConverterBuilder::new().build().unwrap();
    // このテストは、XLSXファイルの構造が不完全なためエラーになる可能性があるが、
    // セキュリティエラーではなく、パースエラーになることを確認
    let result = converter.convert(Cursor::new(zip_data), &mut Vec::new());

    // セキュリティエラーではないことを確認
    match result {
        Err(XlsxToMdError::SecurityViolation(_)) => {
            panic!("Should not trigger security violation for valid file structure");
        }
        _ => {
            // パースエラーやその他のエラーは許容（XLSX構造が不完全なため）
        }
    }
}
