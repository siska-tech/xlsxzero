//! Error Types Module
//!
//! クレート全体で使用する構造化エラー型を定義するモジュール。
//! `thiserror`を使用して、エラーの自動変換とメッセージフォーマットを実現する。

use thiserror::Error;

/// xlsxzeroクレート全体で使用するエラー型
///
/// このエラー型は、Excelファイルの読み込み、解析、変換処理中に発生する
/// すべてのエラーを統一的に扱うために使用されます。
///
/// # エラーの種類
///
/// - `Io`: I/O操作中に発生したエラー（ファイル読み込み失敗など）
/// - `Parse`: Excelファイルの解析中に発生したエラー（calamine由来）
/// - `Config`: 設定の検証に失敗したエラー（無効な範囲指定など）
/// - `UnsupportedFeature`: サポートされていない機能が検出されたエラー
///
/// # 使用例
///
/// ```rust,no_run
/// use xlsxzero::XlsxToMdError;
/// use std::fs::File;
///
/// fn read_excel_file(path: &str) -> Result<(), XlsxToMdError> {
///     let file = File::open(path)?;  // Ioエラーが自動的に変換される
///     // ... 処理 ...
///     Ok(())
/// }
/// ```
#[derive(Error, Debug)]
pub enum XlsxToMdError {
    /// I/O操作中に発生したエラー
    ///
    /// ファイルの読み込み失敗、書き込み失敗など、標準ライブラリの
    /// `std::io::Error`が発生した場合に使用されます。
    ///
    /// `#[from]`属性により、`std::io::Error`から自動的に変換されます。
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// Excelファイルの解析中に発生したエラー
    ///
    /// calamineクレートがExcelファイルを解析する際に発生したエラーです。
    /// ファイル形式が不正、破損したファイル、サポートされていない形式などが
    /// 原因となります。
    ///
    /// `#[from]`属性により、`calamine::Error`から自動的に変換されます。
    #[error("Failed to parse Excel file: {0}")]
    Parse(#[from] calamine::Error),

    /// UTF-8文字列の変換エラー
    ///
    /// XML解析時にUTF-8文字列への変換に失敗した場合に発生します。
    #[error("UTF-8 conversion error: {0}")]
    Utf8(#[from] std::str::Utf8Error),

    /// ZIPアーカイブの解析エラー
    ///
    /// XLSXファイル（ZIPアーカイブ）の解析中に発生したエラーです。
    #[error("ZIP archive error: {0}")]
    Zip(String),

    /// 数値の解析エラー
    ///
    /// 文字列から数値への変換に失敗した場合に発生します。
    #[error("Number parse error: {0}")]
    ParseInt(#[from] std::num::ParseIntError),

    /// 設定の検証に失敗したエラー
    ///
    /// `ConverterBuilder::build()`時に設定を検証し、無効な設定が検出された
    /// 場合に発生します。例えば、セル範囲の開始座標が終了座標より大きい場合や、
    /// カスタム日付形式が不正な場合などです。
    ///
    /// # 例
    ///
    /// ```rust,no_run
    /// use xlsxzero::{ConverterBuilder, XlsxToMdError};
    ///
    /// let result = ConverterBuilder::new()
    ///     .with_range((10, 0), (0, 0))  // 無効な範囲
    ///     .build();
    ///
    /// match result {
    ///     Err(XlsxToMdError::Config(msg)) => {
    ///         println!("設定エラー: {}", msg);
    ///     }
    ///     _ => {}
    /// }
    /// ```
    #[error("Configuration error: {0}")]
    Config(String),

    /// サポートされていない機能が検出されたエラー
    ///
    /// Phase Iでは実装されていない機能（例: ピボットテーブル、高度な数式など）
    /// が検出された場合に発生します。エラーメッセージには、シート名、セル座標、
    /// 詳細なメッセージが含まれます。
    ///
    /// # 例
    ///
    /// ```rust,no_run
    /// use xlsxzero::XlsxToMdError;
    ///
    /// let error = XlsxToMdError::UnsupportedFeature {
    ///     sheet: "Sheet1".to_string(),
    ///     cell: "A1".to_string(),
    ///     message: "Pivot table is not supported in Phase I".to_string(),
    /// };
    ///
    /// println!("{}", error);
    /// // 出力: "Unsupported feature at sheet 'Sheet1', cell A1: Pivot table is not supported in Phase I"
    /// ```
    #[error("Unsupported feature at sheet '{sheet}', cell {cell}: {message}")]
    UnsupportedFeature {
        /// エラーが発生したシート名
        sheet: String,
        /// エラーが発生したセルの座標（A1記法）
        cell: String,
        /// エラーの詳細メッセージ
        message: String,
    },

    /// セキュリティ制限に違反したエラー
    ///
    /// ZIP bomb攻撃、パストラバーサル攻撃、ファイルサイズ制限などの
    /// セキュリティ制限に違反した場合に発生します。
    ///
    /// # 例
    ///
    /// ```rust,no_run
    /// use xlsxzero::XlsxToMdError;
    ///
    /// let error = XlsxToMdError::SecurityViolation(
    ///     "File size exceeds maximum allowed size".to_string()
    /// );
    /// ```
    #[error("Security violation: {0}")]
    SecurityViolation(String),
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io;

    // Ioエラーのテスト
    #[test]
    fn test_io_error() {
        let io_err = io::Error::new(io::ErrorKind::NotFound, "File not found");
        let error: XlsxToMdError = io_err.into();

        match error {
            XlsxToMdError::Io(e) => {
                assert_eq!(e.kind(), io::ErrorKind::NotFound);
                assert_eq!(e.to_string(), "File not found");
            }
            _ => panic!("Expected Io error"),
        }
    }

    #[test]
    fn test_io_error_display() {
        let io_err = io::Error::new(io::ErrorKind::PermissionDenied, "Permission denied");
        let error: XlsxToMdError = io_err.into();

        let error_msg = error.to_string();
        assert!(error_msg.contains("IO error"));
        assert!(error_msg.contains("Permission denied"));
    }

    // Parseエラーのテスト
    #[test]
    fn test_parse_error() {
        let parse_err = calamine::Error::Msg("Invalid file format");
        let error: XlsxToMdError = parse_err.into();

        match error {
            XlsxToMdError::Parse(e) => match e {
                calamine::Error::Msg(msg) => {
                    assert_eq!(msg, "Invalid file format");
                }
                _ => panic!("Expected Msg variant"),
            },
            _ => panic!("Expected Parse error"),
        }
    }

    #[test]
    fn test_parse_error_display() {
        let parse_err = calamine::Error::Msg("Corrupted file");
        let error: XlsxToMdError = parse_err.into();

        let error_msg = error.to_string();
        assert!(error_msg.contains("Failed to parse Excel file"));
        assert!(error_msg.contains("Corrupted file"));
    }

    // Configエラーのテスト
    #[test]
    fn test_config_error() {
        let error = XlsxToMdError::Config("Invalid range: start > end".to_string());

        match error {
            XlsxToMdError::Config(msg) => {
                assert_eq!(msg, "Invalid range: start > end");
            }
            _ => panic!("Expected Config error"),
        }
    }

    #[test]
    fn test_config_error_display() {
        let error = XlsxToMdError::Config("Invalid date format: 'xyz'".to_string());
        let error_msg = error.to_string();

        assert!(error_msg.contains("Configuration error"));
        assert!(error_msg.contains("Invalid date format: 'xyz'"));
    }

    // UnsupportedFeatureエラーのテスト
    #[test]
    fn test_unsupported_feature_error() {
        let error = XlsxToMdError::UnsupportedFeature {
            sheet: "Sheet1".to_string(),
            cell: "A1".to_string(),
            message: "Pivot table is not supported".to_string(),
        };

        match error {
            XlsxToMdError::UnsupportedFeature {
                sheet,
                cell,
                message,
            } => {
                assert_eq!(sheet, "Sheet1");
                assert_eq!(cell, "A1");
                assert_eq!(message, "Pivot table is not supported");
            }
            _ => panic!("Expected UnsupportedFeature error"),
        }
    }

    #[test]
    fn test_unsupported_feature_error_display() {
        let error = XlsxToMdError::UnsupportedFeature {
            sheet: "MySheet".to_string(),
            cell: "B5".to_string(),
            message: "Complex formula not supported in Phase I".to_string(),
        };

        let error_msg = error.to_string();
        assert!(error_msg.contains("Unsupported feature"));
        assert!(error_msg.contains("MySheet"));
        assert!(error_msg.contains("B5"));
        assert!(error_msg.contains("Complex formula not supported in Phase I"));
    }

    // エラー変換のテスト（?演算子の動作確認）
    #[test]
    fn test_error_conversion_with_question_mark() {
        fn io_operation() -> Result<(), XlsxToMdError> {
            let _file = std::fs::File::open("nonexistent_file.xlsx")?;
            Ok(())
        }

        let result = io_operation();
        assert!(result.is_err());

        match result {
            Err(XlsxToMdError::Io(_)) => {}
            _ => panic!("Expected Io error from ? operator"),
        }
    }

    #[test]
    fn test_error_conversion_from_calamine() {
        // calamine::Errorを直接作成してテスト
        let parse_err = calamine::Error::Msg("File not found");
        let error: XlsxToMdError = parse_err.into();

        match error {
            XlsxToMdError::Parse(_) => {}
            _ => panic!("Expected Parse error"),
        }
    }

    // エラーメッセージのフォーマット確認
    #[test]
    fn test_all_error_formats() {
        // Io
        let io_err: XlsxToMdError = io::Error::other("test io").into();
        assert!(io_err.to_string().starts_with("IO error"));

        // Parse
        let parse_err: XlsxToMdError = calamine::Error::Msg("test parse").into();
        assert!(parse_err
            .to_string()
            .starts_with("Failed to parse Excel file"));

        // Config
        let config_err = XlsxToMdError::Config("test config".to_string());
        assert!(config_err.to_string().starts_with("Configuration error"));

        // UnsupportedFeature
        let unsupported_err = XlsxToMdError::UnsupportedFeature {
            sheet: "Sheet1".to_string(),
            cell: "A1".to_string(),
            message: "test unsupported".to_string(),
        };
        assert!(unsupported_err
            .to_string()
            .starts_with("Unsupported feature"));
    }
}
