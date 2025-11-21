//! Formatter Module
//!
//! セル値のフォーマット処理を提供するモジュール。
//! Phase Iでは簡易実装として、Number Format Stringなしで動作します。

use chrono::{Duration, NaiveDate};

use crate::api::DateFormat;
use crate::builder::ConversionConfig;
use crate::error::XlsxToMdError;
use crate::types::{CellValue, RawCellData};

/// セルフォーマッター
///
/// セル値のフォーマット処理のファサードとして機能します。
#[derive(Debug)]
pub(crate) struct CellFormatter {
    /// 日付フォーマッター
    date_formatter: DateFormatter,

    /// 数値フォーマッター
    number_formatter: NumberFormatter,
}

impl CellFormatter {
    /// 新しいCellFormatterインスタンスを生成
    pub fn new() -> Self {
        Self {
            date_formatter: DateFormatter,
            number_formatter: NumberFormatter,
        }
    }

    /// セル値をフォーマット
    ///
    /// # 引数
    ///
    /// * `raw_cell` - パーサーから抽出された生のセルデータ
    /// * `config` - 変換設定
    /// * `is_1904` - 1904年エポックを使用するかどうか（Phase II）
    ///
    /// # 戻り値
    ///
    /// * `Ok(String)` - フォーマット済み文字列
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    pub fn format_cell(
        &self,
        raw_cell: &RawCellData,
        config: &ConversionConfig,
        is_1904: bool,
    ) -> Result<String, XlsxToMdError> {
        use crate::api::FormulaMode;

        // 1. 数式モードの処理
        if config.formula_mode == FormulaMode::Formula {
            if let Some(ref formula) = raw_cell.formula {
                return Ok(formula.clone());
            }
        }

        // 2. 値の種類に応じてフォーマット
        let formatted_value = match &raw_cell.value {
            CellValue::Number(n) => {
                // 日付判定
                if self.is_date_value(*n, &raw_cell.format_id, &raw_cell.format_string) {
                    self.date_formatter.format(*n, config, is_1904)?
                } else {
                    self.number_formatter.format(*n, &raw_cell.format_string)?
                }
            }

            CellValue::String(s) => {
                // リッチテキストが存在する場合は、リッチテキストを使用
                if let Some(ref rich_text_segments) = raw_cell.rich_text {
                    self.format_rich_text(rich_text_segments)
                } else {
                    self.escape_markdown(s)
                }
            }

            CellValue::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),

            CellValue::Error(e) => e.clone(),

            CellValue::Empty => String::new(),
        };

        // 3. ハイパーリンクがある場合はMarkdownリンク構文に変換
        if let Some(ref url) = raw_cell.hyperlink {
            // 表示テキストが空の場合はURLを使用
            let display_text = if formatted_value.is_empty() {
                url.clone()
            } else {
                formatted_value
            };
            Ok(format!("[{}]({})", display_text, url))
        } else {
            Ok(formatted_value)
        }
    }

    /// 数値が日付値かどうかを判定（ヒューリスティック）
    ///
    /// # 引数
    ///
    /// * `value` - 数値
    /// * `format_id` - 数値書式ID（Phase Iでは常にNone）
    /// * `format_string` - カスタム書式文字列（Phase Iでは常にNone）
    ///
    /// # 戻り値
    ///
    /// 日付値と判定された場合は`true`
    fn is_date_value(
        &self,
        _value: f64,
        format_id: &Option<u16>,
        format_string: &Option<String>,
    ) -> bool {
        // 1. 組み込み日付書式IDのチェック
        if let Some(id) = format_id {
            // Excel組み込み日付書式ID
            // 14: "m/d/yy"
            // 15-17: 日付形式
            // 18-21: 時刻形式
            // 22: "m/d/yy h:mm"
            // 45-47: 追加の日付形式
            if matches!(id, 14..=22 | 45..=47) {
                return true;
            }
        }

        // 2. カスタム書式文字列のチェック
        if let Some(ref format_str) = format_string {
            let format_lower = format_str.to_lowercase();
            if format_lower.contains("yy")
                || format_lower.contains("mm")
                || format_lower.contains("dd")
                || format_lower.contains("hh")
            {
                return true;
            }
        }

        // 3. 値の範囲チェック（ヒューリスティック）
        // Phase Iでは、format_idやformat_stringがない場合は、日付として判定しない
        // （誤判定を避けるため）
        // Phase IIでNumber Format Stringが取得できるようになったら、より正確な判定が可能
        false
    }

    /// Markdown特殊文字をエスケープ
    ///
    /// # 引数
    ///
    /// * `s` - エスケープする文字列
    ///
    /// # 戻り値
    ///
    /// エスケープ済み文字列
    fn escape_markdown(&self, s: &str) -> String {
        s.replace('\\', "\\\\")
            .replace('|', "\\|")
            .replace('\n', "<br>")
    }

    /// リッチテキストをMarkdown形式に変換
    ///
    /// # 引数
    ///
    /// * `segments` - リッチテキストセグメントのリスト
    ///
    /// # 戻り値
    ///
    /// Markdown形式の文字列（太字: `**text**`, 斜体: `*text*`）
    fn format_rich_text(&self, segments: &[crate::types::RichTextSegment]) -> String {
        let mut result = String::new();
        for segment in segments {
            let mut text = self.escape_markdown(&segment.text);

            // 書式を適用（太字、斜体）
            if segment.format.bold && segment.format.italic {
                text = format!("***{}***", text);
            } else if segment.format.bold {
                text = format!("**{}**", text);
            } else if segment.format.italic {
                text = format!("*{}*", text);
            }

            result.push_str(&text);
        }
        result
    }
}

impl Default for CellFormatter {
    fn default() -> Self {
        Self::new()
    }
}

/// 日付フォーマッター
///
/// Excelのシリアル日付値を文字列に変換します。
/// Phase Iでは常に1900年エポックとして処理します。
#[derive(Debug)]
pub(crate) struct DateFormatter;

impl DateFormatter {
    /// 日付値をフォーマット
    ///
    /// # 引数
    ///
    /// * `serial_value` - Excelのシリアル日付値
    /// * `config` - 変換設定
    /// * `is_1904` - 1904年エポックを使用するかどうか
    ///
    /// # 戻り値
    ///
    /// * `Ok(String)` - フォーマット済み日付文字列
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    ///
    /// # エポックシステム
    ///
    /// - 1900年システム（デフォルト）: 1899年12月30日起算
    ///   - Excelの1900年うるう年バグを考慮
    ///   - シリアル値0 = 1900年1月0日（存在しない日、Excelのバグ）
    ///   - シリアル値1 = 1900年1月1日
    /// - 1904年システム: 1904年1月1日起算
    ///   - Mac版Excelで使用される
    ///   - シリアル値0 = 1904年1月1日
    ///   - シリアル値1 = 1904年1月2日
    pub fn format(
        &self,
        serial_value: f64,
        config: &ConversionConfig,
        is_1904: bool,
    ) -> Result<String, XlsxToMdError> {
        let (epoch, days_offset) = if is_1904 {
            // 1904年システム: 1904年1月1日起算
            // シリアル値0 = 1904-01-01
            let epoch = NaiveDate::from_ymd_opt(1904, 1, 1)
                .ok_or_else(|| XlsxToMdError::Config("Invalid epoch date".to_string()))?;
            (epoch, 0i64)
        } else {
            // 1900年システム: 1899年12月30日起算
            // Excelの1900年うるう年バグを考慮
            // シリアル値0 = 1900年1月0日（存在しない日、Excelのバグ）
            // シリアル値1 = 1900年1月1日
            // エポック1899-12-30から、シリアル値1で1900-01-01になるように調整
            let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)
                .ok_or_else(|| XlsxToMdError::Config("Invalid epoch date".to_string()))?;
            // シリアル値に対して +1日を加算する必要がある
            (epoch, 1i64)
        };

        // シリアル値からNaiveDateに変換
        let days = serial_value.floor() as i64;
        let date = epoch
            .checked_add_signed(Duration::days(days + days_offset))
            .ok_or_else(|| {
                XlsxToMdError::Config(format!(
                    "Date calculation overflow: serial_value={}, is_1904={}",
                    serial_value, is_1904
                ))
            })?;

        // DateFormatに応じてフォーマット
        let formatted = match &config.date_format {
            DateFormat::Iso8601 => date.format("%Y-%m-%d").to_string(),
            DateFormat::Custom(format_str) => date.format(format_str).to_string(),
        };

        Ok(formatted)
    }
}

/// 数値フォーマッター
///
/// 数値を文字列に変換します。
/// Phase IIでNumber Format Parserを使用してNumber Format Stringを解析・適用します。
#[derive(Debug)]
pub(crate) struct NumberFormatter;

impl NumberFormatter {
    /// 数値をフォーマット
    ///
    /// # 引数
    ///
    /// * `value` - 数値
    /// * `format_string` - カスタム書式文字列（Phase IIで取得可能）
    ///
    /// # 戻り値
    ///
    /// * `Ok(String)` - フォーマット済み数値文字列
    ///
    /// # Phase II実装
    ///
    /// FormatParserを使用してNumber Format Stringを解析・適用します。
    /// format_stringがNoneの場合は`to_string()`でフォールバックします。
    pub fn format(
        &self,
        value: f64,
        format_string: &Option<String>,
    ) -> Result<String, XlsxToMdError> {
        if let Some(ref format_str) = format_string {
            // Number Format Parser を使用
            match crate::format::FormatParser::new(format_str) {
                Ok(parser) => {
                    match parser.format_number(value) {
                        Ok(formatted) => Ok(formatted),
                        Err(_) => {
                            // パースエラーまたはフォーマットエラーの場合はフォールバック
                            Ok(value.to_string())
                        }
                    }
                }
                Err(_) => {
                    // パース失敗の場合はフォールバック
                    Ok(value.to_string())
                }
            }
        } else {
            // format_stringがNoneの場合はフォールバック
            Ok(value.to_string())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::api::{DateFormat, FormulaMode};
    use crate::types::{CellCoord, CellValue, RawCellData};

    fn create_test_config() -> ConversionConfig {
        ConversionConfig::default()
    }

    fn create_test_config_with_date_format(date_format: DateFormat) -> ConversionConfig {
        ConversionConfig {
            date_format,
            ..Default::default()
        }
    }

    fn create_test_config_with_formula_mode(formula_mode: FormulaMode) -> ConversionConfig {
        ConversionConfig {
            formula_mode,
            ..Default::default()
        }
    }

    #[test]
    fn test_escape_markdown() {
        let formatter = CellFormatter::new();
        assert_eq!(formatter.escape_markdown("test"), "test");
        assert_eq!(formatter.escape_markdown("test|value"), "test\\|value");
        assert_eq!(formatter.escape_markdown("test\nvalue"), "test<br>value");
        assert_eq!(formatter.escape_markdown("test\\value"), "test\\\\value");
        assert_eq!(
            formatter.escape_markdown("test|value\nwith\\backslash"),
            "test\\|value<br>with\\\\backslash"
        );
    }

    #[test]
    fn test_is_date_value_by_format_id() {
        let formatter = CellFormatter::new();
        // 組み込み日付書式IDのチェック
        assert!(formatter.is_date_value(1.0, &Some(14), &None));
        assert!(formatter.is_date_value(1.0, &Some(22), &None));
        assert!(formatter.is_date_value(1.0, &Some(45), &None));
        assert!(!formatter.is_date_value(1.0, &Some(1), &None)); // 通常の数値書式
    }

    #[test]
    fn test_is_date_value_by_format_string() {
        let formatter = CellFormatter::new();
        // カスタム書式文字列のチェック
        assert!(formatter.is_date_value(1.0, &None, &Some("yyyy-mm-dd".to_string())));
        assert!(formatter.is_date_value(1.0, &None, &Some("MM/DD/YY".to_string())));
        assert!(formatter.is_date_value(1.0, &None, &Some("hh:mm:ss".to_string())));
        assert!(!formatter.is_date_value(1.0, &None, &Some("#,##0".to_string())));
        // 数値書式
    }

    #[test]
    fn test_is_date_value_by_range() {
        let formatter = CellFormatter::new();
        // Phase Iでは、format_idやformat_stringがない場合は日付として判定しない
        // （誤判定を避けるため）
        assert!(!formatter.is_date_value(1.0, &None, &None));
        assert!(!formatter.is_date_value(100.0, &None, &None));
        assert!(!formatter.is_date_value(10000.0, &None, &None));
        assert!(!formatter.is_date_value(0.0, &None, &None));
        assert!(!formatter.is_date_value(70000.0, &None, &None));
        assert!(!formatter.is_date_value(-1.0, &None, &None));
    }

    #[test]
    fn test_date_formatter_iso8601() {
        let formatter = DateFormatter;
        let config = create_test_config_with_date_format(DateFormat::Iso8601);

        // 1900年1月1日（シリアル値: 1）
        let result = formatter.format(1.0, &config, false).unwrap();
        assert_eq!(result, "1900-01-01");

        // 1900年1月2日（シリアル値: 2）
        let result = formatter.format(2.0, &config, false).unwrap();
        assert_eq!(result, "1900-01-02");

        // 2025年1月2日（シリアル値: 45658）
        // エポック1899-12-30 + (45658+1)日 = 2025-01-02
        let result = formatter.format(45658.0, &config, false).unwrap();
        assert_eq!(result, "2025-01-02");
    }

    #[test]
    fn test_date_formatter_custom() {
        let formatter = DateFormatter;
        let config =
            create_test_config_with_date_format(DateFormat::Custom("%Y/%m/%d".to_string()));

        // 1900年1月1日（シリアル値: 1）
        let result = formatter.format(1.0, &config, false).unwrap();
        assert_eq!(result, "1900/01/01");
    }

    #[test]
    fn test_date_formatter_1904_epoch() {
        let formatter = DateFormatter;
        let config = create_test_config_with_date_format(DateFormat::Iso8601);

        // 1904年エポックシステムのテスト
        // シリアル値0 = 1904年1月1日
        let result = formatter.format(0.0, &config, true).unwrap();
        assert_eq!(result, "1904-01-01");

        // シリアル値1 = 1904年1月2日
        let result = formatter.format(1.0, &config, true).unwrap();
        assert_eq!(result, "1904-01-02");

        // シリアル値365 = 1904年12月31日（1904年はうるう年で366日）
        let result = formatter.format(365.0, &config, true).unwrap();
        assert_eq!(result, "1904-12-31");

        // シリアル値366 = 1905年1月1日
        let result = formatter.format(366.0, &config, true).unwrap();
        assert_eq!(result, "1905-01-01");

        // 1900年エポックシステムとの比較
        // シリアル値0（1900年システム） = 1899年12月31日（1900年1月0日に対応、Excelのバグを考慮）
        let result_1900 = formatter.format(0.0, &config, false).unwrap();
        assert_eq!(result_1900, "1899-12-31");

        // シリアル値1（1900年システム） = 1900年1月1日
        let result_1900 = formatter.format(1.0, &config, false).unwrap();
        assert_eq!(result_1900, "1900-01-01");
    }

    #[test]
    fn test_number_formatter() {
        let formatter = NumberFormatter;
        // Phase I: to_string()でフォールバック
        assert_eq!(formatter.format(123.45, &None).unwrap(), "123.45");
        assert_eq!(formatter.format(0.0, &None).unwrap(), "0");
        assert_eq!(formatter.format(-123.45, &None).unwrap(), "-123.45");
    }

    #[test]
    fn test_format_cell_number() {
        let formatter = CellFormatter::new();
        let config = create_test_config();

        let raw_cell = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Number(123.45),
            format_id: None,
            format_string: None,
            formula: None,
            hyperlink: None,
            rich_text: None,
        };

        let result = formatter.format_cell(&raw_cell, &config, false).unwrap();
        assert_eq!(result, "123.45");
    }

    #[test]
    fn test_format_cell_string() {
        let formatter = CellFormatter::new();
        let config = create_test_config();

        let raw_cell = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::String("test|value".to_string()),
            format_id: None,
            format_string: None,
            formula: None,
            hyperlink: None,
            rich_text: None,
        };

        let result = formatter.format_cell(&raw_cell, &config, false).unwrap();
        assert_eq!(result, "test\\|value");
    }

    #[test]
    fn test_format_cell_bool() {
        let formatter = CellFormatter::new();
        let config = create_test_config();

        let raw_cell_true = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Bool(true),
            format_id: None,
            format_string: None,
            formula: None,
            hyperlink: None,
            rich_text: None,
        };

        let raw_cell_false = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Bool(false),
            format_id: None,
            format_string: None,
            formula: None,
            hyperlink: None,
            rich_text: None,
        };

        assert_eq!(
            formatter
                .format_cell(&raw_cell_true, &config, false)
                .unwrap(),
            "TRUE"
        );
        assert_eq!(
            formatter
                .format_cell(&raw_cell_false, &config, false)
                .unwrap(),
            "FALSE"
        );
    }

    #[test]
    fn test_format_cell_error() {
        let formatter = CellFormatter::new();
        let config = create_test_config();

        let raw_cell = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Error("#DIV/0!".to_string()),
            format_id: None,
            format_string: None,
            formula: None,
            hyperlink: None,
            rich_text: None,
        };

        let result = formatter.format_cell(&raw_cell, &config, false).unwrap();
        assert_eq!(result, "#DIV/0!");
    }

    #[test]
    fn test_format_cell_empty() {
        let formatter = CellFormatter::new();
        let config = create_test_config();

        let raw_cell = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Empty,
            format_id: None,
            format_string: None,
            formula: None,
            hyperlink: None,
            rich_text: None,
        };

        let result = formatter.format_cell(&raw_cell, &config, false).unwrap();
        assert_eq!(result, "");
    }

    #[test]
    fn test_format_cell_date() {
        let formatter = CellFormatter::new();
        let config = create_test_config_with_date_format(DateFormat::Iso8601);

        // 日付値として判定される数値（format_idで日付書式を指定）
        let raw_cell = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Number(1.0),
            format_id: Some(14), // 日付書式ID
            format_string: None,
            formula: None,
            hyperlink: None,
            rich_text: None,
        };

        let result = formatter.format_cell(&raw_cell, &config, false).unwrap();
        assert_eq!(result, "1900-01-01");
    }

    #[test]
    fn test_format_cell_formula_mode_cached() {
        let formatter = CellFormatter::new();
        let config = create_test_config_with_formula_mode(FormulaMode::CachedValue);

        let raw_cell = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Number(100.0),
            format_id: None,
            format_string: None,
            formula: Some("=SUM(A1:A10)".to_string()),
            hyperlink: None,
            rich_text: None,
        };

        // CachedValueモードでは数式を無視して値をフォーマット
        let result = formatter.format_cell(&raw_cell, &config, false).unwrap();
        assert_eq!(result, "100");
    }

    #[test]
    fn test_format_cell_formula_mode_formula() {
        let formatter = CellFormatter::new();
        let config = create_test_config_with_formula_mode(FormulaMode::Formula);

        let raw_cell = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Number(100.0),
            format_id: None,
            format_string: None,
            formula: Some("=SUM(A1:A10)".to_string()),
            hyperlink: None,
            rich_text: None,
        };

        // Formulaモードでは数式をそのまま返す
        let result = formatter.format_cell(&raw_cell, &config, false).unwrap();
        assert_eq!(result, "=SUM(A1:A10)");
    }

    #[test]
    fn test_format_cell_formula_mode_formula_no_formula() {
        let formatter = CellFormatter::new();
        let config = create_test_config_with_formula_mode(FormulaMode::Formula);

        let raw_cell = RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::Number(100.0),
            format_id: None,
            format_string: None,
            formula: None, // 数式がない場合
            hyperlink: None,
            rich_text: None,
        };

        // 数式がない場合は通常の値としてフォーマット
        let result = formatter.format_cell(&raw_cell, &config, false).unwrap();
        assert_eq!(result, "100");
    }

    // プロパティベーステスト: TC-PBT-002
    #[allow(unused_doc_comments)]
    mod property_tests {
        use super::*;
        use proptest::prelude::*;

        #[allow(unused_doc_comments)]
        /// TC-PBT-002: Date Conversion Monotonicity
        ///
        /// 日付変換の単調性を検証します。
        /// シリアル値の大小関係が日付文字列の大小関係と一致することを確認します。
        proptest! {
            #[test]
            fn test_date_conversion_monotonicity(
                serial1 in 1.0f64..50000.0,
                serial2 in 1.0f64..50000.0
            ) {
                let formatter = DateFormatter;
                let config = ConversionConfig {
                    date_format: DateFormat::Iso8601,
                    ..Default::default()
                };

                // 1900年エポックでテスト
                let date1 = formatter.format(serial1, &config, false).unwrap();
                let date2 = formatter.format(serial2, &config, false).unwrap();

                // シリアル値の大小関係が日付文字列の大小関係と一致することを確認
                if serial1 < serial2 {
                    prop_assert!(date1 < date2,
                        "Date monotonicity violated: serial1={} ({}) < serial2={} ({})",
                        serial1, date1, serial2, date2);
                } else if serial1 > serial2 {
                    prop_assert!(date1 > date2,
                        "Date monotonicity violated: serial1={} ({}) > serial2={} ({})",
                        serial1, date1, serial2, date2);
                } else {
                    // シリアル値が等しい場合は、日付文字列も等しいこと
                    prop_assert_eq!(date1.clone(), date2.clone(),
                        "Date equality violated: serial1={} ({}) == serial2={} ({})",
                        serial1, date1, serial2, date2);
                }
            }
        }

        #[allow(unused_doc_comments)]
        /// TC-PBT-002の拡張: 1904年エポックでの単調性テスト
        proptest! {
            #[test]
            fn test_date_conversion_monotonicity_1904(
                serial1 in 0.0f64..50000.0,
                serial2 in 0.0f64..50000.0
            ) {
                let formatter = DateFormatter;
                let config = ConversionConfig {
                    date_format: DateFormat::Iso8601,
                    ..Default::default()
                };

                // 1904年エポックでテスト
                let date1 = formatter.format(serial1, &config, true).unwrap();
                let date2 = formatter.format(serial2, &config, true).unwrap();

                // シリアル値の大小関係が日付文字列の大小関係と一致することを確認
                if serial1 < serial2 {
                    prop_assert!(date1 < date2,
                        "Date monotonicity violated (1904): serial1={} ({}) < serial2={} ({})",
                        serial1, date1, serial2, date2);
                } else if serial1 > serial2 {
                    prop_assert!(date1 > date2,
                        "Date monotonicity violated (1904): serial1={} ({}) > serial2={} ({})",
                        serial1, date1, serial2, date2);
                } else {
                    // シリアル値が等しい場合は、日付文字列も等しいこと
                    prop_assert_eq!(date1.clone(), date2.clone(),
                        "Date equality violated (1904): serial1={} ({}) == serial2={} ({})",
                        serial1, date1, serial2, date2);
                }
            }
        }
    }
}
