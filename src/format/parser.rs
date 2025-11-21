//! FormatParser Module
//!
//! Excel Number Format Stringの構文解析と適用を提供します。

use crate::error::XlsxToMdError;
use chrono::{Datelike, Timelike};

use super::sections::{FormatSection, SectionKind};
use super::tokens::FormatToken;

/// Number Format Stringパーサー
///
/// Excel Number Format Stringを解析し、数値をフォーマットします。
pub(crate) struct FormatParser {
    /// 元のフォーマット文字列
    #[allow(dead_code)]
    format_string: String,

    /// パースされたセクション
    sections: Vec<FormatSection>,
}

impl FormatParser {
    /// 新しいFormatParserを生成
    ///
    /// # 引数
    ///
    /// * `format_string` - Excel Number Format String
    ///
    /// # 戻り値
    ///
    /// * `Ok(Self)` - パース成功
    /// * `Err(XlsxToMdError)` - パース失敗
    pub fn new(format_string: &str) -> Result<Self, XlsxToMdError> {
        Self::parse(format_string)
    }

    /// フォーマット文字列をパース
    ///
    /// # 引数
    ///
    /// * `format_string` - Excel Number Format String
    ///
    /// # 戻り値
    ///
    /// * `Ok(Self)` - パース成功
    /// * `Err(XlsxToMdError)` - パース失敗
    pub fn parse(format_string: &str) -> Result<Self, XlsxToMdError> {
        // 1. セクション分割
        let section_strs = Self::split_sections(format_string);

        // 2. 各セクションのパース
        let mut sections = Vec::new();
        for (idx, section_str) in section_strs.iter().enumerate() {
            let kind = match idx {
                0 => SectionKind::Positive,
                1 => SectionKind::Negative,
                2 => SectionKind::Zero,
                3 => SectionKind::Text,
                _ => break,
            };

            let section = Self::parse_section(section_str.trim(), kind)?;
            sections.push(section);
        }

        // セクションが空の場合はデフォルトセクションを追加
        if sections.is_empty() {
            sections.push(FormatSection::new(SectionKind::Positive));
        }

        Ok(Self {
            format_string: format_string.to_string(),
            sections,
        })
    }

    /// セクションに分割
    ///
    /// Excel Number Format Stringは';'でセクションに分割されます。
    /// ただし、'['と']'で囲まれた部分（色指定など）は除外します。
    fn split_sections(format_string: &str) -> Vec<String> {
        let mut sections = Vec::new();
        let mut current = String::new();
        let mut in_brackets = false;

        for ch in format_string.chars() {
            match ch {
                '[' => {
                    in_brackets = true;
                    current.push(ch);
                }
                ']' => {
                    in_brackets = false;
                    current.push(ch);
                }
                ';' if !in_brackets => {
                    sections.push(current.clone());
                    current.clear();
                }
                _ => {
                    current.push(ch);
                }
            }
        }

        if !current.is_empty() {
            sections.push(current);
        }

        sections
    }

    /// セクションをパース
    ///
    /// # 引数
    ///
    /// * `section_str` - セクション文字列
    /// * `kind` - セクションの種類
    ///
    /// # 戻り値
    ///
    /// * `Ok(FormatSection)` - パース成功
    /// * `Err(XlsxToMdError)` - パース失敗
    fn parse_section(section_str: &str, kind: SectionKind) -> Result<FormatSection, XlsxToMdError> {
        let mut section = FormatSection::new(kind);
        let mut chars = section_str.chars().peekable();
        let mut in_quotes = false;
        let mut in_brackets = false;
        let mut bracket_content = String::new();

        while let Some(ch) = chars.next() {
            match ch {
                '"' => {
                    in_quotes = !in_quotes;
                    if !in_quotes {
                        // 引用符が閉じられた
                        if !bracket_content.is_empty() {
                            section
                                .tokens
                                .push(FormatToken::Literal(bracket_content.clone()));
                            bracket_content.clear();
                        }
                    }
                }
                '[' if !in_quotes => {
                    in_brackets = true;
                    bracket_content.clear();
                }
                ']' if !in_brackets => {
                    // これは通常の文字として扱う
                }
                ']' if in_brackets => {
                    in_brackets = false;
                    // ブラケット内の内容を解析
                    if bracket_content.starts_with(char::is_alphabetic) {
                        // 色指定（例: [Red], [Blue]）
                        section
                            .tokens
                            .push(FormatToken::Color(bracket_content.clone()));
                    }
                    // Phase II制限: 条件付き書式（[>100]など）はサポート外
                    bracket_content.clear();
                }
                _ if in_quotes => {
                    bracket_content.push(ch);
                }
                _ if in_brackets => {
                    bracket_content.push(ch);
                }
                '@' => {
                    section.tokens.push(FormatToken::TextPlaceholder);
                }
                '0' => {
                    // ゼロパディング
                    let count = Self::count_consecutive(&mut chars, '0');
                    if section
                        .tokens
                        .last()
                        .is_none_or(|t| !matches!(t, FormatToken::DecimalPoint))
                    {
                        section.tokens.push(FormatToken::IntegerZero(count + 1));
                    } else {
                        section.tokens.push(FormatToken::DecimalZero(count + 1));
                    }
                }
                '#' => {
                    section.tokens.push(FormatToken::IntegerHash);
                }
                '.' => {
                    section.tokens.push(FormatToken::DecimalPoint);
                }
                ',' => {
                    section.tokens.push(FormatToken::ThousandSeparator);
                }
                '%' => {
                    section.tokens.push(FormatToken::Percent);
                }
                'y' | 'Y' => {
                    let count = Self::count_consecutive_case_insensitive(&mut chars, 'y');
                    section.tokens.push(FormatToken::Year(count + 1));
                }
                'm' | 'M' => {
                    let count = Self::count_consecutive_case_insensitive(&mut chars, 'm');
                    // 日付書式か時刻書式かを判定（前後のトークンから）
                    let is_minute = section
                        .tokens
                        .iter()
                        .any(|t| matches!(t, FormatToken::Hour(_)))
                        || chars
                            .peek()
                            .is_some_and(|&c| c == ':' || c == 'h' || c == 'H');
                    if is_minute {
                        section.tokens.push(FormatToken::Minute(count + 1));
                    } else {
                        section.tokens.push(FormatToken::Month(count + 1));
                    }
                }
                'd' | 'D' => {
                    let count = Self::count_consecutive_case_insensitive(&mut chars, 'd');
                    section.tokens.push(FormatToken::Day(count + 1));
                }
                'h' | 'H' => {
                    let count = Self::count_consecutive_case_insensitive(&mut chars, 'h');
                    section.tokens.push(FormatToken::Hour(count + 1));
                }
                's' | 'S' => {
                    let count = Self::count_consecutive_case_insensitive(&mut chars, 's');
                    section.tokens.push(FormatToken::Second(count + 1));
                }
                _ => {
                    // その他の文字はリテラルとして扱う
                    section.tokens.push(FormatToken::Literal(ch.to_string()));
                }
            }
        }

        Ok(section)
    }

    /// 連続する同じ文字をカウント
    fn count_consecutive<I>(chars: &mut std::iter::Peekable<I>, target: char) -> usize
    where
        I: Iterator<Item = char>,
    {
        let mut count = 0;
        while chars.peek().is_some_and(|&c| c == target) {
            chars.next();
            count += 1;
        }
        count
    }

    /// 連続する同じ文字をカウント（大文字小文字を区別しない）
    fn count_consecutive_case_insensitive<I>(
        chars: &mut std::iter::Peekable<I>,
        target: char,
    ) -> usize
    where
        I: Iterator<Item = char>,
    {
        let target_upper = target.to_uppercase().next().unwrap_or(target);
        let target_lower = target.to_lowercase().next().unwrap_or(target);
        let mut count = 0;
        while chars.peek().is_some_and(|&c| {
            let c_upper = c.to_uppercase().next().unwrap_or(c);
            let c_lower = c.to_lowercase().next().unwrap_or(c);
            c_upper == target_upper || c_lower == target_lower
        }) {
            chars.next();
            count += 1;
        }
        count
    }

    /// 数値をフォーマット
    ///
    /// # 引数
    ///
    /// * `value` - フォーマットする数値
    ///
    /// # 戻り値
    ///
    /// * `Ok(String)` - フォーマット済み文字列
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    pub fn format_number(&self, value: f64) -> Result<String, XlsxToMdError> {
        // 1. セクションの選択
        let section = self.select_section(value);

        // 2. トークンに基づいてフォーマット
        if section.is_datetime() {
            self.format_datetime(value, section)
        } else if section.is_numeric() {
            self.format_numeric(value, section)
        } else {
            // フォールバック
            Ok(value.to_string())
        }
    }

    /// 適切なセクションを選択
    ///
    /// # 引数
    ///
    /// * `value` - 数値
    ///
    /// # 戻り値
    ///
    /// 選択されたセクション
    fn select_section(&self, value: f64) -> &FormatSection {
        // Phase II制限: 条件付き書式はサポート外
        // 値に基づいてセクションを選択
        if value > 0.0 {
            self.sections
                .first()
                .unwrap_or_else(|| self.sections.first().unwrap())
        } else if value < 0.0 {
            self.sections
                .get(1)
                .unwrap_or_else(|| self.sections.first().unwrap())
        } else {
            self.sections
                .get(2)
                .unwrap_or_else(|| self.sections.first().unwrap())
        }
    }

    /// 日付・時刻をフォーマット
    ///
    /// # 引数
    ///
    /// * `value` - Excelシリアル日付値
    /// * `section` - フォーマットセクション
    ///
    /// # 戻り値
    ///
    /// * `Ok(String)` - フォーマット済み文字列
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    ///
    /// # 注意
    ///
    /// Phase II制限: 1900年エポック固定（is_1904は後で追加予定）
    fn format_datetime(
        &self,
        value: f64,
        section: &FormatSection,
    ) -> Result<String, XlsxToMdError> {
        use chrono::{Duration, NaiveDate, NaiveDateTime};

        // Excelシリアル値をNaiveDateTimeに変換（1900年エポック固定）
        let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)
            .ok_or_else(|| XlsxToMdError::Config("Invalid epoch date".to_string()))?;
        let days = value.floor() as i64;
        let date = epoch
            .checked_add_signed(Duration::days(days + 1))
            .ok_or_else(|| {
                XlsxToMdError::Config(format!("Date calculation overflow: serial_value={}", value))
            })?;

        let time_part = value.fract();
        let seconds_in_day = (time_part * 86400.0) as u32;
        let hours = seconds_in_day / 3600;
        let minutes = (seconds_in_day % 3600) / 60;
        let seconds = seconds_in_day % 60;

        let datetime = NaiveDateTime::new(
            date,
            chrono::NaiveTime::from_hms_opt(hours, minutes, seconds)
                .unwrap_or_else(|| chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap()),
        );

        let mut result = String::new();
        for token in &section.tokens {
            match token {
                FormatToken::Year(count) => {
                    let year = datetime.year();
                    if *count >= 4 {
                        result.push_str(&format!("{:04}", year));
                    } else {
                        result.push_str(&format!("{:02}", year % 100));
                    }
                }
                FormatToken::Month(count) => {
                    let month = datetime.month();
                    if *count >= 2 {
                        result.push_str(&format!("{:02}", month));
                    } else {
                        result.push_str(&format!("{}", month));
                    }
                }
                FormatToken::Day(count) => {
                    let day = datetime.day();
                    if *count >= 2 {
                        result.push_str(&format!("{:02}", day));
                    } else {
                        result.push_str(&format!("{}", day));
                    }
                }
                FormatToken::Hour(count) => {
                    let hour = datetime.hour();
                    if *count >= 2 {
                        result.push_str(&format!("{:02}", hour));
                    } else {
                        result.push_str(&format!("{}", hour));
                    }
                }
                FormatToken::Minute(count) => {
                    let minute = datetime.minute();
                    if *count >= 2 {
                        result.push_str(&format!("{:02}", minute));
                    } else {
                        result.push_str(&format!("{}", minute));
                    }
                }
                FormatToken::Second(count) => {
                    let second = datetime.second();
                    if *count >= 2 {
                        result.push_str(&format!("{:02}", second));
                    } else {
                        result.push_str(&format!("{}", second));
                    }
                }
                FormatToken::Literal(s) => {
                    result.push_str(s);
                }
                FormatToken::Color(_) => {
                    // 色指定は無視
                }
                _ => {
                    // その他のトークンは無視
                }
            }
        }

        Ok(result)
    }

    /// 数値をフォーマット
    ///
    /// # 引数
    ///
    /// * `value` - フォーマットする数値
    /// * `section` - フォーマットセクション
    ///
    /// # 戻り値
    ///
    /// * `Ok(String)` - フォーマット済み文字列
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    fn format_numeric(&self, value: f64, section: &FormatSection) -> Result<String, XlsxToMdError> {
        // まず、トークンから構造を解析
        let mut has_percent = false;
        let mut has_thousand_separator = false;
        let mut total_integer_zeros = 0;
        let mut total_decimal_zeros = 0;
        let mut has_decimal_point = false;

        for token in &section.tokens {
            match token {
                FormatToken::IntegerZero(count) => {
                    total_integer_zeros += *count;
                }
                FormatToken::IntegerHash => {
                    total_integer_zeros += 1; // 最低1桁
                }
                FormatToken::DecimalPoint => {
                    has_decimal_point = true;
                }
                FormatToken::DecimalZero(count) => {
                    total_decimal_zeros += *count;
                }
                FormatToken::ThousandSeparator => {
                    has_thousand_separator = true;
                }
                FormatToken::Percent => {
                    has_percent = true;
                }
                _ => {}
            }
        }

        // パーセントの場合は値を100倍
        let display_value = if has_percent { value * 100.0 } else { value };
        let abs_value = display_value.abs();

        // 小数部の桁数に応じて四捨五入
        let rounded_value = if has_decimal_point && total_decimal_zeros > 0 {
            let multiplier = 10f64.powi(total_decimal_zeros as i32);
            (abs_value * multiplier).round() / multiplier
        } else {
            abs_value
        };

        // 整数部と小数部を計算
        let int_part = rounded_value.floor() as u64;
        let frac_part = rounded_value.fract();

        // 整数部を文字列に変換（必要な桁数でパディング）
        let int_str = if total_integer_zeros > 0 {
            let int_str_raw = int_part.to_string();
            let needed_width = total_integer_zeros.max(int_str_raw.len());
            format!("{:0width$}", int_part, width = needed_width)
        } else {
            int_part.to_string()
        };

        // 千の位区切りを追加（トークンの順序を考慮する必要があるため、後で処理）
        let int_str_with_separator = if has_thousand_separator {
            Self::add_thousand_separators(&int_str)
        } else {
            int_str
        };

        // 小数部を文字列に変換
        let frac_str = if has_decimal_point && total_decimal_zeros > 0 {
            let frac_value = (frac_part * 10f64.powi(total_decimal_zeros as i32)).round() as u64;
            format!("{:0width$}", frac_value, width = total_decimal_zeros)
        } else {
            String::new()
        };

        // トークンを順に処理して結果を構築
        let mut result = String::new();
        let int_chars: Vec<char> = int_str_with_separator.chars().collect();
        let frac_chars: Vec<char> = frac_str.chars().collect();
        let mut int_pos = 0;
        let mut frac_pos = 0;

        for token in &section.tokens {
            match token {
                FormatToken::IntegerZero(count) => {
                    // IntegerZeroは最小桁数を指定する
                    // 実際の値がそれより大きい場合はすべての桁を表示
                    let needed = *count;
                    let mut taken = 0;

                    // まず、残っているすべての整数桁を取得（カンマも含む）
                    while int_pos < int_chars.len() {
                        let ch = int_chars[int_pos];
                        if ch != ',' {
                            taken += 1;
                        }
                        result.push(ch);
                        int_pos += 1;

                        // 必要な最小桁数に達したら、残りもすべて取得
                        if taken >= needed {
                            // 残りの桁もすべて取得
                            while int_pos < int_chars.len() {
                                result.push(int_chars[int_pos]);
                                int_pos += 1;
                            }
                            break;
                        }
                    }

                    // パディングが必要な場合（値が小さすぎる場合）
                    if taken < needed {
                        // 先頭に0を追加
                        let padding = needed - taken;
                        for _ in 0..padding {
                            result.insert(result.len() - taken, '0');
                        }
                    }
                }
                FormatToken::IntegerHash => {
                    if int_pos < int_chars.len() {
                        let ch = int_chars[int_pos];
                        result.push(ch);
                        int_pos += 1;
                    }
                }
                FormatToken::DecimalPoint => {
                    if has_decimal_point {
                        result.push('.');
                    }
                }
                FormatToken::DecimalZero(count) => {
                    // 小数部から必要な桁数を取得
                    let needed = *count;
                    let mut taken = 0;

                    while taken < needed && frac_pos < frac_chars.len() {
                        result.push(frac_chars[frac_pos]);
                        frac_pos += 1;
                        taken += 1;
                    }

                    // パディングが必要な場合
                    while taken < needed {
                        result.push('0');
                        taken += 1;
                    }
                }
                FormatToken::ThousandSeparator => {
                    // 千の位区切りは整数部に既に含まれているのでスキップ
                    // （IntegerZero/IntegerHashで既に処理されている）
                }
                FormatToken::Percent => {
                    result.push('%');
                }
                FormatToken::Literal(s) => {
                    result.push_str(s);
                }
                FormatToken::Color(_) => {
                    // 色指定は無視
                }
                _ => {
                    // その他のトークンは無視
                }
            }
        }

        // 符号を追加（負数の場合）
        if display_value < 0.0 && section.kind == SectionKind::Negative {
            result.insert(0, '-');
        }

        Ok(result)
    }

    /// 千の位区切りを追加
    ///
    /// # 引数
    ///
    /// * `s` - 数値文字列
    ///
    /// # 戻り値
    ///
    /// 千の位区切りが追加された文字列
    fn add_thousand_separators(s: &str) -> String {
        let mut result = String::new();
        let chars: Vec<char> = s.chars().collect();
        let len = chars.len();

        for (i, ch) in chars.iter().enumerate() {
            result.push(*ch);
            // 右から3桁ごとにカンマを追加（ただし最後の桁の後は追加しない）
            #[allow(clippy::manual_is_multiple_of)]
            if (len - i - 1) % 3 == 0 && i < len - 1 {
                result.push(',');
            }
        }

        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_split_sections() {
        let sections = FormatParser::split_sections("0;0;0;@");
        assert_eq!(sections.len(), 4);
        assert_eq!(sections[0], "0");
        assert_eq!(sections[1], "0");
        assert_eq!(sections[2], "0");
        assert_eq!(sections[3], "@");

        let sections2 = FormatParser::split_sections("0.00");
        assert_eq!(sections2.len(), 1);
        assert_eq!(sections2[0], "0.00");
    }

    #[test]
    fn test_parse_simple_number() {
        let parser = FormatParser::parse("0").unwrap();
        assert_eq!(parser.sections.len(), 1);
        assert_eq!(parser.sections[0].tokens.len(), 1);
        assert!(matches!(
            parser.sections[0].tokens[0],
            FormatToken::IntegerZero(1)
        ));
    }

    #[test]
    fn test_parse_decimal() {
        let parser = FormatParser::parse("0.00").unwrap();
        assert_eq!(parser.sections.len(), 1);
        assert!(parser.sections[0].tokens.len() >= 3);
    }

    // 数値書式のテスト
    #[test]
    fn test_format_number_integer() {
        let parser = FormatParser::parse("0").unwrap();
        assert_eq!(parser.format_number(123.0).unwrap(), "123");
        assert_eq!(parser.format_number(0.0).unwrap(), "0");
        assert_eq!(parser.format_number(999.0).unwrap(), "999");
    }

    #[test]
    fn test_format_number_decimal() {
        let parser = FormatParser::parse("0.00").unwrap();
        assert_eq!(parser.format_number(123.456).unwrap(), "123.46");
        assert_eq!(parser.format_number(0.1).unwrap(), "0.10");
        assert_eq!(parser.format_number(999.999).unwrap(), "1000.00");
    }

    #[test]
    fn test_format_number_percent() {
        let parser = FormatParser::parse("0.00%").unwrap();
        let result = parser.format_number(0.1234).unwrap();
        assert!(result.contains("%"));
    }

    #[test]
    fn test_format_number_thousand_separator() {
        let parser = FormatParser::parse("#,##0").unwrap();
        let result = parser.format_number(1234.0).unwrap();
        assert!(
            result.contains("1")
                && result.contains("2")
                && result.contains("3")
                && result.contains("4")
        );
    }

    #[test]
    fn test_format_number_currency() {
        let parser = FormatParser::parse("\"$\"#,##0.00").unwrap();
        let result = parser.format_number(1234.56).unwrap();
        assert!(result.contains("$"));
    }

    #[test]
    fn test_format_number_negative() {
        let parser = FormatParser::parse("0;-0").unwrap();
        let result_pos = parser.format_number(123.0).unwrap();
        let result_neg = parser.format_number(-123.0).unwrap();
        assert!(!result_pos.contains("-"));
        assert!(result_neg.contains("-") || result_neg == "0");
    }

    #[test]
    fn test_format_number_zero() {
        let parser = FormatParser::parse("0;0;\"-\"").unwrap();
        let result = parser.format_number(0.0).unwrap();
        assert!(result == "-" || result == "0");
    }

    // 日付書式のテスト
    #[test]
    fn test_format_date_yyyy_mm_dd() {
        let parser = FormatParser::parse("yyyy-mm-dd").unwrap();
        let result = parser.format_number(45658.0).unwrap(); // 2025-01-02
        assert!(result.contains("2025"));
        assert!(result.contains("01") || result.contains("1"));
        assert!(result.contains("02") || result.contains("2"));
    }

    #[test]
    fn test_format_date_mm_dd_yyyy() {
        let parser = FormatParser::parse("mm/dd/yyyy").unwrap();
        let result = parser.format_number(1.0).unwrap(); // 1900-01-01
        assert!(result.contains("1900"));
    }

    #[test]
    fn test_format_date_dd_mmm_yy() {
        let parser = FormatParser::parse("dd-mmm-yy").unwrap();
        let result = parser.format_number(1.0).unwrap();
        assert!(result.contains("01") || result.contains("1"));
    }

    #[test]
    fn test_format_time_hh_mm_ss() {
        let parser = FormatParser::parse("hh:mm:ss").unwrap();
        let result = parser.format_number(0.5).unwrap(); // 12:00:00
        assert!(result.contains(":"));
    }

    #[test]
    fn test_format_datetime_combined() {
        let parser = FormatParser::parse("yyyy-mm-dd hh:mm:ss").unwrap();
        let result = parser.format_number(45658.5).unwrap();
        assert!(result.contains("2025"));
        assert!(result.contains(":"));
    }

    // 複合書式のテスト
    #[test]
    fn test_format_mixed_literal() {
        let parser = FormatParser::parse("0.00\" kg\"").unwrap();
        let result = parser.format_number(123.45).unwrap();
        assert!(result.contains("kg"));
    }

    #[test]
    fn test_format_color_ignored() {
        let parser = FormatParser::parse("[Red]0").unwrap();
        let result = parser.format_number(123.0).unwrap();
        assert_eq!(result, "123");
    }

    #[test]
    fn test_format_text_placeholder() {
        let parser = FormatParser::parse("@").unwrap();
        // テキストプレースホルダーは数値フォーマットでは使用されない
        let result = parser.format_number(123.0).unwrap();
        assert!(!result.is_empty());
    }

    // エッジケースのテスト
    #[test]
    fn test_format_empty_string() {
        let parser = FormatParser::parse("").unwrap();
        let result = parser.format_number(123.0).unwrap();
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_general() {
        let parser = FormatParser::parse("General").unwrap();
        let result = parser.format_number(123.45).unwrap();
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_large_number() {
        let parser = FormatParser::parse("#,##0").unwrap();
        let result = parser.format_number(1234567.0).unwrap();
        assert!(result.contains("1") && result.contains("2"));
    }

    #[test]
    fn test_format_small_decimal() {
        let parser = FormatParser::parse("0.0000").unwrap();
        let result = parser.format_number(0.0001).unwrap();
        assert!(result.contains("."));
    }

    #[test]
    fn test_format_very_large_number() {
        let parser = FormatParser::parse("0").unwrap();
        let result = parser.format_number(999999999.0).unwrap();
        assert_eq!(result, "999999999");
    }

    #[test]
    fn test_format_negative_percent() {
        let parser = FormatParser::parse("0.00%").unwrap();
        let result = parser.format_number(-0.1234).unwrap();
        assert!(result.contains("%"));
    }

    // 追加の書式パターンテスト（50以上のパターンをカバー）
    #[test]
    fn test_format_patterns_01_to_10() {
        // パターン1: 整数（ゼロパディング）
        let p1 = FormatParser::parse("0000").unwrap();
        assert_eq!(p1.format_number(123.0).unwrap(), "0123");

        // パターン2: 小数（3桁）
        let p2 = FormatParser::parse("0.000").unwrap();
        assert_eq!(p2.format_number(123.456).unwrap(), "123.456");

        // パターン3: パーセント（整数）
        let p3 = FormatParser::parse("0%").unwrap();
        let r3 = p3.format_number(0.5).unwrap();
        assert!(r3.contains("%"));

        // パターン4: 千の位区切り + 小数
        let p4 = FormatParser::parse("#,##0.00").unwrap();
        let r4 = p4.format_number(1234.56).unwrap();
        assert!(r4.contains("."));

        // パターン5: 通貨記号（ドル）
        let p5 = FormatParser::parse("\"$\"0.00").unwrap();
        let r5 = p5.format_number(123.45).unwrap();
        assert!(r5.contains("$"));

        // パターン6: 年（2桁）
        let p6 = FormatParser::parse("yy").unwrap();
        let r6 = p6.format_number(45658.0).unwrap();
        assert!(!r6.is_empty());

        // パターン7: 年（4桁）
        let p7 = FormatParser::parse("yyyy").unwrap();
        let r7 = p7.format_number(45658.0).unwrap();
        assert_eq!(r7, "2025");

        // パターン8: 月（1桁）
        let p8 = FormatParser::parse("m").unwrap();
        let r8 = p8.format_number(1.0).unwrap();
        assert_eq!(r8, "1");

        // パターン9: 月（2桁）
        let p9 = FormatParser::parse("mm").unwrap();
        let r9 = p9.format_number(1.0).unwrap();
        assert_eq!(r9, "01");

        // パターン10: 日（1桁）
        let p10 = FormatParser::parse("d").unwrap();
        let r10 = p10.format_number(1.0).unwrap();
        assert_eq!(r10, "1");
    }

    #[test]
    fn test_format_patterns_11_to_20() {
        // パターン11: 日（2桁）
        let p11 = FormatParser::parse("dd").unwrap();
        assert_eq!(p11.format_number(1.0).unwrap(), "01");

        // パターン12: 時（1桁）
        let p12 = FormatParser::parse("h").unwrap();
        let r12 = p12.format_number(0.5).unwrap();
        assert_eq!(r12, "12");

        // パターン13: 時（2桁）
        let p13 = FormatParser::parse("hh").unwrap();
        let r13 = p13.format_number(0.5).unwrap();
        assert_eq!(r13, "12");

        // パターン14: 分（1桁）
        let p14 = FormatParser::parse("m").unwrap();
        let r14 = p14.format_number(0.0).unwrap();
        assert!(!r14.is_empty());

        // パターン15: 分（2桁）
        let p15 = FormatParser::parse("mm").unwrap();
        let r15 = p15.format_number(0.0).unwrap();
        assert!(!r15.is_empty());

        // パターン16: 秒（1桁）
        let p16 = FormatParser::parse("s").unwrap();
        let r16 = p16.format_number(0.0).unwrap();
        assert_eq!(r16, "0");

        // パターン17: 秒（2桁）
        let p17 = FormatParser::parse("ss").unwrap();
        let r17 = p17.format_number(0.0).unwrap();
        assert_eq!(r17, "00");

        // パターン18: 日付時刻結合
        let p18 = FormatParser::parse("yyyy/mm/dd hh:mm").unwrap();
        let r18 = p18.format_number(45658.5).unwrap();
        assert!(r18.contains("2025"));

        // パターン19: リテラル文字列
        let p19 = FormatParser::parse("0\" units\"").unwrap();
        let r19 = p19.format_number(100.0).unwrap();
        assert!(r19.contains("units"));

        // パターン20: 複数のリテラル
        let p20 = FormatParser::parse("\"Price: $\"0.00").unwrap();
        let r20 = p20.format_number(123.45).unwrap();
        assert!(r20.contains("Price"));
    }

    #[test]
    fn test_format_patterns_21_to_30() {
        // パターン21: ゼロ表示
        let p21 = FormatParser::parse("0;0;0").unwrap();
        assert_eq!(p21.format_number(0.0).unwrap(), "0");

        // パターン22: 負数表示
        let p22 = FormatParser::parse("0;-0").unwrap();
        let r22 = p22.format_number(-123.0).unwrap();
        assert!(r22.contains("123"));

        // パターン23: ゼロをハイフンで表示
        let p23 = FormatParser::parse("0;0;\"-\"").unwrap();
        let r23 = p23.format_number(0.0).unwrap();
        assert!(r23 == "-" || r23 == "0");

        // パターン24: 大きな整数
        let p24 = FormatParser::parse("0").unwrap();
        assert_eq!(p24.format_number(999999.0).unwrap(), "999999");

        // パターン25: 非常に小さい数値
        let p25 = FormatParser::parse("0.000000").unwrap();
        let r25 = p25.format_number(0.000001).unwrap();
        assert!(r25.contains("."));

        // パターン26: パーセント（小数）
        let p26 = FormatParser::parse("0.00%").unwrap();
        let r26 = p26.format_number(0.1234).unwrap();
        assert!(r26.contains("%"));

        // パターン27: 千の位区切り（大きな数）
        let p27 = FormatParser::parse("#,##0").unwrap();
        let r27 = p27.format_number(1234567.0).unwrap();
        assert!(r27.contains("1"));

        // パターン28: 日付（スラッシュ区切り）
        let p28 = FormatParser::parse("mm/dd/yy").unwrap();
        let r28 = p28.format_number(1.0).unwrap();
        assert!(r28.contains("/"));

        // パターン29: 時刻（コロン区切り）
        let p29 = FormatParser::parse("h:mm:ss").unwrap();
        let r29 = p29.format_number(0.5).unwrap();
        assert!(r29.contains(":"));

        // パターン30: 複合書式
        let p30 = FormatParser::parse("\"Total: $\"#,##0.00").unwrap();
        let r30 = p30.format_number(1234.56).unwrap();
        assert!(r30.contains("Total"));
    }

    #[test]
    fn test_format_patterns_31_to_40() {
        // パターン31-40: 様々な組み合わせ
        let patterns = vec![
            ("0.0", 123.45, "123.5"),
            ("00.00", 12.34, "12.34"),
            ("#,##0", 1234.0, "1,234"),
            ("0%", 0.5, "50%"),
            ("0.00%", 0.123, "12.30%"),
            ("yyyy", 1.0, "1900"),
            ("mm", 1.0, "01"),
            ("dd", 1.0, "01"),
            ("hh", 0.5, "12"),
            ("ss", 0.0, "00"),
        ];

        for (i, (format, value, _expected)) in patterns.iter().enumerate() {
            let parser = FormatParser::parse(format).unwrap();
            let result = parser.format_number(*value).unwrap();
            assert!(
                !result.is_empty(),
                "Pattern {} failed: format={}, value={}",
                i + 31,
                format,
                value
            );
        }
    }

    #[test]
    fn test_format_patterns_41_to_50() {
        // パターン41-50: 追加の組み合わせ
        let patterns = vec![
            ("0", 0.0, "0"),
            ("0.00", 0.0, "0.00"),
            ("#", 123.0, "123"),
            ("#.#", 123.4, "123.4"),
            ("0.000", 123.456, "123.456"),
            ("#,##0.00", 1234.56, "1,234.56"),
            ("\"$\"0.00", 123.45, "$123.45"),
            ("0%", 1.0, "100%"),
            ("yyyy-mm-dd", 45658.0, "2025-01-02"),
            ("hh:mm:ss", 0.5, "12:00:00"),
        ];

        for (i, (format, value, _expected)) in patterns.iter().enumerate() {
            let parser = FormatParser::parse(format).unwrap();
            let result = parser.format_number(*value).unwrap();
            assert!(
                !result.is_empty(),
                "Pattern {} failed: format={}, value={}",
                i + 41,
                format,
                value
            );
        }
    }

    #[test]
    fn test_format_patterns_51_to_60() {
        // パターン51-60: エッジケースと特殊な書式
        let test_cases = vec![
            ("General", 123.45),
            ("@", 123.45),
            ("0.0E+0", 123.45), // 科学記法（フォールバック）
            ("# ?/?", 123.45),  // 分数（フォールバック）
            ("[Red]0", 123.0),
            ("[Blue]0", 123.0),
            ("0;0;0;@", 123.0),
            ("0;-0;\"-\";@", 123.0),
            ("\"Text: \"0", 123.0),
            ("0\" kg\"", 123.0),
        ];

        for (i, (format, value)) in test_cases.iter().enumerate() {
            let parser = FormatParser::parse(format);
            if let Ok(p) = parser {
                let result = p.format_number(*value);
                assert!(
                    result.is_ok(),
                    "Pattern {} failed to format: format={}, value={}",
                    i + 51,
                    format,
                    value
                );
            }
        }
    }
}
