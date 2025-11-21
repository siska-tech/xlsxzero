//! FormatToken Module
//!
//! Excel Number Format Stringのトークン定義を提供します。

/// フォーマットトークン
///
/// Excel Number Format Stringを解析した際に生成されるトークンです。
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum FormatToken {
    /// 年（例: "yyyy" -> 4桁, "yy" -> 2桁）
    Year(usize),

    /// 月（例: "mm" -> 2桁, "m" -> 1桁）
    Month(usize),

    /// 日（例: "dd" -> 2桁, "d" -> 1桁）
    Day(usize),

    /// 時（例: "hh" -> 2桁, "h" -> 1桁）
    Hour(usize),

    /// 分（例: "mm" -> 2桁, "m" -> 1桁）
    /// 注意: 日付書式では"mm"は月、時刻書式では"mm"は分
    Minute(usize),

    /// 秒（例: "ss" -> 2桁, "s" -> 1桁）
    Second(usize),

    /// 整数部のゼロパディング（例: "0" -> 1桁, "00" -> 2桁）
    IntegerZero(usize),

    /// 整数部の空白パディング（例: "#"）
    IntegerHash,

    /// 小数点
    DecimalPoint,

    /// 小数部のゼロパディング（例: "0" -> 1桁, "00" -> 2桁）
    DecimalZero(usize),

    /// 千の位区切り
    ThousandSeparator,

    /// パーセント記号
    Percent,

    /// リテラル文字列（例: "$", "-", " "）
    Literal(String),

    /// 色指定（例: "[Red]", "[Blue]"）
    /// Phase II制限: 色指定は無視されます
    Color(String),

    /// テキストプレースホルダー（例: "@"）
    TextPlaceholder,
}

impl FormatToken {
    /// トークンが日付・時刻関連かどうかを判定
    pub fn is_datetime(&self) -> bool {
        matches!(
            self,
            FormatToken::Year(_)
                | FormatToken::Month(_)
                | FormatToken::Day(_)
                | FormatToken::Hour(_)
                | FormatToken::Minute(_)
                | FormatToken::Second(_)
        )
    }

    /// トークンが数値関連かどうかを判定
    pub fn is_numeric(&self) -> bool {
        matches!(
            self,
            FormatToken::IntegerZero(_)
                | FormatToken::IntegerHash
                | FormatToken::DecimalPoint
                | FormatToken::DecimalZero(_)
                | FormatToken::ThousandSeparator
                | FormatToken::Percent
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_is_datetime() {
        assert!(FormatToken::Year(4).is_datetime());
        assert!(FormatToken::Month(2).is_datetime());
        assert!(FormatToken::Day(2).is_datetime());
        assert!(FormatToken::Hour(2).is_datetime());
        assert!(FormatToken::Minute(2).is_datetime());
        assert!(FormatToken::Second(2).is_datetime());
        assert!(!FormatToken::IntegerZero(1).is_datetime());
        assert!(!FormatToken::Literal("$".to_string()).is_datetime());
    }

    #[test]
    fn test_is_numeric() {
        assert!(FormatToken::IntegerZero(1).is_numeric());
        assert!(FormatToken::IntegerHash.is_numeric());
        assert!(FormatToken::DecimalPoint.is_numeric());
        assert!(FormatToken::DecimalZero(2).is_numeric());
        assert!(FormatToken::ThousandSeparator.is_numeric());
        assert!(FormatToken::Percent.is_numeric());
        assert!(!FormatToken::Year(4).is_numeric());
        assert!(!FormatToken::Literal("$".to_string()).is_numeric());
    }
}
