//! FormatSection Module
//!
//! Excel Number Format Stringのセクション定義を提供します。

use super::tokens::FormatToken;

/// セクションの種類
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SectionKind {
    /// 正数
    Positive,
    /// 負数
    Negative,
    /// ゼロ
    Zero,
    /// テキスト
    Text,
}

/// 条件（Phase II制限: 条件付き書式はサポート外）
///
/// 将来的な拡張のために定義されていますが、Phase IIでは使用されません。
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum Condition {
    /// 条件付き書式（例: [>100], [<0], [=0]）
    /// Phase II制限: サポート外（フォールバック）
    #[allow(dead_code)]
    Conditional(String),
}

/// フォーマットのセクション（正数、負数、ゼロ、テキスト）
///
/// Excel Number Format Stringは最大4つのセクションに分割されます:
/// 1. 正数
/// 2. 負数
/// 3. ゼロ
/// 4. テキスト
#[derive(Debug, Clone)]
pub(crate) struct FormatSection {
    /// セクションの種類
    pub kind: SectionKind,

    /// 条件（例: [>100]）
    /// Phase II制限: 条件付き書式はサポート外（None固定）
    #[allow(dead_code)]
    pub condition: Option<Condition>,

    /// フォーマットトークン
    pub tokens: Vec<FormatToken>,
}

impl FormatSection {
    /// 新しいセクションを生成
    pub fn new(kind: SectionKind) -> Self {
        Self {
            kind,
            condition: None,
            tokens: Vec::new(),
        }
    }

    /// セクションが日付・時刻書式かどうかを判定
    pub fn is_datetime(&self) -> bool {
        self.tokens.iter().any(|t| t.is_datetime())
    }

    /// セクションが数値書式かどうかを判定
    pub fn is_numeric(&self) -> bool {
        self.tokens.iter().any(|t| t.is_numeric())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_section_new() {
        let section = FormatSection::new(SectionKind::Positive);
        assert_eq!(section.kind, SectionKind::Positive);
        assert!(section.condition.is_none());
        assert!(section.tokens.is_empty());
    }

    #[test]
    fn test_is_datetime() {
        let mut section = FormatSection::new(SectionKind::Positive);
        section.tokens.push(FormatToken::Year(4));
        section.tokens.push(FormatToken::Month(2));
        section.tokens.push(FormatToken::Day(2));
        assert!(section.is_datetime());

        let mut section2 = FormatSection::new(SectionKind::Positive);
        section2.tokens.push(FormatToken::IntegerZero(1));
        assert!(!section2.is_datetime());
    }

    #[test]
    fn test_is_numeric() {
        let mut section = FormatSection::new(SectionKind::Positive);
        section.tokens.push(FormatToken::IntegerZero(1));
        section.tokens.push(FormatToken::DecimalPoint);
        section.tokens.push(FormatToken::DecimalZero(2));
        assert!(section.is_numeric());

        let mut section2 = FormatSection::new(SectionKind::Positive);
        section2.tokens.push(FormatToken::Year(4));
        assert!(!section2.is_numeric());
    }
}
