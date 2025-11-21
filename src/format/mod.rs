//! Format Module
//!
//! Excel Number Format Stringの構文解析と適用を提供します。

mod parser;
mod sections;
mod tokens;

pub(crate) use parser::FormatParser;
