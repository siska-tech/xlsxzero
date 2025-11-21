//! Output Format Module
//!
//! Strategy Patternによる出力フォーマットの抽象化を提供するモジュール。

mod formatters;

use crate::error::XlsxToMdError;
use crate::grid::LogicalGrid;
use crate::types::MergedRegion;
use std::io::Write;

pub use formatters::*;

/// 出力フォーマッター（Strategy Pattern）
///
/// 各出力フォーマット（Markdown, HTML, JSON, CSV）をenumとして表現します。
#[derive(Debug, Clone, Copy)]
pub enum OutputFormatter {
    Markdown,
    Html,
    Json,
    Csv,
}

impl OutputFormatter {
    /// 出力フォーマットからフォーマッターを生成
    pub fn from_format(format: crate::api::OutputFormat) -> Self {
        match format {
            crate::api::OutputFormat::Markdown => OutputFormatter::Markdown,
            crate::api::OutputFormat::Html => OutputFormatter::Html,
            crate::api::OutputFormat::Json => OutputFormatter::Json,
            crate::api::OutputFormat::Csv => OutputFormatter::Csv,
        }
    }

    /// グリッドを指定されたフォーマットで出力する
    ///
    /// # 引数
    ///
    /// * `grid` - 出力するグリッド
    /// * `writer` - 出力先のライター
    /// * `merged_regions` - 結合セル範囲のリスト（HTML形式で使用）
    ///
    /// # 戻り値
    ///
    /// * `Ok(())` - 出力に成功した場合
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    pub fn render<W: Write>(
        &self,
        grid: &LogicalGrid,
        writer: &mut W,
        merged_regions: &[MergedRegion],
    ) -> Result<(), XlsxToMdError> {
        match self {
            OutputFormatter::Markdown => {
                MarkdownFormatter.render(grid, writer, merged_regions)
            }
            OutputFormatter::Html => {
                HtmlFormatter.render(grid, writer, merged_regions)
            }
            OutputFormatter::Json => {
                JsonFormatter.render(grid, writer, merged_regions)
            }
            OutputFormatter::Csv => {
                CsvFormatter.render(grid, writer, merged_regions)
            }
        }
    }
}

