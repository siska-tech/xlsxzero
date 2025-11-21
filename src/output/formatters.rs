//! Output Formatters Implementation
//!
//! 各出力フォーマットの実装を提供するモジュール。

use crate::error::XlsxToMdError;
use crate::grid::LogicalGrid;
use crate::types::MergedRegion;
use std::io::Write;

/// Markdown形式のフォーマッター
pub struct MarkdownFormatter;

impl MarkdownFormatter {
    pub fn render<W: Write>(
        &self,
        grid: &LogicalGrid,
        writer: &mut W,
        _merged_regions: &[MergedRegion],
    ) -> Result<(), XlsxToMdError> {
        grid.render_markdown(writer)
    }
}

/// HTML形式のフォーマッター
pub struct HtmlFormatter;

impl HtmlFormatter {
    pub fn render<W: Write>(
        &self,
        grid: &LogicalGrid,
        writer: &mut W,
        merged_regions: &[MergedRegion],
    ) -> Result<(), XlsxToMdError> {
        grid.render_html(writer, merged_regions)
    }
}

/// JSON形式のフォーマッター
pub struct JsonFormatter;

impl JsonFormatter {
    pub fn render<W: Write>(
        &self,
        grid: &LogicalGrid,
        writer: &mut W,
        _merged_regions: &[MergedRegion],
    ) -> Result<(), XlsxToMdError> {
        use serde_json::json;

        // グリッドの行と列の情報を取得
        let rows = grid.get_rows();
        let cols = grid.get_cols();

        if rows == 0 || cols == 0 {
            // 空のグリッドの場合は空のJSONオブジェクトを出力
            writeln!(writer, "{{}}")?;
            return Ok(());
        }

        // 列名を生成（A, B, C, ...）
        let column_names: Vec<String> = (0..cols)
            .map(|col| {
                let col_letter = col_to_letter(col as u32);
                col_letter
            })
            .collect();

        // 各行をオブジェクトとして構築
        let json_rows: Vec<serde_json::Value> = (0..rows)
            .map(|row_idx| {
                let row = grid.get_row(row_idx);
                let mut row_obj = serde_json::Map::new();

                for (col_idx, cell) in row.iter().enumerate() {
                    let col_name = &column_names[col_idx];
                    // 結合セルの子はスキップ（親セルのみ含める）
                    if !cell.is_merged || cell.merge_parent.is_none() {
                        row_obj.insert(col_name.clone(), json!(cell.content));
                    }
                }

                json!(row_obj)
            })
            .collect();

        // JSONオブジェクトを構築
        let json_output = json!({
            "rows": json_rows
        });

        // JSONを出力
        serde_json::to_writer_pretty(&mut *writer, &json_output).map_err(|e| {
            XlsxToMdError::Config(format!("JSON serialization error: {}", e))
        })?;
        writeln!(writer)?;
        writer.flush()?;

        Ok(())
    }
}

/// CSV形式のフォーマッター
pub struct CsvFormatter;

impl CsvFormatter {
    pub fn render<W: Write>(
        &self,
        grid: &LogicalGrid,
        writer: &mut W,
        _merged_regions: &[MergedRegion],
    ) -> Result<(), XlsxToMdError> {
        let rows = grid.get_rows();
        let cols = grid.get_cols();

        if rows == 0 || cols == 0 {
            return Ok(());
        }

        // 各行をCSV形式で出力
        for row_idx in 0..rows {
            let row = grid.get_row(row_idx);
            let mut first = true;

            for (_col_idx, cell) in row.iter().enumerate() {
                // 結合セルの子はスキップ（親セルのみ含める）
                if cell.is_merged && cell.merge_parent.is_some() {
                    continue;
                }

                if !first {
                    write!(writer, ",")?;
                }
                first = false;

                // CSVエスケープ処理
                let escaped = escape_csv(&cell.content);
                write!(writer, "{}", escaped)?;
            }

            writeln!(writer)?;
        }

        writer.flush()?;
        Ok(())
    }
}

/// 列インデックスをExcel列名（A, B, C, ...）に変換
fn col_to_letter(mut col: u32) -> String {
    let mut result = String::new();
    loop {
        result.push((b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result.chars().rev().collect()
}

/// CSV文字列をエスケープ
///
/// ダブルクォート、改行、カンマを含む場合はダブルクォートで囲み、
/// 内部のダブルクォートは2つにエスケープします。
fn escape_csv(s: &str) -> String {
    if s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r') {
        format!("\"{}\"", s.replace('"', "\"\""))
    } else {
        s.to_string()
    }
}

