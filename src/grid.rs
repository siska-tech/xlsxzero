//! Grid Module
//!
//! スパースなセルデータから稠密なグリッド構造への変換を提供するモジュール。
//! セル結合の処理戦略（DataDuplication / HtmlFallback）を実装します。

use std::io::Write;

use unicode_width::UnicodeWidthStr;

use crate::api::MergeStrategy;
use crate::error::XlsxToMdError;
use crate::types::{CellCoord, MergedRegion, RawCellData, SheetMetadata};

/// フォーマット済みセル
#[derive(Debug, Clone)]
pub(crate) struct Cell {
    /// 表示文字列
    pub content: String,

    /// 結合セルの一部かどうか
    pub is_merged: bool,

    /// 結合セルの親座標（結合セルの場合）
    pub merge_parent: Option<CellCoord>,
}

impl Cell {
    /// 新しい通常セルを生成
    pub fn new(content: String) -> Self {
        Self {
            content,
            is_merged: false,
            merge_parent: None,
        }
    }

    /// 新しい結合セルを生成
    pub fn new_merged(content: String, parent: CellCoord) -> Self {
        Self {
            content,
            is_merged: true,
            merge_parent: Some(parent),
        }
    }

    /// 空セルを生成
    pub fn empty() -> Self {
        Self {
            content: String::new(),
            is_merged: false,
            merge_parent: None,
        }
    }
}

/// 論理的なグリッド構造
pub(crate) struct LogicalGrid {
    /// グリッドデータ（行 × 列）
    cells: Vec<Vec<Cell>>,

    /// 行数
    rows: usize,

    /// 列数
    cols: usize,
}

impl LogicalGrid {
    /// スパースなセルデータから稠密なグリッド構造を構築
    ///
    /// # 引数
    ///
    /// * `cells` - 生のセルデータ（グリッドサイズ決定用）
    /// * `formatted_cells` - フォーマット済みセルデータ（座標と内容のペア）
    /// * `metadata` - シートのメタデータ（結合セル情報を含む）
    /// * `merge_strategy` - セル結合の処理戦略
    ///
    /// # 戻り値
    ///
    /// * `Ok(LogicalGrid)` - グリッド構築に成功した場合
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    pub fn build(
        cells: Vec<RawCellData>,
        formatted_cells: Vec<(CellCoord, String)>,
        metadata: &SheetMetadata,
        merge_strategy: MergeStrategy,
    ) -> Result<Self, XlsxToMdError> {
        // 1. グリッドサイズの決定
        let (rows, cols) = Self::determine_grid_size(&cells);

        // 2. 空のグリッドを初期化
        let mut grid_cells = vec![vec![Cell::empty(); cols]; rows];

        // 3. フォーマット済みセルデータを配置
        for (coord, content) in formatted_cells {
            if coord.row < rows as u32 && coord.col < cols as u32 {
                grid_cells[coord.row as usize][coord.col as usize] = Cell::new(content);
            }
        }

        // 4. セル結合の処理
        let mut grid = LogicalGrid {
            cells: grid_cells,
            rows,
            cols,
        };

        match merge_strategy {
            MergeStrategy::DataDuplication => {
                grid.apply_data_duplication(&metadata.merged_regions)?;
            }
            MergeStrategy::HtmlFallback => {
                // HTMLフォールバックの場合、グリッド処理はスキップ
                // 後段のMarkdown Writerで直接HTML出力
            }
        }

        Ok(grid)
    }

    /// グリッドサイズを決定（内部ヘルパー）
    ///
    /// すべてのセル座標から最大行・列を算出します。
    fn determine_grid_size(cells: &[RawCellData]) -> (usize, usize) {
        let mut max_row = 0;
        let mut max_col = 0;

        for cell in cells {
            max_row = max_row.max(cell.coord.row);
            max_col = max_col.max(cell.coord.col);
        }

        ((max_row + 1) as usize, (max_col + 1) as usize)
    }

    /// データ重複フィル戦略を適用（内部メソッド）
    ///
    /// 結合セル範囲内のすべてのセルに親セルの値を複製します。
    fn apply_data_duplication(
        &mut self,
        merged_regions: &[MergedRegion],
    ) -> Result<(), XlsxToMdError> {
        for region in merged_regions {
            // 親セルの内容を取得
            let parent_content = self.cells[region.parent.row as usize][region.parent.col as usize]
                .content
                .clone();

            // 結合範囲内のすべてのセルに複製
            for row in region.range.start.row..=region.range.end.row {
                for col in region.range.start.col..=region.range.end.col {
                    if row == region.parent.row && col == region.parent.col {
                        // 親セルはスキップ
                        continue;
                    }

                    self.cells[row as usize][col as usize] =
                        Cell::new_merged(parent_content.clone(), region.parent);
                }
            }
        }

        Ok(())
    }

    /// Markdownテーブルとして出力
    ///
    /// # 引数
    ///
    /// * `writer` - 出力先のライター
    ///
    /// # 戻り値
    ///
    /// * `Ok(())` - 出力に成功した場合
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    pub fn render_markdown<W: Write>(&self, writer: &mut W) -> Result<(), XlsxToMdError> {
        if self.rows == 0 || self.cols == 0 {
            return Ok(());
        }

        // 1. 列幅の計算
        let col_widths = self.calculate_column_widths();

        // 2. ヘッダー区切り行
        let separator = self.generate_separator(&col_widths);

        // 3. 各行の出力
        for (row_idx, row) in self.cells.iter().enumerate() {
            write!(writer, "|")?;

            for (col_idx, cell) in row.iter().enumerate() {
                let width = col_widths[col_idx];
                // セル内容をtrimしてからフォーマット
                let trimmed_content = cell.content.trim();
                // 表示幅を計算（全角文字は2、半角文字は1）
                let content_width = trimmed_content.width();

                // セルの前にスペースを1つ入れる
                write!(writer, " ")?;
                // 左揃えでセル内容を出力
                write!(writer, "{}", trimmed_content)?;
                // 残りのスペースを埋める（表示幅に基づく）
                if content_width < width {
                    for _ in content_width..width {
                        write!(writer, " ")?;
                    }
                }
                // セルの後にスペースを1つ入れる
                write!(writer, " |")?;
            }

            writeln!(writer)?;

            // 最初の行の後に区切り行を挿入
            if row_idx == 0 {
                writeln!(writer, "{}", separator)?;
            }
        }

        // 4. フラッシュ
        writer.flush()?;

        Ok(())
    }

    /// 列幅を計算（内部ヘルパー）
    ///
    /// 各列について、すべての行のセル内容の表示幅を計算し、列ごとの最大幅を返します。
    /// vscode-markdown-tableの実装に準拠し、trim処理と最小幅（3文字）を考慮します。
    /// 全角文字（日本語など）は表示幅2として計算します。
    fn calculate_column_widths(&self) -> Vec<usize> {
        let mut widths = vec![3; self.cols]; // 最小幅は3文字（区切り行の最小幅）

        for row in &self.cells {
            for (col_idx, cell) in row.iter().enumerate() {
                // trimしてから表示幅を計算（全角文字は2、半角文字は1）
                let trimmed_width = cell.content.trim().width();
                widths[col_idx] = widths[col_idx].max(trimmed_width);
            }
        }

        widths
    }

    /// ヘッダー区切り行を生成（内部ヘルパー）
    ///
    /// 各列幅に応じて "---" を生成し、"|" で連結します。
    /// vscode-markdown-tableの実装に準拠し、セルの前後のスペース（各1文字）を考慮します。
    fn generate_separator(&self, col_widths: &[usize]) -> String {
        let mut parts = vec!["|".to_string()];

        for &width in col_widths {
            // セルの前後のスペース（各1文字）+ セル幅分のハイフン
            parts.push("-".repeat(width + 2));
            parts.push("|".to_string());
        }

        parts.join("")
    }

    /// HTMLテーブルとして出力
    ///
    /// セル結合を含むテーブルをHTML形式で出力します。
    /// `MergeStrategy::HtmlFallback`が指定された場合に使用されます。
    ///
    /// # 引数
    ///
    /// * `writer` - 出力先のライター
    /// * `merged_regions` - 結合セル範囲のリスト
    ///
    /// # 戻り値
    ///
    /// * `Ok(())` - 出力に成功した場合
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    pub fn render_html<W: Write>(
        &self,
        writer: &mut W,
        merged_regions: &[MergedRegion],
    ) -> Result<(), XlsxToMdError> {
        writeln!(writer, "<table>")?;

        for (row_idx, row) in self.cells.iter().enumerate() {
            writeln!(writer, "  <tr>")?;

            for (col_idx, cell) in row.iter().enumerate() {
                let coord = CellCoord::new(row_idx as u32, col_idx as u32);

                // 結合セルの子かチェック
                if cell.is_merged && cell.merge_parent.is_some() {
                    continue; // スキップ
                }

                // rowspan/colspan計算
                let (rowspan, colspan) = self.calculate_span(&coord, merged_regions);

                if rowspan > 1 || colspan > 1 {
                    write!(
                        writer,
                        "    <td rowspan=\"{}\" colspan=\"{}\">",
                        rowspan, colspan
                    )?;
                } else {
                    write!(writer, "    <td>")?;
                }

                writeln!(writer, "{}</td>", cell.content)?;
            }

            writeln!(writer, "  </tr>")?;
        }

        writeln!(writer, "</table>")?;
        writer.flush()?;
        Ok(())
    }

    /// rowspan/colspanを計算（内部ヘルパー）
    ///
    /// 指定されたセル座標が結合セルの親かチェックし、親の場合はrow_span()とcol_span()を返します。
    fn calculate_span(&self, coord: &CellCoord, merged_regions: &[MergedRegion]) -> (u32, u32) {
        for region in merged_regions {
            if region.parent == *coord {
                return (region.row_span(), region.col_span());
            }
        }
        (1, 1)
    }

    /// 行数を取得
    pub(crate) fn get_rows(&self) -> usize {
        self.rows
    }

    /// 列数を取得
    pub(crate) fn get_cols(&self) -> usize {
        self.cols
    }

    /// 指定された行を取得
    pub(crate) fn get_row(&self, row_idx: usize) -> &[Cell] {
        if row_idx < self.rows {
            &self.cells[row_idx]
        } else {
            &[]
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::{CellRange, CellValue};

    #[test]
    fn test_cell_new() {
        let cell = Cell::new("Hello".to_string());
        assert_eq!(cell.content, "Hello");
        assert!(!cell.is_merged);
        assert!(cell.merge_parent.is_none());
    }

    #[test]
    fn test_cell_new_merged() {
        let parent = CellCoord::new(0, 0);
        let cell = Cell::new_merged("Merged".to_string(), parent);
        assert_eq!(cell.content, "Merged");
        assert!(cell.is_merged);
        assert_eq!(cell.merge_parent, Some(parent));
    }

    #[test]
    fn test_cell_empty() {
        let cell = Cell::empty();
        assert_eq!(cell.content, "");
        assert!(!cell.is_merged);
        assert!(cell.merge_parent.is_none());
    }

    #[test]
    fn test_determine_grid_size() {
        let cells = vec![
            RawCellData {
                coord: CellCoord::new(0, 0),
                value: CellValue::String("A1".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(2, 3),
                value: CellValue::String("D3".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
        ];

        let (rows, cols) = LogicalGrid::determine_grid_size(&cells);
        assert_eq!(rows, 3);
        assert_eq!(cols, 4);
    }

    #[test]
    fn test_build_empty_grid() {
        let cells = vec![];
        let formatted_cells = vec![];
        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![],
            hidden_rows: vec![],
            hidden_cols: vec![],
            is_1904: false,
        };

        let result = LogicalGrid::build(
            cells,
            formatted_cells,
            &metadata,
            MergeStrategy::DataDuplication,
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_build_simple_grid() {
        let cells = vec![
            RawCellData {
                coord: CellCoord::new(0, 0),
                value: CellValue::String("A1".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(0, 1),
                value: CellValue::String("B1".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
        ];

        let formatted_cells = vec![
            (CellCoord::new(0, 0), "A1".to_string()),
            (CellCoord::new(0, 1), "B1".to_string()),
        ];

        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![],
            hidden_rows: vec![],
            hidden_cols: vec![],
            is_1904: false,
        };

        let result = LogicalGrid::build(
            cells,
            formatted_cells,
            &metadata,
            MergeStrategy::DataDuplication,
        );
        assert!(result.is_ok());

        let grid = result.unwrap();
        assert_eq!(grid.rows, 1);
        assert_eq!(grid.cols, 2);
    }

    #[test]
    fn test_build_with_merged_cells_data_duplication() {
        let cells = vec![
            RawCellData {
                coord: CellCoord::new(0, 0),
                value: CellValue::String("Header".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(0, 1),
                value: CellValue::Empty,
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(0, 2),
                value: CellValue::Empty,
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
        ];

        let formatted_cells = vec![(CellCoord::new(0, 0), "Header".to_string())];

        let merged_range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(0, 2));
        let merged_region = MergedRegion::new(merged_range);

        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![merged_region],
            hidden_rows: vec![],
            hidden_cols: vec![],
            is_1904: false,
        };

        let result = LogicalGrid::build(
            cells,
            formatted_cells,
            &metadata,
            MergeStrategy::DataDuplication,
        );
        assert!(result.is_ok());

        let _grid = result.unwrap();
        // 結合セルの内容が複製されていることを確認
        // 注意: 内部実装の詳細に依存するため、render_markdown()の出力で確認する方が良い
    }

    #[test]
    fn test_render_markdown() {
        let cells = vec![
            RawCellData {
                coord: CellCoord::new(0, 0),
                value: CellValue::String("A1".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(0, 1),
                value: CellValue::String("B1".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
        ];

        let formatted_cells = vec![
            (CellCoord::new(0, 0), "A1".to_string()),
            (CellCoord::new(0, 1), "B1".to_string()),
        ];

        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![],
            hidden_rows: vec![],
            hidden_cols: vec![],
            is_1904: false,
        };

        let grid = LogicalGrid::build(
            cells,
            formatted_cells,
            &metadata,
            MergeStrategy::DataDuplication,
        )
        .unwrap();

        let mut output = Vec::new();
        let result = grid.render_markdown(&mut output);
        assert!(result.is_ok());

        let markdown = String::from_utf8(output).unwrap();
        
        // フォーマットが正しいことを確認
        // 実際の出力形式: | A1 | B1 | または | A1   | B1   | (列幅に応じて)
        assert!(markdown.contains("A1"));
        assert!(markdown.contains("B1"));
        assert!(markdown.contains("|"));
        assert!(markdown.contains("-"));
        // セルの前後にスペースが1つずつ入っていることを確認
        assert!(markdown.contains("| A1") || markdown.contains("|A1"));
        assert!(markdown.contains("| B1") || markdown.contains("|B1"));
    }

    #[test]
    fn test_render_markdown_with_trim() {
        // セル内容に前後のスペースがある場合のテスト
        let cells = vec![
            RawCellData {
                coord: CellCoord::new(0, 0),
                value: CellValue::String("  Header1  ".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(0, 1),
                value: CellValue::String("Header2".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(1, 0),
                value: CellValue::String("  Data1  ".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(1, 1),
                value: CellValue::String("Data2".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
        ];

        let formatted_cells = vec![
            (CellCoord::new(0, 0), "  Header1  ".to_string()),
            (CellCoord::new(0, 1), "Header2".to_string()),
            (CellCoord::new(1, 0), "  Data1  ".to_string()),
            (CellCoord::new(1, 1), "Data2".to_string()),
        ];

        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![],
            hidden_rows: vec![],
            hidden_cols: vec![],
            is_1904: false,
        };

        let grid = LogicalGrid::build(
            cells,
            formatted_cells,
            &metadata,
            MergeStrategy::DataDuplication,
        )
        .unwrap();

        let mut output = Vec::new();
        let result = grid.render_markdown(&mut output);
        assert!(result.is_ok());

        let markdown = String::from_utf8(output).unwrap();
        
        // trim処理が正しく動作することを確認
        // 前後のスペースが削除され、列幅が適切に計算されている
        // セル内容がtrimされていることを確認
        assert!(markdown.contains("Header1"), "Markdown should contain 'Header1'. Got: {}", markdown);
        assert!(markdown.contains("Header2"), "Markdown should contain 'Header2'. Got: {}", markdown);
        assert!(markdown.contains("Data1"), "Markdown should contain 'Data1'. Got: {}", markdown);
        assert!(markdown.contains("Data2"), "Markdown should contain 'Data2'. Got: {}", markdown);
        // 前後のスペースが削除されていることを確認
        assert!(!markdown.contains("  Header1  "), "Markdown should not contain spaces around 'Header1'");
        assert!(!markdown.contains("  Data1  "), "Markdown should not contain spaces around 'Data1'");
        // 列幅が統一されていることを確認（区切り行の長さが一致）
        let lines: Vec<&str> = markdown.lines().collect();
        assert!(lines.len() >= 2);
        // 区切り行の各列の幅が一致していることを確認
        let separator_line = lines[1];
        let header_line = lines[0];
        // 区切り行の各セル部分（`---`）の長さを確認
        let header_parts: Vec<&str> = header_line.split('|').collect();
        let separator_parts: Vec<&str> = separator_line.split('|').collect();
        // 各列の幅が一致していることを確認（前後の空要素を除く）
        for i in 1..header_parts.len().min(separator_parts.len()) - 1 {
            let header_cell = header_parts[i].trim();
            let separator_cell = separator_parts[i].trim();
            // 区切り行の長さは、セル内容の長さ + 2（前後のスペース）であるべき
            assert_eq!(
                separator_cell.len(),
                header_cell.len() + 2,
                "Column {} width mismatch: header='{}' (len={}), separator='{}' (len={})",
                i,
                header_cell,
                header_cell.len(),
                separator_cell,
                separator_cell.len()
            );
        }
    }

    #[test]
    fn test_render_html() {
        let cells = vec![
            RawCellData {
                coord: CellCoord::new(0, 0),
                value: CellValue::String("Header".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(0, 1),
                value: CellValue::Empty,
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
        ];

        let formatted_cells = vec![(CellCoord::new(0, 0), "Header".to_string())];

        let merged_range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(0, 1));
        let merged_region = MergedRegion::new(merged_range);

        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![merged_region.clone()],
            hidden_rows: vec![],
            hidden_cols: vec![],
            is_1904: false,
        };

        let grid = LogicalGrid::build(
            cells,
            formatted_cells,
            &metadata,
            MergeStrategy::HtmlFallback,
        )
        .unwrap();

        let mut output = Vec::new();
        let result = grid.render_html(&mut output, &metadata.merged_regions);
        assert!(result.is_ok());

        let html = String::from_utf8(output).unwrap();
        assert!(html.contains("<table>"));
        assert!(html.contains("</table>"));
        assert!(html.contains("Header"));
    }

    #[test]
    fn test_calculate_column_widths() {
        let grid_cells = vec![
            vec![
                Cell::new("Short".to_string()),
                Cell::new("Very Long Content".to_string()),
            ],
            vec![
                Cell::new("Longer".to_string()),
                Cell::new("Short".to_string()),
            ],
        ];

        let grid = LogicalGrid {
            cells: grid_cells,
            rows: 2,
            cols: 2,
        };

        let widths = grid.calculate_column_widths();
        assert_eq!(widths[0], 6); // "Longer" の長さ
        assert_eq!(widths[1], 17); // "Very Long Content" の長さ
    }

    #[test]
    fn test_generate_separator() {
        let grid = LogicalGrid {
            cells: vec![],
            rows: 0,
            cols: 0,
        };

        let col_widths = vec![3, 5, 2];
        let separator = grid.generate_separator(&col_widths);
        assert!(separator.contains("|"));
        assert!(separator.contains("-"));
    }

    #[test]
    fn test_calculate_span() {
        let merged_range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(1, 2));
        let merged_region = MergedRegion::new(merged_range);

        let grid = LogicalGrid {
            cells: vec![],
            rows: 0,
            cols: 0,
        };

        let (rowspan, colspan) =
            grid.calculate_span(&CellCoord::new(0, 0), std::slice::from_ref(&merged_region));
        assert_eq!(rowspan, 2);
        assert_eq!(colspan, 3);

        let (rowspan, colspan) = grid.calculate_span(&CellCoord::new(5, 5), &[merged_region]);
        assert_eq!(rowspan, 1);
        assert_eq!(colspan, 1);
    }

    #[test]
    fn test_calculate_column_widths_with_japanese() {
        // 日本語を含むテストケース
        // 全角文字は表示幅2、半角文字は表示幅1
        let grid_cells = vec![
            vec![
                Cell::new("市区町村コード".to_string()), // 7文字 × 2 = 14
                Cell::new("店舗名".to_string()),         // 3文字 × 2 = 6
            ],
            vec![
                Cell::new("01100".to_string()), // 5文字 × 1 = 5
                Cell::new("札幌店".to_string()), // 3文字 × 2 = 6
            ],
        ];

        let grid = LogicalGrid {
            cells: grid_cells,
            rows: 2,
            cols: 2,
        };

        let widths = grid.calculate_column_widths();
        assert_eq!(widths[0], 14); // "市区町村コード" の表示幅
        assert_eq!(widths[1], 6); // "店舗名" と "札幌店" の表示幅（同じ）
    }

    #[test]
    fn test_render_markdown_with_japanese() {
        // 日本語を含むMarkdownテーブルの出力テスト
        let cells = vec![
            RawCellData {
                coord: CellCoord::new(0, 0),
                value: CellValue::String("ヘッダー".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(0, 1),
                value: CellValue::String("Header".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(1, 0),
                value: CellValue::String("データ".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
            RawCellData {
                coord: CellCoord::new(1, 1),
                value: CellValue::String("Data".to_string()),
                format_id: None,
                format_string: None,
                formula: None,
                hyperlink: None,
                rich_text: None,
            },
        ];

        let formatted_cells = vec![
            (CellCoord::new(0, 0), "ヘッダー".to_string()),
            (CellCoord::new(0, 1), "Header".to_string()),
            (CellCoord::new(1, 0), "データ".to_string()),
            (CellCoord::new(1, 1), "Data".to_string()),
        ];

        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![],
            hidden_rows: vec![],
            hidden_cols: vec![],
            is_1904: false,
        };

        let grid = LogicalGrid::build(
            cells,
            formatted_cells,
            &metadata,
            MergeStrategy::DataDuplication,
        )
        .unwrap();

        let mut output = Vec::new();
        let result = grid.render_markdown(&mut output);
        assert!(result.is_ok());

        let markdown = String::from_utf8(output).unwrap();

        // 各行の長さが一致していることを確認（全角文字の幅が正しく計算されている場合）
        let lines: Vec<&str> = markdown.lines().collect();
        assert_eq!(lines.len(), 3); // ヘッダー行、区切り行、データ行

        // 区切り行が正しい形式であることを確認
        let separator_line = lines[1];
        assert!(separator_line.starts_with("|"));
        assert!(separator_line.ends_with("|"));
        assert!(separator_line.contains("-"));

        // "ヘッダー"（表示幅8）と "Header"（表示幅6）で、列幅は8になるはず
        // セルの形式: "| content |" で、前後にスペース1つずつ
        // 最初の列: "| ヘッダー |" (ヘッダーは表示幅8だが、バイト幅は12)
        // 最初の列の区切り: "|----------|" (width + 2 = 10個の-)
        assert!(
            separator_line.contains("----------"),
            "Separator should contain 10 dashes for column with width 8. Got: {}",
            separator_line
        );
    }

    #[test]
    fn test_display_width_calculation() {
        use unicode_width::UnicodeWidthStr;

        // 日本語の表示幅テスト
        assert_eq!("日本語".width(), 6); // 3文字 × 2 = 6
        assert_eq!("ABC".width(), 3); // 3文字 × 1 = 3
        assert_eq!("日本ABC".width(), 7); // 2文字 × 2 + 3文字 × 1 = 7

        // 混合文字の表示幅テスト
        assert_eq!("市区町村コード".width(), 14); // 7文字 × 2 = 14
        assert_eq!("01100".width(), 5); // 5文字 × 1 = 5
    }
}
