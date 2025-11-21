//! Public API Types
//!
//! 公開APIで使用する列挙型を定義するモジュール。

/// セル結合の処理戦略
///
/// Excelの結合セルをMarkdownに変換する際の処理方法を指定します。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum MergeStrategy {
    /// 結合セル範囲内のすべてのセルに親セルの値を複製
    ///
    /// 純粋なMarkdownテーブルとして出力します。結合セル範囲（例: A1:C1）を検出し、
    /// 親セル（A1）の値を結合範囲内のすべてのセルに複製します。
    ///
    /// # 利点
    ///
    /// - LLMが理解しやすい
    /// - トークン効率が高い
    /// - RAGシステムに最適
    ///
    /// # 出力例
    ///
    /// ```markdown
    /// | Header1 | Header1 | Header1 |
    /// | ------- | ------- | ------- |
    /// | Data1   | Data2   | Data3   |
    /// ```
    DataDuplication,

    /// HTMLテーブル（rowspan/colspan属性）として出力
    ///
    /// 結合セルを検出した場合、テーブル全体をHTMLとして出力します。
    /// `<td rowspan="...">`および`<td colspan="...">`属性を使用します。
    ///
    /// # 利点
    ///
    /// - 構造的忠実性を完全に維持
    /// - 視覚的な結合情報を保持
    ///
    /// # 出力例
    ///
    /// ```html
    /// <table>
    ///   <tr>
    ///     <th colspan="3">Header1</th>
    ///   </tr>
    ///   <tr>
    ///     <td>Data1</td>
    ///     <td>Data2</td>
    ///     <td>Data3</td>
    ///   </tr>
    /// </table>
    /// ```
    HtmlFallback,
}

/// 日付の出力形式
///
/// Excelの日付セルをMarkdownに変換する際の出力形式を指定します。
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum DateFormat {
    /// ISO 8601形式（YYYY-MM-DD）
    ///
    /// 例: `2025-11-20`
    Iso8601,

    /// カスタム形式（chrono互換フォーマット文字列）
    ///
    /// chrono互換のフォーマット文字列を使用して、カスタム日付形式を指定します。
    ///
    /// # フォーマット指定子（主要なもの）
    ///
    /// - `%Y`: 4桁の年（例: 2025）
    /// - `%y`: 2桁の年（例: 25）
    /// - `%m`: 2桁の月（01-12）
    /// - `%d`: 2桁の日（01-31）
    /// - `%H`: 24時間形式の時（00-23）
    /// - `%M`: 分（00-59）
    /// - `%S`: 秒（00-59）
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::{ConverterBuilder, DateFormat};
    ///
    /// # fn main() -> Result<(), xlsxzero::XlsxToMdError> {
    /// let converter = ConverterBuilder::new()
    ///     .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()))
    ///     .build()?;
    /// # Ok(())
    /// # }
    /// ```
    Custom(String),
}

/// 数式セルの出力モード
///
/// Excelの数式セルをMarkdownに変換する際の出力方法を指定します。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum FormulaMode {
    /// キャッシュされた結果値を出力（デフォルト）
    ///
    /// 数式セルの計算結果（キャッシュされた値）を出力します。
    /// 例: `=SUM(A1:A10)` → `100`
    CachedValue,

    /// 数式文字列を出力
    ///
    /// 数式そのものを文字列として出力します。
    /// 例: `=SUM(A1:A10)` → `=SUM(A1:A10)`
    Formula,
}

/// シート選択方式
///
/// 変換対象のシートを選択する方法を指定します。
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum SheetSelector {
    /// すべてのシートを変換（デフォルト）
    All,

    /// インデックス指定（0始まり）
    ///
    /// 例: `SheetSelector::Index(0)` は最初のシートを選択
    Index(usize),

    /// シート名指定
    ///
    /// 例: `SheetSelector::Name("Sheet1".to_string())`
    Name(String),

    /// 複数のインデックス指定
    ///
    /// 例: `SheetSelector::Indices(vec![0, 2, 4])`
    Indices(Vec<usize>),

    /// 複数のシート名指定
    ///
    /// 例: `SheetSelector::Names(vec!["Sheet1".to_string(), "Sheet2".to_string()])`
    Names(Vec<String>),
}

/// 出力フォーマット
///
/// Excelファイルを変換する際の出力形式を指定します。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum OutputFormat {
    /// Markdown形式（デフォルト）
    ///
    /// Markdownテーブル形式で出力します。
    /// セル結合は`MergeStrategy`に基づいて処理されます。
    ///
    /// # 出力例
    ///
    /// ```markdown
    /// | Header1 | Header2 |
    /// | ------- | ------- |
    /// | Data1   | Data2   |
    /// ```
    Markdown,

    /// HTML形式
    ///
    /// HTMLテーブル形式で出力します。
    /// セル結合は`rowspan`と`colspan`属性を使用します。
    ///
    /// # 出力例
    ///
    /// ```html
    /// <table>
    ///   <tr>
    ///     <td>Header1</td>
    ///     <td>Header2</td>
    ///   </tr>
    ///   <tr>
    ///     <td>Data1</td>
    ///     <td>Data2</td>
    ///   </tr>
    /// </table>
    /// ```
    Html,

    /// JSON形式
    ///
    /// セルデータをJSON形式で出力します。
    /// 各シートは配列として、各行はオブジェクトとして表現されます。
    ///
    /// # 出力例
    ///
    /// ```json
    /// {
    ///   "sheet_name": "Sheet1",
    ///   "rows": [
    ///     {"A": "Header1", "B": "Header2"},
    ///     {"A": "Data1", "B": "Data2"}
    ///   ]
    /// }
    /// ```
    Json,

    /// CSV形式
    ///
    /// CSV（Comma-Separated Values）形式で出力します。
    /// 各シートは独立したCSVとして出力されます。
    ///
    /// # 出力例
    ///
    /// ```csv
    /// Header1,Header2
    /// Data1,Data2
    /// ```
    Csv,
}
