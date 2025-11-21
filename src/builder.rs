//! Builder Module
//!
//! Fluent Builder APIを提供し、`Converter`インスタンスを段階的に構築する。

use crate::api::{DateFormat, FormulaMode, MergeStrategy, OutputFormat, SheetSelector};
use crate::error::XlsxToMdError;
use crate::types::CellRange;
use chrono::NaiveDate;
use rayon::prelude::*;
use std::io::{Cursor, Read, Seek, Write};

/// 変換処理の設定を保持する内部構造体
#[derive(Debug, Clone)]
pub(crate) struct ConversionConfig {
    /// シート選択方式
    pub sheet_selector: SheetSelector,

    /// セル結合戦略
    pub merge_strategy: MergeStrategy,

    /// 日付形式
    pub date_format: DateFormat,

    /// 数式出力モード
    pub formula_mode: FormulaMode,

    /// 非表示要素を含めるか
    pub include_hidden: bool,

    /// セル範囲制限（Option: Noneの場合は全範囲）
    pub range: Option<CellRange>,

    /// 出力フォーマット
    pub output_format: OutputFormat,
}

impl Default for ConversionConfig {
    fn default() -> Self {
        Self {
            sheet_selector: SheetSelector::All,
            merge_strategy: MergeStrategy::DataDuplication,
            date_format: DateFormat::Iso8601,
            formula_mode: FormulaMode::CachedValue,
            include_hidden: false,
            range: None,
            output_format: OutputFormat::Markdown,
        }
    }
}

/// Fluent Builder APIを提供する構造体
///
/// `Converter`インスタンスを段階的に構築するためのビルダーです。
/// すべての設定項目にデフォルト値が設定されており、必要な設定のみをオーバーライドできます。
///
/// # 使用例
///
/// ```rust,no_run
/// use xlsxzero::{ConverterBuilder, SheetSelector, MergeStrategy};
///
/// # fn main() -> Result<(), xlsxzero::XlsxToMdError> {
/// let converter = ConverterBuilder::new()
///     .with_sheet_selector(SheetSelector::Index(0))
///     .with_merge_strategy(MergeStrategy::HtmlFallback)
///     .build()?;
/// # Ok(())
/// # }
/// ```
#[derive(Debug)]
pub struct ConverterBuilder {
    /// 内部設定（構築中）
    config: ConversionConfig,
}

impl Default for ConverterBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl ConverterBuilder {
    /// デフォルト設定を持つビルダーインスタンスを生成する
    ///
    /// # デフォルト設定
    ///
    /// - シート選択: すべてのシート
    /// - セル結合戦略: データ重複フィル
    /// - 日付形式: ISO 8601 (YYYY-MM-DD)
    /// - 非表示要素: スキップ
    /// - 数式モード: キャッシュ値を出力
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::ConverterBuilder;
    ///
    /// let builder = ConverterBuilder::new();
    /// ```
    pub fn new() -> Self {
        Self {
            config: ConversionConfig::default(),
        }
    }

    /// 変換対象のシートを選択する
    ///
    /// # 引数
    ///
    /// * `selector: SheetSelector`: シート選択方式
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::{ConverterBuilder, SheetSelector};
    ///
    /// // 単一シートをインデックスで指定
    /// let builder = ConverterBuilder::new()
    ///     .with_sheet_selector(SheetSelector::Index(0));
    ///
    /// // 単一シートを名前で指定
    /// let builder = ConverterBuilder::new()
    ///     .with_sheet_selector(SheetSelector::Name("Sheet1".to_string()));
    ///
    /// // 複数シートを指定
    /// let builder = ConverterBuilder::new()
    ///     .with_sheet_selector(SheetSelector::Indices(vec![0, 2, 4]));
    /// ```
    pub fn with_sheet_selector(mut self, selector: SheetSelector) -> Self {
        self.config.sheet_selector = selector;
        self
    }

    /// セル結合の処理戦略を指定する
    ///
    /// # 引数
    ///
    /// * `strategy: MergeStrategy`: セル結合戦略
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::{ConverterBuilder, MergeStrategy};
    ///
    /// let builder = ConverterBuilder::new()
    ///     .with_merge_strategy(MergeStrategy::HtmlFallback);
    /// ```
    pub fn with_merge_strategy(mut self, strategy: MergeStrategy) -> Self {
        self.config.merge_strategy = strategy;
        self
    }

    /// 日付の出力形式を指定する
    ///
    /// # 引数
    ///
    /// * `format: DateFormat`: 日付形式
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::{ConverterBuilder, DateFormat};
    ///
    /// // ISO 8601形式（デフォルト）
    /// let builder = ConverterBuilder::new()
    ///     .with_date_format(DateFormat::Iso8601);
    ///
    /// // カスタム形式
    /// let builder = ConverterBuilder::new()
    ///     .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()));
    /// ```
    pub fn with_date_format(mut self, format: DateFormat) -> Self {
        self.config.date_format = format;
        self
    }

    /// 数式セルの出力モードを指定する
    ///
    /// # 引数
    ///
    /// * `mode: FormulaMode`: 数式出力モード
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::{ConverterBuilder, FormulaMode};
    ///
    /// let builder = ConverterBuilder::new()
    ///     .with_formula_mode(FormulaMode::Formula);
    /// ```
    pub fn with_formula_mode(mut self, mode: FormulaMode) -> Self {
        self.config.formula_mode = mode;
        self
    }

    /// 非表示要素（非表示シート、行、列）を出力に含めるかを指定する
    ///
    /// # 引数
    ///
    /// * `include: bool`:
    ///   * `true`: 非表示要素を含める
    ///   * `false`: 非表示要素をスキップ（デフォルト）
    ///
    /// # Phase I制限事項
    ///
    /// Phase Iでは `calamine` ライブラリの制限により、非表示行・非表示列の情報を取得できません。
    /// そのため、`include_hidden(false)` を指定しても**非表示行・列のフィルタリングは機能しません**。
    /// 非表示シートのみがフィルタリング対象となります。
    /// Phase IIで `xl/worksheets/sheet*.xml` から `hidden="1"` 属性を直接パースすることで完全対応予定です。
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::ConverterBuilder;
    ///
    /// let builder = ConverterBuilder::new()
    ///     .include_hidden(true);
    /// ```
    pub fn include_hidden(mut self, include: bool) -> Self {
        self.config.include_hidden = include;
        self
    }

    /// 処理対象のセル範囲を制限する
    ///
    /// 範囲外のセルは無視されます。
    ///
    /// # 引数
    ///
    /// * `start: (u32, u32)`: 開始セル座標 (row, col)（0始まり）
    /// * `end: (u32, u32)`: 終了セル座標 (row, col)（0始まり）
    ///
    /// # 制約
    ///
    /// * `start.0 <= end.0` かつ `start.1 <= end.1` でなければならない
    /// * 制約違反の場合、`build()`時に`XlsxToMdError::Config`を返す
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::ConverterBuilder;
    ///
    /// // A1:C10の範囲を処理（0始まりなので、row 0-9, col 0-2）
    /// let builder = ConverterBuilder::new()
    ///     .with_range((0, 0), (9, 2));
    /// ```
    pub fn with_range(mut self, start: (u32, u32), end: (u32, u32)) -> Self {
        use crate::types::CellCoord;
        self.config.range = Some(CellRange::new(
            CellCoord::new(start.0, start.1),
            CellCoord::new(end.0, end.1),
        ));
        self
    }

    /// 出力フォーマットを指定する
    ///
    /// # 引数
    ///
    /// * `format: OutputFormat`: 出力フォーマット（Markdown, HTML, JSON, CSV）
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::{ConverterBuilder, OutputFormat};
    ///
    /// // JSON形式で出力
    /// let builder = ConverterBuilder::new()
    ///     .with_output_format(OutputFormat::Json);
    ///
    /// // CSV形式で出力
    /// let builder = ConverterBuilder::new()
    ///     .with_output_format(OutputFormat::Csv);
    /// ```
    pub fn with_output_format(mut self, format: OutputFormat) -> Self {
        self.config.output_format = format;
        self
    }

    /// 設定を検証し、`Converter`インスタンスを生成する
    ///
    /// # 戻り値
    ///
    /// * `Ok(Converter)`: 設定が有効な場合、Converterインスタンス
    /// * `Err(XlsxToMdError::Config)`: 設定が無効な場合（例: 範囲指定の開始 > 終了）
    ///
    /// # 発生し得るエラー
    ///
    /// * `XlsxToMdError::Config(String)`: 設定の検証に失敗した場合
    ///   * 範囲指定の開始座標が終了座標より大きい
    ///   * カスタム日付形式が不正な書式文字列
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use xlsxzero::{ConverterBuilder, SheetSelector};
    ///
    /// # fn main() -> Result<(), xlsxzero::XlsxToMdError> {
    /// let converter = ConverterBuilder::new()
    ///     .with_sheet_selector(SheetSelector::Index(0))
    ///     .build()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn build(self) -> Result<Converter, XlsxToMdError> {
        // 1. セル範囲の検証
        if let Some(range) = &self.config.range {
            if range.start.row > range.end.row {
                return Err(XlsxToMdError::Config(format!(
                    "Invalid range: start row ({}) > end row ({})",
                    range.start.row, range.end.row
                )));
            }

            if range.start.col > range.end.col {
                return Err(XlsxToMdError::Config(format!(
                    "Invalid range: start col ({}) > end col ({})",
                    range.start.col, range.end.col
                )));
            }
        }

        // 2. カスタム日付形式の検証
        if let DateFormat::Custom(ref format_str) = self.config.date_format {
            // テスト用の日付でフォーマット試行
            let test_date = NaiveDate::from_ymd_opt(2025, 1, 1)
                .ok_or_else(|| XlsxToMdError::Config("Failed to create test date".to_string()))?;
            let formatted = test_date.format(format_str).to_string();
            if formatted.is_empty() {
                return Err(XlsxToMdError::Config(format!(
                    "Invalid date format string: '{}'",
                    format_str
                )));
            }
        }

        // 3. Converterインスタンス生成
        Ok(Converter::new(self.config))
    }
}

/// 変換処理のファサード
///
/// ExcelファイルをMarkdown形式に変換するためのメインエントリーポイントです。
/// `ConverterBuilder`を使用して構築された設定に基づいて変換処理を実行します。
///
/// # 使用例
///
/// ```rust,no_run
/// use xlsxzero::ConverterBuilder;
/// use std::fs::File;
///
/// # fn main() -> Result<(), xlsxzero::XlsxToMdError> {
/// let converter = ConverterBuilder::new().build()?;
/// let input = File::open("example.xlsx")?;
/// let mut output = Vec::new();
/// converter.convert(input, &mut output)?;
/// # Ok(())
/// # }
/// ```
#[derive(Debug)]
pub struct Converter {
    /// 変換設定
    config: ConversionConfig,

    /// セルフォーマッター
    formatter: crate::formatter::CellFormatter,
}

impl Converter {
    pub(crate) fn new(config: ConversionConfig) -> Self {
        Self {
            formatter: crate::formatter::CellFormatter::new(),
            config,
        }
    }

    /// ExcelファイルをMarkdown形式に変換
    ///
    /// # 引数
    ///
    /// * `input` - Excelファイルを読み込むためのリーダー（Read + Seekトレイトを実装）
    /// * `output` - Markdown出力先のライター（Writeトレイトを実装）
    ///
    /// # 戻り値
    ///
    /// * `Ok(())` - 変換に成功した場合
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    ///
    /// # 処理フロー
    ///
    /// 1. WorkbookParserの初期化
    /// 2. シート選択
    /// 3. 各シートについて処理（ループ）
    ///    - シートのパース
    ///    - セルのフォーマット
    ///    - グリッドの構築
    ///    - Markdown/HTML出力（戦略に応じて）
    /// 4. 出力バッファをフラッシュ
    ///
    /// # 使用例
    ///
    /// ## ファイルからファイルへの変換
    ///
    /// ```rust,no_run
    /// use xlsxzero::ConverterBuilder;
    /// use std::fs::File;
    ///
    /// # fn main() -> Result<(), xlsxzero::XlsxToMdError> {
    /// let converter = ConverterBuilder::new().build()?;
    /// let input = File::open("example.xlsx")?;
    /// let output = File::create("output.md")?;
    /// converter.convert(input, output)?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// 注意: `File`は`Clone`を実装していないため、メモリバッファを使用する場合は`Cursor`を使用してください。
    ///
    /// ## メモリバッファからの変換
    ///
    /// ```rust,no_run
    /// use xlsxzero::ConverterBuilder;
    /// use std::io::Cursor;
    ///
    /// # fn main() -> Result<(), xlsxzero::XlsxToMdError> {
    /// let converter = ConverterBuilder::new().build()?;
    /// let excel_data: Vec<u8> = vec![]; // Excelファイルのバイト列
    /// let mut markdown_output = Vec::new();
    /// converter.convert(Cursor::new(excel_data), &mut markdown_output)?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// ## 標準出力への変換
    ///
    /// ```rust,no_run
    /// use xlsxzero::ConverterBuilder;
    /// use std::fs::File;
    ///
    /// # fn main() -> Result<(), xlsxzero::XlsxToMdError> {
    /// let converter = ConverterBuilder::new().build()?;
    /// let input = File::open("example.xlsx")?;
    /// converter.convert(input, std::io::stdout())?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// 注意: `File`は`Clone`を実装していないため、実際の使用では`File`を直接使用できますが、
    /// メモリバッファを使用する場合は`Cursor`を使用してください。
    pub fn convert<R: Read + Seek, W: Write>(
        &self,
        mut input: R,
        mut output: W,
    ) -> Result<(), XlsxToMdError> {
        use std::io::{BufWriter, Write};

        // 1. 入力データをメモリに読み込む（並列処理のため）
        use crate::security::SecurityConfig;
        let security_config = SecurityConfig::default();
        let mut buffer = Vec::new();
        let bytes_read = input.read_to_end(&mut buffer)?;

        if bytes_read as u64 > security_config.max_input_file_size {
            return Err(XlsxToMdError::SecurityViolation(format!(
                "Input file size exceeds maximum: {} bytes (max: {} bytes)",
                bytes_read, security_config.max_input_file_size
            )));
        }

        // 2. WorkbookParserの初期化（Phase II: メタデータも取得）
        // 非表示行・列の情報を取得するため、open_with_metadata()を使用
        let parser = crate::parser::WorkbookParser::open_with_metadata(Cursor::new(buffer.clone()))?;

        // 3. シート選択
        let sheet_names =
            parser.select_sheets(&self.config.sheet_selector, self.config.include_hidden)?;

        // 4. メタデータを1回だけ解析して再利用（並列処理の効率化）
        // メタデータを抽出（WorkbookParserから取得）
        let metadata = parser.metadata()
            .ok_or_else(|| XlsxToMdError::Config("Metadata not available".to_string()))?
            .clone();

        // 5. 各シートの処理を並列化
        // 各シートの処理結果（出力文字列）を並列に計算
        let sheet_outputs: Result<Vec<(usize, String)>, XlsxToMdError> = sheet_names
            .par_iter()
            .enumerate()
            .map(|(sheet_idx, sheet_name)| {
                // 各シート処理でワークブックを再オープン（メモリ内のデータを使用）
                // メタデータは既に解析済みなので再利用
                let mut parser = crate::parser::WorkbookParser::open_with_existing_metadata(
                    Cursor::new(buffer.clone()),
                    metadata.clone(),
                )?;

                // シートのパース
                let (metadata, raw_cells) = parser.parse_sheet(sheet_name, &self.config)?;

                // セルのフォーマット
                let mut formatted_cells = Vec::new();
                for raw_cell in &raw_cells {
                    let content =
                        self.formatter
                            .format_cell(raw_cell, &self.config, metadata.is_1904)?;
                    formatted_cells.push((raw_cell.coord, content));
                }

                // グリッドの構築
                let grid = crate::grid::LogicalGrid::build(
                    raw_cells,
                    formatted_cells,
                    &metadata,
                    self.config.merge_strategy,
                )?;

                // 出力フォーマッターを取得
                let formatter = crate::output::OutputFormatter::from_format(self.config.output_format);

                // 出力フォーマットに応じて出力
                let mut output_buffer = Vec::new();
                formatter.render(&grid, &mut output_buffer, &metadata.merged_regions)?;

                let output_string = String::from_utf8(output_buffer).map_err(|e| {
                    XlsxToMdError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
                })?;

                Ok((sheet_idx, output_string))
            })
            .collect();

        let mut sheet_outputs = sheet_outputs?;

        // 結果をインデックス順にソート（並列処理の順序を保証）
        sheet_outputs.sort_by_key(|(idx, _)| *idx);

        // 5. 結果を順序付きで出力
        let mut writer = BufWriter::new(&mut output);
        for (sheet_idx, (_, sheet_output)) in sheet_outputs.iter().enumerate() {
            // シート間の区切り（Markdown形式の場合のみ）
            if sheet_idx > 0 && self.config.output_format == crate::api::OutputFormat::Markdown {
                writeln!(writer, "\n---\n")?;
            } else if sheet_idx > 0 {
                // 他のフォーマットの場合は改行のみ
                writeln!(writer)?;
            }

            // シート名をヘッダーとして出力（Markdown形式の場合のみ）
            if self.config.output_format == crate::api::OutputFormat::Markdown {
                writeln!(writer, "# {}\n", sheet_names[sheet_idx])?;
            } else if self.config.output_format == crate::api::OutputFormat::Json {
                // JSON形式の場合は、シート名を含める（既にformatterで処理済みの場合はスキップ）
                // ここでは既にJSONが生成されているので、そのまま出力
            } else {
                // CSV/HTML形式の場合は、シート名をコメントとして出力
                if self.config.output_format == crate::api::OutputFormat::Csv {
                    writeln!(writer, "# Sheet: {}\n", sheet_names[sheet_idx])?;
                } else if self.config.output_format == crate::api::OutputFormat::Html {
                    writeln!(writer, "<!-- Sheet: {} -->\n", sheet_names[sheet_idx])?;
                }
            }

            // シートの出力
            write!(writer, "{}", sheet_output)?;
        }

        // 6. フラッシュ
        writer.flush()?;

        Ok(())
    }

    /// ExcelファイルをMarkdown形式の文字列に変換
    ///
    /// # 引数
    ///
    /// * `input` - Excelファイルを読み込むためのリーダー（Read + Seekトレイトを実装）
    ///
    /// # 戻り値
    ///
    /// * `Ok(String)` - 変換されたMarkdown文字列
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    ///
    /// # 使用例
    ///
    /// ```rust,no_run
    /// use std::fs::File;
    /// use xlsxzero::ConverterBuilder;
    ///
    /// # fn main() -> Result<(), xlsxzero::XlsxToMdError> {
    /// let converter = ConverterBuilder::new().build()?;
    /// let input = File::open("example.xlsx")?;
    /// let markdown = converter.convert_to_string(input)?;
    /// println!("{}", markdown);
    /// # Ok(())
    /// # }
    /// ```
    pub fn convert_to_string<R: Read + Seek>(&self, input: R) -> Result<String, XlsxToMdError> {
        let mut buffer = Vec::new();
        self.convert(input, &mut buffer)?;

        let result = String::from_utf8(buffer).map_err(|e| {
            XlsxToMdError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
        })?;

        Ok(result)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::CellCoord;

    #[test]
    fn test_converter_builder_new() {
        let builder = ConverterBuilder::new();
        assert_eq!(builder.config.sheet_selector, SheetSelector::All);
        assert_eq!(
            builder.config.merge_strategy,
            MergeStrategy::DataDuplication
        );
        assert_eq!(builder.config.date_format, DateFormat::Iso8601);
        assert_eq!(builder.config.formula_mode, FormulaMode::CachedValue);
        assert!(!builder.config.include_hidden);
        assert!(builder.config.range.is_none());
    }

    #[test]
    fn test_with_sheet_selector() {
        let builder = ConverterBuilder::new().with_sheet_selector(SheetSelector::Index(0));
        assert!(matches!(
            builder.config.sheet_selector,
            SheetSelector::Index(0)
        ));

        let builder =
            ConverterBuilder::new().with_sheet_selector(SheetSelector::Name("Sheet1".to_string()));
        assert!(matches!(
            builder.config.sheet_selector,
            SheetSelector::Name(ref name) if name == "Sheet1"
        ));
    }

    #[test]
    fn test_with_merge_strategy() {
        let builder = ConverterBuilder::new().with_merge_strategy(MergeStrategy::HtmlFallback);
        assert_eq!(builder.config.merge_strategy, MergeStrategy::HtmlFallback);
    }

    #[test]
    fn test_with_date_format() {
        let builder = ConverterBuilder::new()
            .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()));
        assert!(matches!(
            builder.config.date_format,
            DateFormat::Custom(ref s) if s == "%Y年%m月%d日"
        ));
    }

    #[test]
    fn test_with_formula_mode() {
        let builder = ConverterBuilder::new().with_formula_mode(FormulaMode::Formula);
        assert_eq!(builder.config.formula_mode, FormulaMode::Formula);
    }

    #[test]
    fn test_include_hidden() {
        let builder = ConverterBuilder::new().include_hidden(true);
        assert!(builder.config.include_hidden);
    }

    #[test]
    fn test_with_range() {
        let builder = ConverterBuilder::new().with_range((0, 0), (9, 2));
        assert!(builder.config.range.is_some());
        let range = builder.config.range.unwrap();
        assert_eq!(range.start, CellCoord::new(0, 0));
        assert_eq!(range.end, CellCoord::new(9, 2));
    }

    #[test]
    fn test_build_success() {
        let result = ConverterBuilder::new().build();
        assert!(result.is_ok());
    }

    #[test]
    fn test_build_with_invalid_range_row() {
        let result = ConverterBuilder::new().with_range((10, 0), (0, 0)).build();
        assert!(result.is_err());
        match result {
            Err(XlsxToMdError::Config(msg)) => {
                assert!(msg.contains("start row"));
            }
            _ => panic!("Expected Config error"),
        }
    }

    #[test]
    fn test_build_with_invalid_range_col() {
        let result = ConverterBuilder::new().with_range((0, 10), (0, 0)).build();
        assert!(result.is_err());
        match result {
            Err(XlsxToMdError::Config(msg)) => {
                assert!(msg.contains("start col"));
            }
            _ => panic!("Expected Config error"),
        }
    }

    #[test]
    fn test_build_with_valid_custom_date_format() {
        let result = ConverterBuilder::new()
            .with_date_format(DateFormat::Custom("%Y-%m-%d".to_string()))
            .build();
        assert!(result.is_ok());
    }

    #[test]
    fn test_build_with_invalid_custom_date_format() {
        // 空のフォーマット文字列は無効
        let result = ConverterBuilder::new()
            .with_date_format(DateFormat::Custom("".to_string()))
            .build();
        assert!(result.is_err());
        match result {
            Err(XlsxToMdError::Config(msg)) => {
                assert!(msg.contains("Invalid date format"));
            }
            _ => panic!("Expected Config error"),
        }
    }

    #[test]
    fn test_builder_method_chaining() {
        let builder = ConverterBuilder::new()
            .with_sheet_selector(SheetSelector::Index(0))
            .with_merge_strategy(MergeStrategy::HtmlFallback)
            .with_date_format(DateFormat::Iso8601)
            .with_formula_mode(FormulaMode::Formula)
            .include_hidden(true)
            .with_range((0, 0), (10, 5));

        assert!(matches!(
            builder.config.sheet_selector,
            SheetSelector::Index(0)
        ));
        assert_eq!(builder.config.merge_strategy, MergeStrategy::HtmlFallback);
        assert_eq!(builder.config.date_format, DateFormat::Iso8601);
        assert_eq!(builder.config.formula_mode, FormulaMode::Formula);
        assert!(builder.config.include_hidden);
        assert!(builder.config.range.is_some());
    }

    #[test]
    fn test_build_with_all_settings() {
        let result = ConverterBuilder::new()
            .with_sheet_selector(SheetSelector::Name("Sheet1".to_string()))
            .with_merge_strategy(MergeStrategy::DataDuplication)
            .with_date_format(DateFormat::Custom("%Y/%m/%d".to_string()))
            .with_formula_mode(FormulaMode::CachedValue)
            .include_hidden(false)
            .with_range((0, 0), (99, 9))
            .build();

        assert!(result.is_ok());
    }

    // Converter構造体のテスト
    #[test]
    fn test_converter_new() {
        let _converter = ConverterBuilder::new().build().unwrap();
        // Converterが正常に構築されることを確認
        // 実際の変換処理は統合テストで検証（Issue #11）
    }

    #[test]
    fn test_converter_convert_to_string_with_invalid_input() {
        let converter = ConverterBuilder::new().build().unwrap();
        // 無効な入力データ（空のVec）
        let invalid_input: Vec<u8> = vec![];
        let result = converter.convert_to_string(std::io::Cursor::new(invalid_input));
        // エラーが返されることを確認
        assert!(result.is_err());
    }
}
