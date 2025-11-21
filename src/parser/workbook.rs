//! Parser Module
//!
//! calamineを使用したExcelファイル解析の基礎実装。
//! ストリーミング処理により、メモリ効率的にセルデータを抽出します。

use calamine::{open_workbook_auto_from_rs, Data, Range, Reader, Sheets, Xlsx};
use std::io::{Cursor, Read, Seek};

use crate::api::SheetSelector;
use crate::builder::ConversionConfig;
use crate::error::XlsxToMdError;
use crate::parser::XlsxMetadataParser;
use crate::types::{CellCoord, CellRange, CellValue, MergedRegion, RawCellData, SheetMetadata};

/// ワークブックパーサー
///
/// calamineのラッパーとして、ワークブックレベルの操作を提供します。
/// Phase IIでは、XlsxMetadataParserを統合してXMLメタデータも取得します。
pub(crate) struct WorkbookParser<R: Read + Seek + Clone> {
    /// calamineのワークブック（XLSX形式のみサポート）
    workbook: Xlsx<R>,
    /// XMLメタデータパーサー（Phase II）
    metadata: Option<XlsxMetadataParser>,
}

impl WorkbookParser<std::io::Cursor<Vec<u8>>> {
    /// ワークブックを開き、XMLメタデータも解析する（Phase II）
    ///
    /// # 引数
    ///
    /// * `reader` - Excelファイルを読み込むためのリーダー（Read + Seekトレイトを実装）
    ///
    /// # 戻り値
    ///
    /// * `Ok(WorkbookParser)` - ワークブックとメタデータの読み込みに成功した場合
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    ///
    /// # 注意
    ///
    /// このメソッドは、ファイルを2回読み込む必要があるため、効率が悪い可能性があります。
    /// 将来的には、1回の読み込みで両方を処理する最適化を検討します。
    pub fn open_with_metadata<R: Read + Seek>(mut reader: R) -> Result<Self, XlsxToMdError> {
        use crate::security::SecurityConfig;

        // セキュリティチェック: 入力ファイルサイズの上限
        let security_config = SecurityConfig::default();

        // ファイル全体をメモリに読み込む（効率化のため）
        // セキュリティ: ファイルサイズ制限を適用
        let mut buffer = Vec::new();
        let bytes_read = reader.read_to_end(&mut buffer)?;

        if bytes_read as u64 > security_config.max_input_file_size {
            return Err(XlsxToMdError::SecurityViolation(format!(
                "Input file size exceeds maximum: {} bytes (max: {} bytes)",
                bytes_read, security_config.max_input_file_size
            )));
        }

        // calamineでワークブックを開く
        let sheets = open_workbook_auto_from_rs(Cursor::new(buffer.clone()))
            .map_err(XlsxToMdError::Parse)?;
        let workbook = match sheets {
            Sheets::Xlsx(workbook) => workbook,
            _ => {
                return Err(XlsxToMdError::Config(
                    "Only XLSX format is supported".to_string(),
                ))
            }
        };

        // XMLメタデータを解析
        let metadata = Some(XlsxMetadataParser::new(Cursor::new(buffer))?);

        Ok(WorkbookParser { workbook, metadata })
    }

    /// ワークブックを開き、既存のメタデータを再利用する
    ///
    /// # 引数
    ///
    /// * `reader` - Excelファイルを読み込むためのリーダー（Read + Seekトレイトを実装）
    /// * `metadata` - 再利用するメタデータ
    ///
    /// # 戻り値
    ///
    /// * `Ok(WorkbookParser)` - ワークブックの読み込みに成功した場合
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    pub fn open_with_existing_metadata<R: Read + Seek>(
        mut reader: R,
        metadata: XlsxMetadataParser,
    ) -> Result<Self, XlsxToMdError> {
        use crate::security::SecurityConfig;

        // セキュリティチェック: 入力ファイルサイズの上限
        let security_config = SecurityConfig::default();

        // ファイル全体をメモリに読み込む（効率化のため）
        // セキュリティ: ファイルサイズ制限を適用
        let mut buffer = Vec::new();
        let bytes_read = reader.read_to_end(&mut buffer)?;

        if bytes_read as u64 > security_config.max_input_file_size {
            return Err(XlsxToMdError::SecurityViolation(format!(
                "Input file size exceeds maximum: {} bytes (max: {} bytes)",
                bytes_read, security_config.max_input_file_size
            )));
        }

        // calamineでワークブックを開く
        let sheets = open_workbook_auto_from_rs(Cursor::new(buffer))
            .map_err(XlsxToMdError::Parse)?;
        let workbook = match sheets {
            Sheets::Xlsx(workbook) => workbook,
            _ => {
                return Err(XlsxToMdError::Config(
                    "Only XLSX format is supported".to_string(),
                ))
            }
        };

        // 既存のメタデータを再利用
        let metadata = Some(metadata);

        Ok(WorkbookParser { workbook, metadata })
    }
}

impl<R: Read + Seek + Clone> WorkbookParser<R> {
    /// ワークブックを開く（Phase I互換）
    ///
    /// # 引数
    ///
    /// * `reader` - Excelファイルを読み込むためのリーダー（Read + Seek + Cloneトレイトを実装）
    ///
    /// # 戻り値
    ///
    /// * `Ok(WorkbookParser)` - ワークブックの読み込みに成功した場合（XLSX形式のみサポート）
    /// * `Err(XlsxToMdError::Parse)` - ワークブックの読み込みに失敗した場合、またはXLSX形式でない場合
    #[allow(dead_code)]
    pub fn open(reader: R) -> Result<Self, XlsxToMdError> {
        let sheets = open_workbook_auto_from_rs(reader).map_err(XlsxToMdError::Parse)?;
        match sheets {
            Sheets::Xlsx(workbook) => Ok(Self {
                workbook,
                metadata: None, // Phase I: メタデータなし
            }),
            _ => Err(XlsxToMdError::Config(
                "Only XLSX format is supported".to_string(),
            )),
        }
    }

    /// すべてのシート名を取得
    ///
    /// # 戻り値
    ///
    /// シート名のベクター
    pub fn get_sheet_names(&self) -> Vec<String> {
        self.workbook.sheet_names().to_vec()
    }

    /// メタデータを取得（並列処理での再利用用）
    ///
    /// # 戻り値
    ///
    /// * `Some(&XlsxMetadataParser)` - メタデータが存在する場合
    /// * `None` - メタデータが存在しない場合
    pub fn metadata(&self) -> Option<&XlsxMetadataParser> {
        self.metadata.as_ref()
    }

    /// シート選択方式に基づいてシートを選択
    ///
    /// # 引数
    ///
    /// * `selector` - シート選択方式
    /// * `include_hidden` - 非表示シートを含めるかどうか（Phase Iでは常にfalseとして扱う）
    ///
    /// # 戻り値
    ///
    /// * `Ok(Vec<String>)` - 選択されたシート名のリスト
    /// * `Err(XlsxToMdError::Config)` - シートが見つからない、またはインデックスが範囲外の場合
    pub fn select_sheets(
        &self,
        selector: &SheetSelector,
        _include_hidden: bool,
    ) -> Result<Vec<String>, XlsxToMdError> {
        let all_sheet_names = self.get_sheet_names();

        match selector {
            SheetSelector::All => {
                // Phase I: 非表示シートの判定は未実装のため、すべてのシートを返す
                Ok(all_sheet_names)
            }

            SheetSelector::Index(index) => {
                if *index >= all_sheet_names.len() {
                    return Err(XlsxToMdError::Config(format!(
                        "Sheet index {} is out of range (total: {})",
                        index,
                        all_sheet_names.len()
                    )));
                }
                Ok(vec![all_sheet_names[*index].clone()])
            }

            SheetSelector::Name(name) => {
                if !all_sheet_names.contains(name) {
                    return Err(XlsxToMdError::Config(format!("Sheet '{}' not found", name)));
                }
                Ok(vec![name.clone()])
            }

            SheetSelector::Indices(indices) => {
                let mut result = Vec::new();
                for &index in indices {
                    if index >= all_sheet_names.len() {
                        return Err(XlsxToMdError::Config(format!(
                            "Sheet index {} is out of range (total: {})",
                            index,
                            all_sheet_names.len()
                        )));
                    }
                    result.push(all_sheet_names[index].clone());
                }
                Ok(result)
            }

            SheetSelector::Names(names) => {
                for name in names {
                    if !all_sheet_names.contains(name) {
                        return Err(XlsxToMdError::Config(format!("Sheet '{}' not found", name)));
                    }
                }
                Ok(names.clone())
            }
        }
    }

    /// シートをパースして、メタデータとセルデータを抽出
    ///
    /// # 引数
    ///
    /// * `sheet_name` - パースするシート名
    /// * `config` - 変換設定
    ///
    /// # 戻り値
    ///
    /// * `Ok((SheetMetadata, Vec<RawCellData>))` - メタデータとセルデータのペア
    /// * `Err(XlsxToMdError)` - パースエラーが発生した場合
    pub fn parse_sheet(
        &mut self,
        sheet_name: &str,
        config: &ConversionConfig,
    ) -> Result<(SheetMetadata, Vec<RawCellData>), XlsxToMdError> {
        // 1. シートの取得
        let range = self
            .workbook
            .worksheet_range(sheet_name)
            .map_err(|e| XlsxToMdError::Parse(e.into()))?;

        // 2. メタデータの収集
        let metadata = self.collect_metadata(sheet_name)?;

        // 3. 数式情報を事前に取得（全セルで再利用するため）
        // 注意: 各セルごとにworksheet_formula()を呼び出すと非常に遅いため、
        // 1回だけ呼び出して結果を全セルで再利用する
        let formula_range = self.workbook.worksheet_formula(sheet_name).ok();

        // 4. セルデータの抽出（ストリーミング処理）
        let mut cells = Vec::new();

        for (row_idx, row) in range.rows().enumerate() {
            let row_idx = row_idx as u32;

            // 非表示行のスキップ（Phase I: hidden_rowsは常に空リスト）
            if !config.include_hidden && metadata.hidden_rows.contains(&row_idx) {
                continue;
            }

            for (col_idx, cell) in row.iter().enumerate() {
                let col_idx = col_idx as u32;

                // 非表示列のスキップ（Phase I: hidden_colsは常に空リスト）
                if !config.include_hidden && metadata.hidden_cols.contains(&col_idx) {
                    continue;
                }

                let coord = CellCoord::new(row_idx, col_idx);

                // 範囲制限のチェック
                if let Some(range) = &config.range {
                    if !range.contains(coord) {
                        continue;
                    }
                }

                // RawCellDataの生成
                let raw_cell = self.extract_cell_data_with_formula(coord, cell, sheet_name, &formula_range)?;
                cells.push(raw_cell);
            }
        }

        Ok((metadata, cells))
    }

    /// セルデータを抽出（内部ヘルパーメソッド）
    ///
    /// # 引数
    ///
    /// * `coord` - セル座標
    /// * `cell` - calamineのセルデータ
    /// * `sheet_name` - シート名（数式取得用）
    ///
    /// # 戻り値
    ///
    /// * `Ok(RawCellData)` - 抽出されたセルデータ
    #[allow(dead_code)]
    fn extract_cell_data(
        &mut self,
        coord: CellCoord,
        cell: &Data,
        sheet_name: &str,
    ) -> Result<RawCellData, XlsxToMdError> {
        self.extract_cell_data_with_formula(coord, cell, sheet_name, &None)
    }

    /// セルデータを抽出（数式範囲を事前に取得したバージョン）
    ///
    /// # 引数
    ///
    /// * `coord` - セル座標
    /// * `cell` - calamineのセルデータ
    /// * `sheet_name` - シート名（数式取得用）
    /// * `formula_range` - 事前に取得した数式範囲（Noneの場合は取得を試みる）
    ///
    /// # 戻り値
    ///
    /// * `Ok(RawCellData)` - 抽出されたセルデータ
    fn extract_cell_data_with_formula(
        &mut self,
        coord: CellCoord,
        cell: &Data,
        sheet_name: &str,
        formula_range: &Option<Range<String>>,
    ) -> Result<RawCellData, XlsxToMdError> {
        // 1. 値の変換
        let value = match cell {
            Data::Int(i) => CellValue::Number(*i as f64),
            Data::Float(f) => CellValue::Number(*f),
            Data::String(s) => CellValue::String(s.clone()),
            Data::Bool(b) => CellValue::Bool(*b),
            Data::Error(e) => CellValue::Error(format!("{:?}", e)),
            Data::Empty => CellValue::Empty,
            _ => CellValue::Empty,
        };

        // 2. 書式情報の取得
        // Phase II: XlsxMetadataParserでxl/styles.xmlから取得
        let (format_id, format_string) = if let Some(ref metadata) = self.metadata {
            // calamineからstyle_idを取得（現在は未対応のため、None）
            // 将来的には、calamineのAPI拡張を待つか、XMLから直接取得する必要がある
            let style_id = None; // TODO: calamineからstyle_idを取得
            if let Some(style_id) = style_id {
                let fmt_str = metadata.get_format_string(style_id);
                (Some(style_id as u16), fmt_str.map(|s| s.to_string()))
            } else {
                (None, None)
            }
        } else {
            (None, None) // Phase I: メタデータなし
        };

        // 3. 数式情報の取得
        // Phase I: calamine 0.26以降のworksheet_formula() APIで取得可能
        // 事前に取得した数式範囲を使用（各セルごとに呼び出すと非常に遅い）
        let formula = if let Some(ref formula_range) = formula_range {
            // 座標をcalamineの形式に変換（(row, col)）
            let calamine_coord = (coord.row as usize, coord.col as usize);
            formula_range.get(calamine_coord).and_then(|f| {
                if f.is_empty() {
                    None
                } else {
                    Some(f.clone())
                }
            })
        } else {
            None
        };

        // 4. ハイパーリンク情報の取得
        // Phase II: XlsxMetadataParserでxl/worksheets/*.xmlと_rels/*.xml.relsから取得
        let hyperlink = if let Some(ref metadata) = self.metadata {
            metadata.hyperlinks.get(sheet_name).and_then(|sheet_links| {
                sheet_links
                    .get(&(coord.row, coord.col))
                    .map(|h| h.url.clone())
            })
        } else {
            None
        };

        // 5. リッチテキスト情報の取得
        // Phase II: XlsxMetadataParserでxl/sharedStrings.xmlとxl/worksheets/*.xmlから取得
        let rich_text = if let Some(ref metadata) = self.metadata {
            metadata
                .cell_string_indices
                .get(sheet_name)
                .and_then(|sheet_indices| {
                    sheet_indices
                        .get(&(coord.row, coord.col))
                        .and_then(|index| metadata.shared_strings.get(index).cloned())
                })
        } else {
            None
        };

        Ok(RawCellData {
            coord,
            value,
            format_id,
            format_string,
            formula,
            hyperlink,
            rich_text,
        })
    }

    /// シートのメタデータを収集
    ///
    /// # 引数
    ///
    /// * `sheet_name` - シート名
    ///
    /// # 戻り値
    ///
    /// * `Ok(SheetMetadata)` - シートのメタデータ
    /// * `Err(XlsxToMdError)` - エラーが発生した場合
    fn collect_metadata(&mut self, sheet_name: &str) -> Result<SheetMetadata, XlsxToMdError> {
        // 1. シートインデックスの取得
        let index = self
            .workbook
            .sheet_names()
            .iter()
            .position(|name| name == sheet_name)
            .ok_or_else(|| XlsxToMdError::Config(format!("Sheet '{}' not found", sheet_name)))?;

        // 2. 非表示フラグの取得
        // Phase I: calamine APIで非表示シート情報は未サポート
        // Phase II: XlsxMetadataParserで取得予定
        let hidden = false; // Phase I: false固定

        // 3. 結合セル範囲の取得
        // Phase I: calamine 0.26以降で完全対応
        self.workbook
            .load_merged_regions()
            .map_err(|e| XlsxToMdError::Parse(e.into()))?;
        let merged_regions = match self.workbook.worksheet_merge_cells(sheet_name) {
            Some(Ok(regions)) => regions
                .iter()
                .map(|dims| {
                    let start = CellCoord::new(dims.start.0, dims.start.1);
                    let end = CellCoord::new(dims.end.0, dims.end.1);
                    let range = CellRange::new(start, end);
                    MergedRegion::new(range)
                })
                .collect(),
            Some(Err(_)) | None => Vec::new(),
        };

        // 4. 非表示行・列のリスト
        // Phase II: XlsxMetadataParserでxl/worksheets/*.xmlから取得
        let (hidden_rows, hidden_cols) = if let Some(ref metadata) = self.metadata {
            let rows: Vec<u32> = metadata
                .hidden_rows
                .get(sheet_name)
                .map(|set| set.iter().copied().collect())
                .unwrap_or_default();
            let cols: Vec<u32> = metadata
                .hidden_cols
                .get(sheet_name)
                .map(|set| set.iter().copied().collect())
                .unwrap_or_default();
            (rows, cols)
        } else {
            (Vec::new(), Vec::new()) // Phase I: 空リスト
        };

        // 5. 1904年エポックフラグ
        // Phase II: XlsxMetadataParserでxl/workbook.xmlから取得
        let is_1904 = self.metadata.as_ref().map(|m| m.is_1904()).unwrap_or(false); // Phase I: デフォルトはfalse

        Ok(SheetMetadata {
            name: sheet_name.to_string(),
            index,
            hidden,
            merged_regions,
            hidden_rows,
            hidden_cols,
            is_1904,
        })
    }
}

// テストは統合テスト（tests/）で実装します。
// 実際のXLSXファイルが必要なため、単体テストではなく統合テストとして実装します。
