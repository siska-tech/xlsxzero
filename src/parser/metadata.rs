//! XML Metadata Parser Module
//!
//! XLSX内部のXMLファイルから、calamineで取得不可能な情報を抽出するモジュール。
//! Number Format String、非表示行/列、1904年エポック判定などを提供します。

use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};
use zip::ZipArchive;

use crate::error::XlsxToMdError;
use crate::security::{validate_zip_path, SecurityConfig};
use crate::types::{RichTextFormat, RichTextSegment};

/// セルスタイル情報（cellXfs要素）
#[derive(Debug, Clone)]
pub(crate) struct CellXf {
    pub num_fmt_id: u32,
    #[allow(dead_code)]
    pub font_id: Option<u32>,
    #[allow(dead_code)]
    pub fill_id: Option<u32>,
    #[allow(dead_code)]
    pub border_id: Option<u32>,
}

/// ハイパーリンク情報
#[derive(Debug, Clone)]
pub(crate) struct Hyperlink {
    /// URL
    pub url: String,
    /// 表示テキスト（セルの値、またはURLと同じ）
    #[allow(dead_code)]
    pub display: Option<String>,
}

/// XLSXメタデータパーサー
///
/// XLSXファイル（ZIPアーカイブ）からXMLを直接解析し、
/// calamineで取得できない情報を抽出します。
#[derive(Clone)]
pub(crate) struct XlsxMetadataParser {
    /// numFmtId -> formatCode のマッピング
    num_formats: HashMap<u32, String>,
    /// styleId -> CellXf のマッピング
    cell_xfs: Vec<CellXf>,
    /// シート名 -> 非表示行インデックスのセット
    pub(crate) hidden_rows: HashMap<String, HashSet<u32>>,
    /// シート名 -> 非表示列インデックスのセット
    pub(crate) hidden_cols: HashMap<String, HashSet<u32>>,
    /// シート名 -> セル座標 -> ハイパーリンク情報のマッピング
    pub(crate) hyperlinks: HashMap<String, HashMap<(u32, u32), Hyperlink>>,
    /// 1904年エポックを使用するかどうか
    is_1904: bool,
    /// 共有文字列インデックス -> リッチテキストセグメントのマッピング
    /// 通常のテキストの場合は、1つのプレーンテキストセグメントを含む
    pub(crate) shared_strings: HashMap<u32, Vec<RichTextSegment>>,
    /// シート名 -> セル座標 -> 共有文字列インデックスのマッピング
    pub(crate) cell_string_indices: HashMap<String, HashMap<(u32, u32), u32>>,
}

impl XlsxMetadataParser {
    /// XLSXファイル（ZIPアーカイブ）からメタデータを解析
    ///
    /// # 引数
    ///
    /// * `xlsx_reader` - XLSXファイルを読み込むためのリーダー（Read + Seekトレイトを実装）
    ///
    /// # 戻り値
    ///
    /// * `Ok(XlsxMetadataParser)` - メタデータの解析に成功した場合
    /// * `Err(XlsxToMdError)` - 解析エラーが発生した場合
    pub fn new<R: Read + Seek>(xlsx_reader: R) -> Result<Self, XlsxToMdError> {
        let security_config = SecurityConfig::default();

        let mut archive =
            ZipArchive::new(xlsx_reader).map_err(|e| XlsxToMdError::Zip(format!("{}", e)))?;

        // セキュリティチェック: ファイル数の上限
        if archive.len() > security_config.max_file_count {
            return Err(XlsxToMdError::SecurityViolation(format!(
                "ZIP archive contains too many files: {} (max: {})",
                archive.len(),
                security_config.max_file_count
            )));
        }

        // セキュリティチェック: 各ファイルのパス検証とサイズチェック
        let mut total_decompressed_size = 0u64;
        for i in 0..archive.len() {
            let file = archive
                .by_index(i)
                .map_err(|e| XlsxToMdError::Zip(format!("{}", e)))?;

            // パストラバーサル対策
            let file_name = file.name();
            validate_zip_path(file_name).map_err(|e| {
                XlsxToMdError::SecurityViolation(format!("Invalid ZIP path: {}", e))
            })?;

            // ファイルサイズチェック
            let file_size = file.size();
            if file_size > security_config.max_file_size {
                return Err(XlsxToMdError::SecurityViolation(format!(
                    "File '{}' exceeds maximum size: {} bytes (max: {} bytes)",
                    file_name, file_size, security_config.max_file_size
                )));
            }

            // 展開後のサイズ累計をチェック
            total_decompressed_size =
                total_decompressed_size
                    .checked_add(file_size)
                    .ok_or_else(|| {
                        XlsxToMdError::SecurityViolation(
                            "Total decompressed size calculation overflow".to_string(),
                        )
                    })?;

            if total_decompressed_size > security_config.max_decompressed_size {
                return Err(XlsxToMdError::SecurityViolation(format!(
                    "Total decompressed size exceeds maximum: {} bytes (max: {} bytes)",
                    total_decompressed_size, security_config.max_decompressed_size
                )));
            }
        }

        // 1. xl/styles.xml を解析
        let (num_formats, cell_xfs) = Self::parse_styles(&mut archive)?;

        // 2. xl/sharedStrings.xml を解析
        let shared_strings = Self::parse_shared_strings(&mut archive)?;

        // 3. xl/worksheets/*.xml を解析
        let (hidden_rows, hidden_cols, cell_string_indices) = Self::parse_worksheets(&mut archive)?;

        // 4. ハイパーリンク情報を解析
        let hyperlinks = Self::parse_hyperlinks(&mut archive)?;

        // 5. xl/workbook.xml を解析
        let is_1904 = Self::parse_workbook(&mut archive)?;

        Ok(Self {
            num_formats,
            cell_xfs,
            hidden_rows,
            hidden_cols,
            hyperlinks,
            is_1904,
            shared_strings,
            cell_string_indices,
        })
    }

    /// styleIdからNumber Format Stringを取得
    ///
    /// # 引数
    ///
    /// * `style_id` - スタイルID（0始まり）
    ///
    /// # 戻り値
    ///
    /// * `Some(&str)` - フォーマット文字列が見つかった場合
    /// * `None` - スタイルIDが範囲外、またはフォーマットが見つからない場合
    pub fn get_format_string(&self, style_id: u32) -> Option<&str> {
        self.cell_xfs.get(style_id as usize).and_then(|xf| {
            // ビルトイン書式ID（0-163）の場合はハードコードマッピングを使用
            if xf.num_fmt_id < 164 {
                get_builtin_format(xf.num_fmt_id)
            } else {
                // カスタム書式ID（>= 164）の場合はnum_formatsから取得
                self.num_formats.get(&xf.num_fmt_id).map(|s| s.as_str())
            }
        })
    }

    /// 行が非表示かどうかを判定
    ///
    /// # 引数
    ///
    /// * `sheet_name` - シート名
    /// * `row` - 行インデックス（0始まり）
    ///
    /// # 戻り値
    ///
    /// * `true` - 行が非表示の場合
    /// * `false` - 行が表示されている、または情報が取得できない場合
    #[allow(dead_code)]
    pub fn is_row_hidden(&self, sheet_name: &str, row: u32) -> bool {
        self.hidden_rows
            .get(sheet_name)
            .map(|rows| rows.contains(&row))
            .unwrap_or(false)
    }

    /// 列が非表示かどうかを判定
    ///
    /// # 引数
    ///
    /// * `sheet_name` - シート名
    /// * `col` - 列インデックス（0始まり）
    ///
    /// # 戻り値
    ///
    /// * `true` - 列が非表示の場合
    /// * `false` - 列が表示されている、または情報が取得できない場合
    #[allow(dead_code)]
    pub fn is_col_hidden(&self, sheet_name: &str, col: u32) -> bool {
        self.hidden_cols
            .get(sheet_name)
            .map(|cols| cols.contains(&col))
            .unwrap_or(false)
    }

    /// 1904年エポックを使用するかどうかを取得
    ///
    /// # 戻り値
    ///
    /// * `true` - 1904年エポックを使用する場合
    /// * `false` - 1900年エポックを使用する場合（デフォルト）
    pub fn is_1904(&self) -> bool {
        self.is_1904
    }

    /// xl/sharedStrings.xml の解析（プライベート）
    ///
    /// `<sst>` 要素を解析し、リッチテキスト情報を抽出します。
    fn parse_shared_strings<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
    ) -> Result<HashMap<u32, Vec<RichTextSegment>>, XlsxToMdError> {
        let mut shared_strings = HashMap::new();

        // xl/sharedStrings.xml を開く
        let mut shared_strings_file = match archive.by_name("xl/sharedStrings.xml") {
            Ok(file) => file,
            Err(_) => {
                // sharedStrings.xmlが存在しない場合は空の結果を返す
                return Ok(shared_strings);
            }
        };

        // ZIPファイルの内容を一度メモリに読み込む
        use std::io::Read;
        let mut xml_content = Vec::new();
        shared_strings_file.read_to_end(&mut xml_content)?;

        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_reader(xml_content.as_slice());
        reader.trim_text(true);

        let mut buf = Vec::new();
        let mut in_si = false;
        let mut in_r = false;
        let mut in_t = false;
        let mut current_index: u32 = 0;
        let mut current_segments: Vec<RichTextSegment> = Vec::new();
        let mut current_segment_text = String::new();
        let mut current_format = RichTextFormat::new();
        let mut has_r_element = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    match e.name().as_ref() {
                        b"si" => {
                            // <si> 要素の開始
                            in_si = true;
                            current_segments.clear();
                            current_segment_text.clear();
                            current_format = RichTextFormat::new();
                            has_r_element = false;
                        }
                        b"r" if in_si => {
                            // <r> 要素の開始（リッチテキストセグメント）
                            in_r = true;
                            has_r_element = true;
                            current_segment_text.clear();
                            current_format = RichTextFormat::new();
                        }
                        b"rPr" if in_r => {
                            // <rPr> 要素の開始（書式プロパティ）
                            // 書式プロパティは子要素で定義される
                        }
                        b"b" if in_r => {
                            // <b/> 要素（太字）
                            current_format.bold = true;
                        }
                        b"i" if in_r => {
                            // <i/> 要素（斜体）
                            current_format.italic = true;
                        }
                        b"t" if in_si => {
                            // <t> 要素の開始（テキスト）
                            in_t = true;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_t {
                        let text = e
                            .unescape()
                            .map_err(|e| XlsxToMdError::Config(format!("XML text error: {}", e)))?;
                        current_segment_text.push_str(&text);
                    }
                }
                Ok(Event::End(e)) => {
                    match e.name().as_ref() {
                        b"si" => {
                            // <si> 要素の終了
                            if has_r_element {
                                // <r>要素がある場合：現在のセグメントを追加
                                if !current_segment_text.is_empty() {
                                    current_segments.push(RichTextSegment::new(
                                        current_segment_text.clone(),
                                        current_format.clone(),
                                    ));
                                }
                            } else {
                                // 通常のテキスト（<r>要素がない場合）
                                if !current_segment_text.is_empty() {
                                    current_segments
                                        .push(RichTextSegment::plain(current_segment_text.clone()));
                                }
                            }
                            if !current_segments.is_empty() {
                                shared_strings.insert(current_index, current_segments.clone());
                            }
                            current_index += 1;
                            in_si = false;
                            current_segments.clear();
                            current_segment_text.clear();
                            current_format = RichTextFormat::new();
                            has_r_element = false;
                        }
                        b"r" if in_r => {
                            // <r> 要素の終了（リッチテキストセグメント）
                            if !current_segment_text.is_empty() {
                                current_segments.push(RichTextSegment::new(
                                    current_segment_text.clone(),
                                    current_format.clone(),
                                ));
                                current_segment_text.clear();
                            }
                            in_r = false;
                            current_format = RichTextFormat::new();
                        }
                        b"t" if in_t => {
                            // <t> 要素の終了
                            in_t = false;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxToMdError::Config(format!("XML parse error: {}", e))),
                _ => {}
            }
        }

        Ok(shared_strings)
    }

    /// xl/styles.xml の解析（プライベート）
    ///
    /// `<numFmts>` と `<cellXfs>` を解析し、Number Format Stringのマッピングを構築します。
    fn parse_styles<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
    ) -> Result<(HashMap<u32, String>, Vec<CellXf>), XlsxToMdError> {
        let mut num_formats = HashMap::new();
        let mut cell_xfs = Vec::new();

        // xl/styles.xml を開く（パストラバーサル対策済み）
        let mut styles_file = match archive.by_name("xl/styles.xml") {
            Ok(file) => file,
            Err(_) => {
                // styles.xmlが存在しない場合は空の結果を返す
                return Ok((num_formats, cell_xfs));
            }
        };

        // ZIPファイルの内容を一度メモリに読み込む
        use std::io::Read;
        let mut xml_content = Vec::new();
        styles_file.read_to_end(&mut xml_content)?;

        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_reader(xml_content.as_slice());
        reader.trim_text(true);

        let mut buf = Vec::new();
        let mut in_num_fmts = false;
        let mut in_cell_xfs = false;
        let mut current_num_fmt_id: Option<u32> = None;
        let mut current_num_fmt_code: Option<String> = None;
        let mut current_xf: Option<CellXf> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    match e.name().as_ref() {
                        b"numFmts" => {
                            in_num_fmts = true;
                        }
                        b"numFmt" if in_num_fmts => {
                            // <numFmt numFmtId="165" formatCode="0.000"/>
                            current_num_fmt_id = None;
                            current_num_fmt_code = None;
                            for attr in e.attributes() {
                                let attr = attr.map_err(|e| {
                                    XlsxToMdError::Config(format!("XML attribute error: {}", e))
                                })?;
                                match attr.key.as_ref() {
                                    b"numFmtId" => {
                                        let id_str = std::str::from_utf8(&attr.value)?;
                                        current_num_fmt_id = Some(id_str.parse()?);
                                    }
                                    b"formatCode" => {
                                        current_num_fmt_code =
                                            Some(std::str::from_utf8(&attr.value)?.to_string());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"cellXfs" => {
                            in_cell_xfs = true;
                        }
                        b"xf" if in_cell_xfs => {
                            // <xf numFmtId="165" fontId="0" fillId="0" borderId="0"/>
                            let mut num_fmt_id = 0u32;
                            let mut font_id = None;
                            let mut fill_id = None;
                            let mut border_id = None;

                            for attr in e.attributes() {
                                let attr = attr.map_err(|e| {
                                    XlsxToMdError::Config(format!("XML attribute error: {}", e))
                                })?;
                                match attr.key.as_ref() {
                                    b"numFmtId" => {
                                        let id_str = std::str::from_utf8(&attr.value)?;
                                        num_fmt_id = id_str.parse()?;
                                    }
                                    b"fontId" => {
                                        let id_str = std::str::from_utf8(&attr.value)?;
                                        font_id = Some(id_str.parse()?);
                                    }
                                    b"fillId" => {
                                        let id_str = std::str::from_utf8(&attr.value)?;
                                        fill_id = Some(id_str.parse()?);
                                    }
                                    b"borderId" => {
                                        let id_str = std::str::from_utf8(&attr.value)?;
                                        border_id = Some(id_str.parse()?);
                                    }
                                    _ => {}
                                }
                            }

                            current_xf = Some(CellXf {
                                num_fmt_id,
                                font_id,
                                fill_id,
                                border_id,
                            });
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => {
                    match e.name().as_ref() {
                        b"numFmts" => {
                            in_num_fmts = false;
                        }
                        b"numFmt" if in_num_fmts => {
                            if let (Some(id), Some(code)) =
                                (current_num_fmt_id, current_num_fmt_code.take())
                            {
                                // カスタム書式ID（>= 164）のみ保存
                                if id >= 164 {
                                    num_formats.insert(id, code);
                                }
                            }
                        }
                        b"cellXfs" => {
                            in_cell_xfs = false;
                        }
                        b"xf" if in_cell_xfs => {
                            if let Some(xf) = current_xf.take() {
                                cell_xfs.push(xf);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxToMdError::Config(format!("XML parse error: {}", e))),
                _ => {}
            }
        }

        Ok((num_formats, cell_xfs))
    }

    /// xl/worksheets/*.xml の解析（プライベート）
    ///
    /// すべてのワークシートXMLファイルを解析し、非表示行・列の情報を収集します。
    #[allow(clippy::type_complexity)]
    fn parse_worksheets<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
    ) -> Result<
        (
            HashMap<String, HashSet<u32>>,
            HashMap<String, HashSet<u32>>,
            HashMap<String, HashMap<(u32, u32), u32>>,
        ),
        XlsxToMdError,
    > {
        let mut hidden_rows: HashMap<String, HashSet<u32>> = HashMap::new();
        let mut hidden_cols: HashMap<String, HashSet<u32>> = HashMap::new();
        let mut cell_string_indices: HashMap<String, HashMap<(u32, u32), u32>> = HashMap::new();

        // すべてのワークシートXMLファイルを検索
        for i in 0..archive.len() {
            let file_name = archive
                .by_index(i)
                .map_err(|e| XlsxToMdError::Zip(format!("{}", e)))?
                .name()
                .to_string();

            // セキュリティ: パストラバーサル対策（既にnew()で検証済みだが、念のため再確認）
            validate_zip_path(&file_name).map_err(|e| {
                XlsxToMdError::SecurityViolation(format!("Invalid ZIP path: {}", e))
            })?;

            if file_name.starts_with("xl/worksheets/sheet") && file_name.ends_with(".xml") {
                // シート名を抽出（例: "xl/worksheets/sheet1.xml" -> "Sheet1"）
                // 実際のシート名はworkbook.xmlから取得する必要があるが、
                // ここではファイル名から推測する（簡易実装）
                let sheet_name = Self::extract_sheet_name_from_path(&file_name);

                // 非表示行・列と共有文字列インデックスを解析
                let mut file = archive
                    .by_name(&file_name)
                    .map_err(|e| XlsxToMdError::Zip(format!("{}", e)))?;
                let (rows, cols, string_indices) = Self::parse_worksheet_xml(&mut file)?;
                if !rows.is_empty() {
                    hidden_rows.insert(sheet_name.clone(), rows);
                }
                if !cols.is_empty() {
                    hidden_cols.insert(sheet_name.clone(), cols);
                }
                if !string_indices.is_empty() {
                    cell_string_indices.insert(sheet_name, string_indices);
                }
            }
        }

        Ok((hidden_rows, hidden_cols, cell_string_indices))
    }

    /// ワークシートXMLファイルから非表示行・列と共有文字列インデックスを解析
    #[allow(clippy::type_complexity)]
    fn parse_worksheet_xml(
        reader: &mut zip::read::ZipFile<'_>,
    ) -> Result<(HashSet<u32>, HashSet<u32>, HashMap<(u32, u32), u32>), XlsxToMdError> {
        use quick_xml::events::Event;
        use quick_xml::Reader;
        use std::io::Read;

        // ZIPファイルの内容を一度メモリに読み込む
        let mut xml_content = Vec::new();
        reader.read_to_end(&mut xml_content)?;

        let mut xml_reader = Reader::from_reader(xml_content.as_slice());
        xml_reader.trim_text(true);

        let mut buf = Vec::new();
        let mut hidden_rows = HashSet::new();
        let mut hidden_cols = HashSet::new();
        let mut cell_string_indices = HashMap::new();
        let mut in_cols = false;
        let mut in_row = false;
        let mut in_cell = false;
        let mut current_row_num: Option<u32> = None;
        let mut current_col_num: Option<u32> = None;
        let mut current_cell_type: Option<String> = None;
        let mut current_cell_value: Option<String> = None;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    match e.name().as_ref() {
                        b"cols" => {
                            in_cols = true;
                        }
                        b"col" if in_cols => {
                            // <col min="3" max="3" hidden="1"/>
                            let mut current_col_min: Option<u32> = None;
                            let mut current_col_max: Option<u32> = None;
                            let mut is_hidden = false;

                            for attr in e.attributes() {
                                let attr = attr.map_err(|e| {
                                    XlsxToMdError::Config(format!("XML attribute error: {}", e))
                                })?;
                                match attr.key.as_ref() {
                                    b"min" => {
                                        let min_str = std::str::from_utf8(&attr.value)?;
                                        // Excelの列番号は1始まりなので、0始まりに変換
                                        current_col_min = Some(min_str.parse::<u32>()? - 1);
                                    }
                                    b"max" => {
                                        let max_str = std::str::from_utf8(&attr.value)?;
                                        current_col_max = Some(max_str.parse::<u32>()? - 1);
                                    }
                                    b"hidden" => {
                                        let hidden_str = std::str::from_utf8(&attr.value)?;
                                        is_hidden = hidden_str == "1" || hidden_str == "true";
                                    }
                                    _ => {}
                                }
                            }

                            if is_hidden {
                                if let (Some(min), Some(max)) = (current_col_min, current_col_max) {
                                    for col in min..=max {
                                        hidden_cols.insert(col);
                                    }
                                }
                            }
                        }
                        b"row" => {
                            // <row r="15" hidden="1">
                            in_row = true;
                            current_row_num = None;
                            let mut is_hidden = false;

                            for attr in e.attributes() {
                                let attr = attr.map_err(|e| {
                                    XlsxToMdError::Config(format!("XML attribute error: {}", e))
                                })?;
                                match attr.key.as_ref() {
                                    b"r" => {
                                        let r_str = std::str::from_utf8(&attr.value)?;
                                        // Excelの行番号は1始まりなので、0始まりに変換
                                        current_row_num = Some(r_str.parse::<u32>()? - 1);
                                    }
                                    b"hidden" => {
                                        let hidden_str = std::str::from_utf8(&attr.value)?;
                                        is_hidden = hidden_str == "1" || hidden_str == "true";
                                    }
                                    _ => {}
                                }
                            }

                            if is_hidden {
                                if let Some(row) = current_row_num {
                                    hidden_rows.insert(row);
                                }
                            }
                        }
                        b"c" if in_row => {
                            // <c r="A1" t="s">
                            in_cell = true;
                            current_col_num = None;
                            current_cell_type = None;
                            current_cell_value = None;

                            for attr in e.attributes() {
                                let attr = attr.map_err(|e| {
                                    XlsxToMdError::Config(format!("XML attribute error: {}", e))
                                })?;
                                match attr.key.as_ref() {
                                    b"r" => {
                                        let ref_str = std::str::from_utf8(&attr.value)?;
                                        // セル参照から行・列を抽出（例: "A1" -> (0, 0)）
                                        if let Some((row, col)) = Self::parse_cell_ref(ref_str) {
                                            current_row_num = Some(row);
                                            current_col_num = Some(col);
                                        }
                                    }
                                    b"t" => {
                                        let t_str = std::str::from_utf8(&attr.value)?;
                                        current_cell_type = Some(t_str.to_string());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"v" if in_cell => {
                            // <v>0</v> - 共有文字列インデックス
                            // テキストを読み込む準備
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_cell {
                        let text = e
                            .unescape()
                            .map_err(|e| XlsxToMdError::Config(format!("XML text error: {}", e)))?;
                        current_cell_value = Some(text.to_string());
                    }
                }
                Ok(Event::End(e)) => {
                    match e.name().as_ref() {
                        b"cols" => {
                            in_cols = false;
                        }
                        b"row" => {
                            in_row = false;
                            current_row_num = None;
                        }
                        b"c" if in_cell => {
                            // セルが終了したら、共有文字列インデックスを保存
                            if let (Some(row), Some(col), Some(cell_type), Some(cell_value)) = (
                                current_row_num,
                                current_col_num,
                                current_cell_type,
                                current_cell_value.take(),
                            ) {
                                if cell_type == "s" {
                                    // セルタイプが"s"（shared string）の場合
                                    if let Ok(index) = cell_value.parse::<u32>() {
                                        cell_string_indices.insert((row, col), index);
                                    }
                                }
                            }
                            in_cell = false;
                            current_col_num = None;
                            current_cell_type = None;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxToMdError::Config(format!("XML parse error: {}", e))),
                _ => {}
            }
        }

        Ok((hidden_rows, hidden_cols, cell_string_indices))
    }

    /// ハイパーリンク情報を解析
    ///
    /// ワークシートXMLとリレーションシップファイルからハイパーリンク情報を取得します。
    #[allow(clippy::type_complexity)]
    fn parse_hyperlinks<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
    ) -> Result<HashMap<String, HashMap<(u32, u32), Hyperlink>>, XlsxToMdError> {
        let mut hyperlinks: HashMap<String, HashMap<(u32, u32), Hyperlink>> = HashMap::new();

        // 1回のループでリレーションシップファイルとワークシートXMLの両方を処理
        // リレーションシップファイルを先に収集してから、ワークシートXMLを処理
        let mut rels_map: HashMap<String, HashMap<String, String>> = HashMap::new();
        let mut worksheet_files: Vec<(String, String)> = Vec::new(); // (file_name, sheet_name)

        for i in 0..archive.len() {
            let file_name = archive
                .by_index(i)
                .map_err(|e| XlsxToMdError::Zip(format!("{}", e)))?
                .name()
                .to_string();

            // セキュリティ: パストラバーサル対策（既にnew()で検証済みだが、念のため再確認）
            validate_zip_path(&file_name).map_err(|e| {
                XlsxToMdError::SecurityViolation(format!("Invalid ZIP path: {}", e))
            })?;

            // リレーションシップファイルの処理
            if file_name.contains("_rels") && file_name.ends_with(".xml.rels") {
                // ワークシートのリレーションシップファイルのみを処理
                if file_name.contains("worksheets/_rels/sheet") {
                    let sheet_name = Self::extract_sheet_name_from_rels_path(&file_name);
                    let mut file = archive
                        .by_name(&file_name)
                        .map_err(|e| XlsxToMdError::Zip(format!("{}", e)))?;
                    let rels = Self::parse_relationships(&mut file)?;
                    if !rels.is_empty() {
                        rels_map.insert(sheet_name, rels);
                    }
                }
            }
            // ワークシートXMLファイルの収集
            else if file_name.starts_with("xl/worksheets/sheet") && file_name.ends_with(".xml") {
                let sheet_name = Self::extract_sheet_name_from_path(&file_name);
                worksheet_files.push((file_name, sheet_name));
            }
        }

        // 2. ワークシートXMLからハイパーリンク要素を解析
        for (file_name, sheet_name) in worksheet_files {
            // 対応するリレーションシップファイルを探す
            // ファイル名から番号を抽出（例: "sheet1.xml" -> 1）
            let sheet_num = if let Some(name) = file_name.strip_prefix("xl/worksheets/sheet") {
                if let Some(num_str) = name.strip_suffix(".xml") {
                    num_str.parse::<usize>().ok()
                } else {
                    None
                }
            } else {
                None
            };

            // リレーションシップファイル名を構築（例: "xl/worksheets/_rels/sheet1.xml.rels"）
            let rels_file_name = if let Some(num) = sheet_num {
                format!("xl/worksheets/_rels/sheet{}.xml.rels", num)
            } else {
                continue;
            };

            // リレーションシップを取得（先に収集したマップから取得、なければ直接読み込む）
            let rels_for_sheet = if let Some(rels) = rels_map.get(&sheet_name) {
                Some(rels.clone())
            } else if let Ok(mut rels_file) = archive.by_name(&rels_file_name) {
                Self::parse_relationships(&mut rels_file).ok()
            } else {
                None
            };

            let mut file = archive
                .by_name(&file_name)
                .map_err(|e| XlsxToMdError::Zip(format!("{}", e)))?;
            let sheet_hyperlinks =
                Self::parse_worksheet_hyperlinks(&mut file, &rels_for_sheet.as_ref())?;

            if !sheet_hyperlinks.is_empty() {
                hyperlinks.insert(sheet_name, sheet_hyperlinks);
            }
        }

        Ok(hyperlinks)
    }

    /// リレーションシップファイルを解析
    fn parse_relationships(
        reader: &mut zip::read::ZipFile<'_>,
    ) -> Result<HashMap<String, String>, XlsxToMdError> {
        use quick_xml::events::Event;
        use quick_xml::Reader;
        use std::io::Read;

        let mut xml_content = Vec::new();
        reader.read_to_end(&mut xml_content)?;

        let mut xml_reader = Reader::from_reader(xml_content.as_slice());
        xml_reader.trim_text(true);

        let mut buf = Vec::new();
        let mut relationships = HashMap::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    // Event::Emptyは自己終了タグの場合に発生
                    let name = e.name();
                    if name.as_ref() == b"Relationship" {
                        let mut id = None;
                        let mut target = None;

                        for attr in e.attributes() {
                            let attr = attr.map_err(|e| {
                                XlsxToMdError::Config(format!("XML attribute error: {}", e))
                            })?;
                            match attr.key.as_ref() {
                                b"Id" => {
                                    id = Some(std::str::from_utf8(&attr.value)?.to_string());
                                }
                                b"Target" => {
                                    target = Some(std::str::from_utf8(&attr.value)?.to_string());
                                }
                                _ => {}
                            }
                        }

                        match (id, target) {
                            (Some(id_val), Some(target_val)) => {
                                relationships.insert(id_val, target_val);
                            }
                            _ => {
                                // リレーションシップIDまたはターゲットが欠落している場合はスキップ
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxToMdError::Config(format!("XML parse error: {}", e))),
                _ => {}
            }
        }

        Ok(relationships)
    }

    /// ワークシートXMLからハイパーリンク要素を解析
    fn parse_worksheet_hyperlinks(
        reader: &mut zip::read::ZipFile<'_>,
        relationships: &Option<&HashMap<String, String>>,
    ) -> Result<HashMap<(u32, u32), Hyperlink>, XlsxToMdError> {
        use quick_xml::events::Event;
        use quick_xml::Reader;
        use std::io::Read;

        let mut xml_content = Vec::new();
        reader.read_to_end(&mut xml_content)?;

        let mut xml_reader = Reader::from_reader(xml_content.as_slice());
        xml_reader.trim_text(true);

        let mut buf = Vec::new();
        let mut hyperlinks = HashMap::new();
        let mut in_hyperlinks = false; // <hyperlinks>要素内にいるかどうか

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    // Event::Emptyは自己終了タグ（<hyperlink ... />）の場合に発生
                    let name = e.name();
                    let name_bytes = name.as_ref();

                    // <hyperlinks>要素の開始を検出
                    if name_bytes == b"hyperlinks" {
                        in_hyperlinks = true;
                        continue; // <hyperlinks>要素自体は処理しない
                    }

                    // <hyperlinks>要素内の<hyperlink>要素を検出
                    // 注意: <hyperlink>要素は自己終了タグ（<hyperlink ... />）の可能性がある
                    if in_hyperlinks && name_bytes == b"hyperlink" {
                        let mut ref_attr = None;
                        let mut relationship_id = None;

                        for attr_result in e.attributes() {
                            let attr = attr_result.map_err(|e| {
                                XlsxToMdError::Config(format!("XML attribute error: {}", e))
                            })?;
                            let key_bytes = attr.key.as_ref();

                            if key_bytes == b"ref" {
                                // セル参照（例: "A1"）
                                ref_attr = Some(std::str::from_utf8(&attr.value)?.to_string());
                            } else if key_bytes == b"r:id" {
                                // リレーションシップID（"r:id"）
                                relationship_id =
                                    Some(std::str::from_utf8(&attr.value)?.to_string());
                            }
                        }

                        if let Some(ref_str) = ref_attr {
                            // セル参照を座標に変換（例: "A1" -> (0, 0)）
                            if let Some(coord) = Self::parse_cell_ref(&ref_str) {
                                let url = if let Some(rel_id) = relationship_id {
                                    // リレーションシップからURLを取得
                                    relationships
                                        .and_then(|rels| rels.get(&rel_id))
                                        .cloned()
                                        .unwrap_or_default()
                                } else {
                                    // リレーションシップIDがない場合は、ref属性をそのまま使用（外部URLの場合）
                                    // ただし、通常はリレーションシップIDが必要
                                    String::new()
                                };

                                if !url.is_empty() {
                                    hyperlinks.insert(
                                        coord,
                                        Hyperlink {
                                            url,
                                            display: None, // 表示テキストはセルの値から取得
                                        },
                                    );
                                }
                            }
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    // <hyperlinks>要素の終了を検出
                    if e.name().as_ref() == b"hyperlinks" {
                        in_hyperlinks = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxToMdError::Config(format!("XML parse error: {}", e))),
                _ => {}
            }
        }

        Ok(hyperlinks)
    }

    /// セル参照文字列を座標に変換（例: "A1" -> (0, 0)）
    fn parse_cell_ref(ref_str: &str) -> Option<(u32, u32)> {
        // 簡単な実装: "A1"形式を想定
        let mut col_str = String::new();
        let mut row_str = String::new();

        for ch in ref_str.chars() {
            if ch.is_ascii_alphabetic() {
                col_str.push(ch);
            } else if ch.is_ascii_digit() {
                row_str.push(ch);
            }
        }

        if col_str.is_empty() || row_str.is_empty() {
            return None;
        }

        // 列を数値に変換（A=0, B=1, ..., Z=25, AA=26, ...）
        let col = col_str
            .chars()
            .rev()
            .enumerate()
            .map(|(i, ch)| {
                let val = (ch as u32) - ('A' as u32) + 1;
                val * 26_u32.pow(i as u32)
            })
            .sum::<u32>()
            - 1;

        // 行を数値に変換（1始まりなので0始まりに変換）
        let row = row_str.parse::<u32>().ok()? - 1;

        Some((row, col))
    }

    /// リレーションシップファイルパスからシート名を抽出
    fn extract_sheet_name_from_rels_path(path: &str) -> String {
        // "xl/worksheets/_rels/sheet1.xml.rels" -> "Sheet1"
        if let Some(name) = path.strip_prefix("xl/worksheets/_rels/sheet") {
            if let Some(num_str) = name.strip_suffix(".xml.rels") {
                if let Ok(num) = num_str.parse::<usize>() {
                    return format!("Sheet{}", num);
                }
            }
        }
        path.to_string()
    }

    /// ファイルパスからシート名を抽出（簡易実装）
    ///
    /// 実際の実装では、workbook.xmlからシート名とファイル名のマッピングを取得すべきですが、
    /// ここでは簡易的にファイル名から推測します。
    fn extract_sheet_name_from_path(path: &str) -> String {
        // "xl/worksheets/sheet1.xml" -> "Sheet1"
        if let Some(name) = path.strip_prefix("xl/worksheets/sheet") {
            if let Some(num_str) = name.strip_suffix(".xml") {
                if let Ok(num) = num_str.parse::<usize>() {
                    return format!("Sheet{}", num);
                }
            }
        }
        // フォールバック: パスをそのまま使用
        path.to_string()
    }

    /// xl/workbook.xml の解析（プライベート）
    ///
    /// `<workbookPr date1904="true"/>` を解析し、1904年エポックフラグを取得します。
    fn parse_workbook<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<bool, XlsxToMdError> {
        let mut workbook_file = match archive.by_name("xl/workbook.xml") {
            Ok(file) => file,
            Err(_) => {
                // workbook.xmlが存在しない場合はデフォルト（false）を返す
                return Ok(false);
            }
        };

        // ZIPファイルの内容を一度メモリに読み込む
        use std::io::Read;
        let mut xml_content = Vec::new();
        workbook_file.read_to_end(&mut xml_content)?;

        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_reader(xml_content.as_slice());
        reader.trim_text(true);

        let mut buf = Vec::new();
        let mut is_1904 = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"workbookPr" {
                        // <workbookPr date1904="true"/>
                        for attr in e.attributes() {
                            let attr = attr.map_err(|e| {
                                XlsxToMdError::Config(format!("XML attribute error: {}", e))
                            })?;
                            if attr.key.as_ref() == b"date1904" {
                                let value_str = std::str::from_utf8(&attr.value)?;
                                is_1904 = value_str == "1" || value_str == "true";
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxToMdError::Config(format!("XML parse error: {}", e))),
                _ => {}
            }
        }

        Ok(is_1904)
    }
}

/// ビルトイン書式ID（0-163）のマッピング
///
/// Excelの標準書式IDとフォーマット文字列の対応表です。
/// このマッピングは、Excelの仕様に基づいています。
fn get_builtin_format(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        5 => Some("$#,##0_);($#,##0)"),
        6 => Some("$#,##0_);[Red]($#,##0)"),
        7 => Some("$#,##0.00_);($#,##0.00)"),
        8 => Some("$#,##0.00_);[Red]($#,##0.00)"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0_);(#,##0)"),
        38 => Some("#,##0_);[Red](#,##0)"),
        39 => Some("#,##0.00_);(#,##0.00)"),
        40 => Some("#,##0.00_);[Red](#,##0.00)"),
        41 => Some("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
        42 => Some("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"),
        43 => Some("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"),
        44 => Some("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mm:ss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None, // その他のビルトイン書式IDは未実装
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_get_builtin_format() {
        assert_eq!(get_builtin_format(0), Some("General"));
        assert_eq!(get_builtin_format(1), Some("0"));
        assert_eq!(get_builtin_format(14), Some("mm-dd-yy"));
        assert_eq!(get_builtin_format(49), Some("@"));
        assert_eq!(get_builtin_format(50), None);
        assert_eq!(get_builtin_format(163), None);
        assert_eq!(get_builtin_format(164), None);
    }

    #[test]
    fn test_extract_sheet_name_from_path() {
        assert_eq!(
            XlsxMetadataParser::extract_sheet_name_from_path("xl/worksheets/sheet1.xml"),
            "Sheet1"
        );
        assert_eq!(
            XlsxMetadataParser::extract_sheet_name_from_path("xl/worksheets/sheet2.xml"),
            "Sheet2"
        );
    }
}
