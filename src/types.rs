//! Types Module
//!
//! クレート全体で使用する共通データ型を定義するモジュール。

/// セルの値を表す列挙型
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum CellValue {
    /// 数値（f64）
    Number(f64),

    /// 文字列
    String(String),

    /// 論理値
    Bool(bool),

    /// エラー値（例: #DIV/0!）
    Error(String),

    /// 空セル
    Empty,
}

impl CellValue {
    /// 値が空かどうかを判定
    #[allow(dead_code)]
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// 値を文字列として取得（書式適用前）
    #[allow(dead_code)]
    pub fn as_raw_string(&self) -> String {
        match self {
            CellValue::Number(n) => n.to_string(),
            CellValue::String(s) => s.clone(),
            CellValue::Bool(b) => b.to_string(),
            CellValue::Error(e) => e.clone(),
            CellValue::Empty => String::new(),
        }
    }
}

/// セル座標（0始まり）
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct CellCoord {
    pub row: u32,
    pub col: u32,
}

impl CellCoord {
    /// 新しい座標を生成
    pub fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }

    /// A1形式の文字列に変換（例: (0, 0) -> "A1"）
    #[allow(dead_code, clippy::wrong_self_convention)]
    pub fn to_a1_notation(&self) -> String {
        let col_str = Self::col_index_to_letter(self.col);
        format!("{}{}", col_str, self.row + 1)
    }

    /// 列インデックスを文字列に変換（0 -> "A", 25 -> "Z", 26 -> "AA"）
    #[allow(dead_code)]
    fn col_index_to_letter(mut col: u32) -> String {
        let mut result = String::new();
        loop {
            let remainder = col % 26;
            result.insert(0, (b'A' + remainder as u8) as char);
            if col < 26 {
                break;
            }
            col = col / 26 - 1;
        }
        result
    }
}

/// セル範囲
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CellRange {
    pub start: CellCoord,
    pub end: CellCoord,
}

impl CellRange {
    /// 新しい範囲を生成
    pub fn new(start: CellCoord, end: CellCoord) -> Self {
        Self { start, end }
    }

    /// 指定された座標が範囲内にあるかを判定
    #[allow(dead_code)]
    pub fn contains(&self, coord: CellCoord) -> bool {
        coord.row >= self.start.row
            && coord.row <= self.end.row
            && coord.col >= self.start.col
            && coord.col <= self.end.col
    }

    /// 範囲のサイズ（行数 × 列数）を計算
    #[allow(dead_code)]
    pub fn size(&self) -> (u32, u32) {
        let rows = self.end.row - self.start.row + 1;
        let cols = self.end.col - self.start.col + 1;
        (rows, cols)
    }
}

/// セル結合範囲の情報
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct MergedRegion {
    /// 結合範囲
    pub range: CellRange,

    /// 親セル（左上セル）の座標
    pub parent: CellCoord,
}

impl MergedRegion {
    /// 新しい結合範囲を生成
    pub fn new(range: CellRange) -> Self {
        Self {
            parent: range.start,
            range,
        }
    }

    /// 指定された座標が結合範囲内にあるかを判定
    #[allow(dead_code)]
    pub fn contains(&self, coord: CellCoord) -> bool {
        self.range.contains(coord)
    }

    /// 結合セルの行数
    pub fn row_span(&self) -> u32 {
        self.range.end.row - self.range.start.row + 1
    }

    /// 結合セルの列数
    pub fn col_span(&self) -> u32 {
        self.range.end.col - self.range.start.col + 1
    }
}

/// リッチテキストの書式情報
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct RichTextFormat {
    /// 太字かどうか
    pub bold: bool,
    /// 斜体かどうか
    pub italic: bool,
}

impl RichTextFormat {
    /// 新しいRichTextFormatを生成
    pub fn new() -> Self {
        Self {
            bold: false,
            italic: false,
        }
    }
}

impl Default for RichTextFormat {
    fn default() -> Self {
        Self::new()
    }
}

/// リッチテキストのセグメント（書式付きテキスト）
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct RichTextSegment {
    /// テキスト内容
    pub text: String,
    /// 書式情報
    pub format: RichTextFormat,
}

impl RichTextSegment {
    /// 新しいRichTextSegmentを生成
    pub fn new(text: String, format: RichTextFormat) -> Self {
        Self { text, format }
    }

    /// 通常のテキストセグメントを生成（書式なし）
    pub fn plain(text: String) -> Self {
        Self {
            text,
            format: RichTextFormat::new(),
        }
    }
}

/// パーサーから抽出された生のセルデータ
#[derive(Debug, Clone)]
pub(crate) struct RawCellData {
    /// セル座標
    pub coord: CellCoord,

    /// セルの値
    pub value: CellValue,

    /// 数値書式ID（calamineから取得）
    pub format_id: Option<u16>,

    /// カスタム書式文字列（存在する場合）
    pub format_string: Option<String>,

    /// 数式文字列（数式セルの場合）
    pub formula: Option<String>,

    /// ハイパーリンク情報（存在する場合）
    pub hyperlink: Option<String>,

    /// リッチテキスト情報（存在する場合）
    /// リッチテキストが存在する場合、valueはStringで通常のテキストが格納される
    pub rich_text: Option<Vec<RichTextSegment>>,
}

/// シートのメタデータ
#[derive(Debug, Clone)]
pub(crate) struct SheetMetadata {
    /// シート名
    #[allow(dead_code)]
    pub name: String,

    /// シートインデックス（0始まり）
    #[allow(dead_code)]
    pub index: usize,

    /// シートが非表示かどうか
    #[allow(dead_code)]
    pub hidden: bool,

    /// セル結合範囲のリスト
    pub merged_regions: Vec<MergedRegion>,

    /// 非表示行のインデックスリスト
    /// Phase I: 空リスト（Phase IIで実装）
    pub hidden_rows: Vec<u32>,

    /// 非表示列のインデックスリスト
    /// Phase I: 空リスト（Phase IIで実装）
    pub hidden_cols: Vec<u32>,

    /// 1904年エポックを使用するか（ワークブック全体の設定）
    /// Phase I: 常にfalse（Phase IIで実装）
    pub is_1904: bool,
}

#[cfg(test)]
mod tests {
    use super::*;

    // CellValue のテスト
    #[test]
    fn test_cell_value_is_empty() {
        assert!(CellValue::Empty.is_empty());
        assert!(!CellValue::Number(42.0).is_empty());
        assert!(!CellValue::String("test".to_string()).is_empty());
        assert!(!CellValue::Bool(true).is_empty());
        assert!(!CellValue::Error("#DIV/0!".to_string()).is_empty());
    }

    #[test]
    fn test_cell_value_as_raw_string() {
        assert_eq!(CellValue::Empty.as_raw_string(), "");
        assert_eq!(CellValue::Number(42.5).as_raw_string(), "42.5");
        assert_eq!(
            CellValue::String("hello".to_string()).as_raw_string(),
            "hello"
        );
        assert_eq!(CellValue::Bool(true).as_raw_string(), "true");
        assert_eq!(
            CellValue::Error("#DIV/0!".to_string()).as_raw_string(),
            "#DIV/0!"
        );
    }

    // CellCoord のテスト
    #[test]
    fn test_cell_coord_new() {
        let coord = CellCoord::new(0, 0);
        assert_eq!(coord.row, 0);
        assert_eq!(coord.col, 0);
    }

    #[test]
    fn test_cell_coord_to_a1_notation() {
        assert_eq!(CellCoord::new(0, 0).to_a1_notation(), "A1");
        assert_eq!(CellCoord::new(0, 25).to_a1_notation(), "Z1");
        assert_eq!(CellCoord::new(0, 26).to_a1_notation(), "AA1");
        assert_eq!(CellCoord::new(99, 701).to_a1_notation(), "ZZ100");
        assert_eq!(CellCoord::new(0, 51).to_a1_notation(), "AZ1");
        assert_eq!(CellCoord::new(0, 52).to_a1_notation(), "BA1");
        assert_eq!(CellCoord::new(0, 701).to_a1_notation(), "ZZ1");
    }

    #[test]
    fn test_cell_coord_col_index_to_letter() {
        // プライベートメソッドのテストは、公開メソッドを通じて間接的にテスト
        assert_eq!(CellCoord::new(0, 0).to_a1_notation(), "A1");
        assert_eq!(CellCoord::new(0, 25).to_a1_notation(), "Z1");
        assert_eq!(CellCoord::new(0, 26).to_a1_notation(), "AA1");
    }

    // CellRange のテスト
    #[test]
    fn test_cell_range_new() {
        let start = CellCoord::new(0, 0);
        let end = CellCoord::new(10, 5);
        let range = CellRange::new(start, end);
        assert_eq!(range.start, start);
        assert_eq!(range.end, end);
    }

    #[test]
    fn test_cell_range_contains() {
        let range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(10, 5));

        // 範囲内の座標
        assert!(range.contains(CellCoord::new(0, 0)));
        assert!(range.contains(CellCoord::new(5, 3)));
        assert!(range.contains(CellCoord::new(10, 5)));

        // 範囲外の座標
        assert!(!range.contains(CellCoord::new(11, 5)));
        assert!(!range.contains(CellCoord::new(5, 6)));
        assert!(!range.contains(CellCoord::new(0, 6)));
    }

    #[test]
    fn test_cell_range_size() {
        let range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(10, 5));
        assert_eq!(range.size(), (11, 6));

        let range2 = CellRange::new(CellCoord::new(5, 3), CellCoord::new(7, 4));
        assert_eq!(range2.size(), (3, 2));

        // 単一セル
        let range3 = CellRange::new(CellCoord::new(0, 0), CellCoord::new(0, 0));
        assert_eq!(range3.size(), (1, 1));
    }

    // MergedRegion のテスト
    #[test]
    fn test_merged_region_new() {
        let range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(2, 3));
        let merged = MergedRegion::new(range);
        assert_eq!(merged.range, range);
        assert_eq!(merged.parent, CellCoord::new(0, 0));
    }

    #[test]
    fn test_merged_region_contains() {
        let range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(2, 3));
        let merged = MergedRegion::new(range);

        // 範囲内の座標
        assert!(merged.contains(CellCoord::new(0, 0)));
        assert!(merged.contains(CellCoord::new(1, 2)));
        assert!(merged.contains(CellCoord::new(2, 3)));

        // 範囲外の座標
        assert!(!merged.contains(CellCoord::new(3, 3)));
        assert!(!merged.contains(CellCoord::new(1, 4)));
    }

    #[test]
    fn test_merged_region_row_span() {
        let range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(2, 3));
        let merged = MergedRegion::new(range);
        assert_eq!(merged.row_span(), 3);

        let range2 = CellRange::new(CellCoord::new(5, 1), CellCoord::new(5, 1));
        let merged2 = MergedRegion::new(range2);
        assert_eq!(merged2.row_span(), 1);
    }

    #[test]
    fn test_merged_region_col_span() {
        let range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(2, 3));
        let merged = MergedRegion::new(range);
        assert_eq!(merged.col_span(), 4);

        let range2 = CellRange::new(CellCoord::new(5, 1), CellCoord::new(5, 1));
        let merged2 = MergedRegion::new(range2);
        assert_eq!(merged2.col_span(), 1);
    }

    // RawCellData のテスト
    #[test]
    fn test_raw_cell_data() {
        let coord = CellCoord::new(0, 0);
        let value = CellValue::Number(42.0);
        let cell_data = RawCellData {
            coord,
            value: value.clone(),
            format_id: Some(1),
            format_string: Some("0.00".to_string()),
            formula: None,
            hyperlink: None,
            rich_text: None,
        };

        assert_eq!(cell_data.coord, coord);
        assert_eq!(cell_data.value, value);
        assert_eq!(cell_data.format_id, Some(1));
        assert_eq!(cell_data.format_string, Some("0.00".to_string()));
        assert_eq!(cell_data.formula, None);
    }

    #[test]
    fn test_raw_cell_data_with_formula() {
        let coord = CellCoord::new(1, 1);
        let value = CellValue::Number(100.0);
        let cell_data = RawCellData {
            coord,
            value: value.clone(),
            format_id: None,
            format_string: None,
            formula: Some("=A1*2".to_string()),
            hyperlink: None,
            rich_text: None,
        };

        assert_eq!(cell_data.formula, Some("=A1*2".to_string()));
    }

    // SheetMetadata のテスト
    #[test]
    fn test_sheet_metadata_phase_i() {
        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![],
            hidden_rows: vec![], // Phase I: 空リスト
            hidden_cols: vec![], // Phase I: 空リスト
            is_1904: false,      // Phase I: 常にfalse
        };

        assert_eq!(metadata.name, "Sheet1");
        assert_eq!(metadata.index, 0);
        assert!(!metadata.hidden);
        assert!(metadata.merged_regions.is_empty());
        assert!(metadata.hidden_rows.is_empty());
        assert!(metadata.hidden_cols.is_empty());
        assert!(!metadata.is_1904);
    }

    #[test]
    fn test_sheet_metadata_with_merged_regions() {
        let range1 = CellRange::new(CellCoord::new(0, 0), CellCoord::new(0, 2));
        let range2 = CellRange::new(CellCoord::new(2, 0), CellCoord::new(3, 1));
        let merged1 = MergedRegion::new(range1);
        let merged2 = MergedRegion::new(range2);

        let metadata = SheetMetadata {
            name: "Sheet1".to_string(),
            index: 0,
            hidden: false,
            merged_regions: vec![merged1.clone(), merged2.clone()],
            hidden_rows: vec![],
            hidden_cols: vec![],
            is_1904: false,
        };

        assert_eq!(metadata.merged_regions.len(), 2);
        assert_eq!(metadata.merged_regions[0], merged1);
        assert_eq!(metadata.merged_regions[1], merged2);
    }

    // プロパティベーステスト: TC-PBT-001
    #[allow(unused_doc_comments)]
    mod property_tests {
        use super::*;
        use proptest::prelude::*;

        #[allow(unused_doc_comments)]
        /// TC-PBT-001: A1 Notation Round-Trip
        ///
        /// ランダムな座標値でA1記法に変換し、形式を検証します。
        /// 注意: A1記法のパース関数がまだ実装されていないため、
        /// 完全なround-tripテストではなく、形式検証のみを行います。
        proptest! {
            #[test]
            fn test_a1_notation_round_trip(row in 0u32..10000, col in 0u32..10000) {
                let coord = CellCoord::new(row, col);
                let a1 = coord.to_a1_notation();

                // A1記法の形式検証
                // 1. 最初の文字が大文字のアルファベットであること
                prop_assert!(a1.chars().next().unwrap().is_ascii_uppercase());

                // 2. 最後の文字が数字であること
                prop_assert!(a1.chars().last().unwrap().is_ascii_digit());

                // 3. 列部分（アルファベット）と行部分（数字）が分離されていること
                // アルファベット部分と数字部分の境界を確認
                let mut found_digit = false;
                for (i, ch) in a1.chars().enumerate() {
                    if ch.is_ascii_digit() {
                        found_digit = true;
                        // 数字部分の前はすべてアルファベットであること
                        prop_assert!(i > 0, "A1 notation should have at least one letter");
                    } else {
                        // 数字が見つかった後はすべて数字であること
                        prop_assert!(!found_digit, "A1 notation should not have letters after digits");
                    }
                }

                // 4. 空でないこと
                prop_assert!(!a1.is_empty());

                // 5. 行番号が1以上であること（0始まりの座標を1始まりに変換）
                let row_part: String = a1.chars().filter(|c| c.is_ascii_digit()).collect();
                let row_num: u32 = row_part.parse().unwrap();
                prop_assert!(row_num >= 1);
                prop_assert_eq!(row_num, row + 1);
            }
        }
    }
}
