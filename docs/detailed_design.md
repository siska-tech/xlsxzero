# **詳細設計書 (Detailed Design Document)**

## **目的**

本設計書は、[architecture.md](architecture.md)で定義されたアーキテクチャと[interface.md](interface.md)で定義された公開APIの内部実装を具体化する。開発者が直接コーディングする際の指針となる詳細な仕様を提供する。

**主要な目標:**
* 各モジュールの内部実装（プライベートメソッド、内部データ構造）を明確化
* 複雑なアルゴリズムやビジネスロジックの処理手順を定義
* エラーハンドリングの詳細な実装方針を提示
* 開発者が迷わず実装できる具体的な指示を提供

**関連文書:**
* [requirements.md](requirements.md): 要件定義
* [architecture.md](architecture.md): アーキテクチャ設計
* [interface.md](interface.md): 公開API設計

---

## **1. Builder Module (`builder.rs`)**

### **1.1. 内部データ構造**

#### **ConversionConfig構造体**

**目的:** ビルダーで構築された設定を保持する内部構造体。

```rust
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
}
```

**デフォルト実装:**
```rust
impl Default for ConversionConfig {
    fn default() -> Self {
        Self {
            sheet_selector: SheetSelector::All,
            merge_strategy: MergeStrategy::DataDuplication,
            date_format: DateFormat::Iso8601,
            formula_mode: FormulaMode::CachedValue,
            include_hidden: false,
            range: None,
        }
    }
}
```

---

### **1.2. ConverterBuilder構造体**

**内部状態:**
```rust
/// Fluent Builder APIを提供する構造体
#[derive(Debug)]
pub struct ConverterBuilder {
    /// 内部設定（構築中）
    config: ConversionConfig,
}
```

---

### **1.3. 設定検証ロジック**

#### **build()メソッドの内部実装**

**シグネチャ:**
```rust
pub fn build(self) -> Result<Converter, XlsxToMdError>
```

**処理フロー:**
```
1. セル範囲の妥当性検証
   ├─ range.is_some() の場合
   │  ├─ start.row <= end.row をチェック
   │  ├─ start.col <= end.col をチェック
   │  └─ 違反があれば Err(XlsxToMdError::Config) を返す
   └─ range.is_none() の場合は検証スキップ

2. カスタム日付形式の検証
   ├─ DateFormat::Custom(format_str) の場合
   │  ├─ chrono::format::strftime でパース試行
   │  └─ 失敗した場合は Err(XlsxToMdError::Config) を返す
   └─ DateFormat::Iso8601 の場合は検証スキップ

3. Converterインスタンスの生成
   └─ Ok(Converter::new(self.config))
```

**疑似コード:**
```rust
pub fn build(self) -> Result<Converter, XlsxToMdError> {
    // 1. セル範囲の検証
    if let Some(range) = &self.config.range {
        if range.start.row > range.end.row {
            return Err(XlsxToMdError::Config(
                format!(
                    "Invalid range: start row ({}) > end row ({})",
                    range.start.row, range.end.row
                )
            ));
        }

        if range.start.col > range.end.col {
            return Err(XlsxToMdError::Config(
                format!(
                    "Invalid range: start col ({}) > end col ({})",
                    range.start.col, range.end.col
                )
            ));
        }
    }

    // 2. カスタム日付形式の検証
    if let DateFormat::Custom(ref format_str) = self.config.date_format {
        // テスト用の日付でフォーマット試行
        let test_date = NaiveDate::from_ymd_opt(2025, 1, 1).unwrap();
        if test_date.format(format_str).to_string().is_empty() {
            return Err(XlsxToMdError::Config(
                format!("Invalid date format string: '{}'", format_str)
            ));
        }
    }

    // 3. Converterインスタンス生成
    Ok(Converter::new(self.config))
}
```

---

## **2. Types Module (`types.rs`)**

### **2.1. 基本データ型**

#### **CellValue列挙型**

```rust
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
```

**実装メソッド:**
```rust
impl CellValue {
    /// 値が空かどうかを判定
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// 値を文字列として取得（書式適用前）
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
```

---

#### **CellCoord構造体**

```rust
/// セル座標（0始まり）
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct CellCoord {
    pub row: u32,
    pub col: u32,
}
```

**実装メソッド:**
```rust
impl CellCoord {
    /// 新しい座標を生成
    pub fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }

    /// A1形式の文字列に変換（例: (0, 0) -> "A1"）
    pub fn to_a1_notation(&self) -> String {
        let col_str = Self::col_index_to_letter(self.col);
        format!("{}{}", col_str, self.row + 1)
    }

    /// 列インデックスを文字列に変換（0 -> "A", 25 -> "Z", 26 -> "AA"）
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
```

**テストケース:**
```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_to_a1_notation() {
        assert_eq!(CellCoord::new(0, 0).to_a1_notation(), "A1");
        assert_eq!(CellCoord::new(0, 25).to_a1_notation(), "Z1");
        assert_eq!(CellCoord::new(0, 26).to_a1_notation(), "AA1");
        assert_eq!(CellCoord::new(99, 701).to_a1_notation(), "ZZ100");
    }
}
```

---

#### **CellRange構造体**

```rust
/// セル範囲
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CellRange {
    pub start: CellCoord,
    pub end: CellCoord,
}
```

**実装メソッド:**
```rust
impl CellRange {
    /// 新しい範囲を生成
    pub fn new(start: CellCoord, end: CellCoord) -> Self {
        Self { start, end }
    }

    /// 指定された座標が範囲内にあるかを判定
    pub fn contains(&self, coord: CellCoord) -> bool {
        coord.row >= self.start.row
            && coord.row <= self.end.row
            && coord.col >= self.start.col
            && coord.col <= self.end.col
    }

    /// 範囲のサイズ（行数 × 列数）を計算
    pub fn size(&self) -> (u32, u32) {
        let rows = self.end.row - self.start.row + 1;
        let cols = self.end.col - self.start.col + 1;
        (rows, cols)
    }
}
```

---

#### **MergedRegion構造体**

```rust
/// セル結合範囲の情報
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct MergedRegion {
    /// 結合範囲
    pub range: CellRange,

    /// 親セル（左上セル）の座標
    pub parent: CellCoord,
}
```

**実装メソッド:**
```rust
impl MergedRegion {
    /// 新しい結合範囲を生成
    pub fn new(range: CellRange) -> Self {
        Self {
            parent: range.start,
            range,
        }
    }

    /// 指定された座標が結合範囲内にあるかを判定
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
```

---

### **2.2. パーサー層のデータ型**

#### **RawCellData構造体**

```rust
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
}
```

---

#### **SheetMetadata構造体**

```rust
/// シートのメタデータ
#[derive(Debug, Clone)]
pub(crate) struct SheetMetadata {
    /// シート名
    pub name: String,

    /// シートインデックス（0始まり）
    pub index: usize,

    /// シートが非表示かどうか
    pub hidden: bool,

    /// セル結合範囲のリスト
    pub merged_regions: Vec<MergedRegion>,

    /// 非表示行のインデックスリスト
    pub hidden_rows: Vec<u32>,

    /// 非表示列のインデックスリスト
    pub hidden_cols: Vec<u32>,

    /// 1904年エポックを使用するか（ワークブック全体の設定）
    /// Phase II: Issue #14で実装
    pub is_1904: bool,
}
```

---

## **3. Parser Module (`parser.rs`)**

### **3.1. WorkbookParser構造体**

**目的:** calamineのラッパーとして、ワークブックレベルの操作を提供。

```rust
/// ワークブックパーサー
pub(crate) struct WorkbookParser<R: Read + Seek> {
    /// calamineのワークブック
    workbook: Xlsx<R>,
}
```

---

### **3.2. 主要メソッドの実装**

#### **open()メソッド**

**シグネチャ:**
```rust
pub fn open<R: Read + Seek>(reader: R) -> Result<Self, XlsxToMdError>
```

**処理フロー:**
```
1. calamine::open_workbook_auto() でワークブックを開く
   └─ 失敗した場合は Err(XlsxToMdError::Parse) を返す

2. WorkbookParser インスタンスを生成
   └─ Ok(Self { workbook })
```

**実装:**
```rust
pub fn open<R: Read + Seek>(reader: R) -> Result<Self, XlsxToMdError> {
    let workbook = calamine::open_workbook_auto(reader)
        .map_err(XlsxToMdError::Parse)?;

    Ok(Self { workbook })
}
```

---

#### **get_sheet_names()メソッド**

**シグネチャ:**
```rust
pub fn get_sheet_names(&self) -> Vec<String>
```

**処理フロー:**
```
1. workbook.sheet_names() を呼び出し
2. Vec<String> として返却
```

**実装:**
```rust
pub fn get_sheet_names(&self) -> Vec<String> {
    self.workbook.sheet_names().to_vec()
}
```

---

#### **select_sheets()メソッド**

**シグネチャ:**
```rust
pub fn select_sheets(
    &mut self,
    selector: &SheetSelector,
    include_hidden: bool,
) -> Result<Vec<String>, XlsxToMdError>
```

**処理フロー:**
```
1. SheetSelector のバリアントに応じて処理を分岐
   ├─ SheetSelector::All
   │  ├─ すべてのシート名を取得
   │  └─ include_hidden が false の場合、非表示シートをフィルタ
   │
   ├─ SheetSelector::Index(index)
   │  ├─ インデックスの範囲チェック
   │  ├─ 範囲外なら Err(XlsxToMdError::Config)
   │  └─ シート名を返却
   │
   ├─ SheetSelector::Name(name)
   │  ├─ シート名の存在チェック
   │  ├─ 存在しなければ Err(XlsxToMdError::Config)
   │  └─ シート名を返却
   │
   ├─ SheetSelector::Indices(indices)
   │  ├─ 各インデックスの範囲チェック
   │  └─ シート名のリストを返却
   │
   └─ SheetSelector::Names(names)
      ├─ 各シート名の存在チェック
      └─ シート名のリストを返却

2. 選択されたシート名のリストを返却
```

**実装:**
```rust
pub fn select_sheets(
    &mut self,
    selector: &SheetSelector,
    include_hidden: bool,
) -> Result<Vec<String>, XlsxToMdError> {
    let all_sheet_names = self.get_sheet_names();

    match selector {
        SheetSelector::All => {
            if include_hidden {
                Ok(all_sheet_names)
            } else {
                // 非表示シートをフィルタ
                Ok(all_sheet_names
                    .into_iter()
                    .filter(|name| !self.is_sheet_hidden(name))
                    .collect())
            }
        }

        SheetSelector::Index(index) => {
            if *index >= all_sheet_names.len() {
                return Err(XlsxToMdError::Config(
                    format!(
                        "Sheet index {} is out of range (total: {})",
                        index, all_sheet_names.len()
                    )
                ));
            }
            Ok(vec![all_sheet_names[*index].clone()])
        }

        SheetSelector::Name(name) => {
            if !all_sheet_names.contains(name) {
                return Err(XlsxToMdError::Config(
                    format!("Sheet '{}' not found", name)
                ));
            }
            Ok(vec![name.clone()])
        }

        SheetSelector::Indices(indices) => {
            let mut result = Vec::new();
            for &index in indices {
                if index >= all_sheet_names.len() {
                    return Err(XlsxToMdError::Config(
                        format!(
                            "Sheet index {} is out of range (total: {})",
                            index, all_sheet_names.len()
                        )
                    ));
                }
                result.push(all_sheet_names[index].clone());
            }
            Ok(result)
        }

        SheetSelector::Names(names) => {
            for name in names {
                if !all_sheet_names.contains(name) {
                    return Err(XlsxToMdError::Config(
                        format!("Sheet '{}' not found", name)
                    ));
                }
            }
            Ok(names.clone())
        }
    }
}
```

---

#### **parse_sheet()メソッド**

**シグネチャ:**
```rust
pub fn parse_sheet(
    &mut self,
    sheet_name: &str,
    config: &ConversionConfig,
) -> Result<(SheetMetadata, Vec<RawCellData>), XlsxToMdError>
```

**処理フロー:**
```
1. シートの取得
   ├─ workbook.worksheet_range(sheet_name) を呼び出し
   └─ 失敗した場合は Err(XlsxToMdError::Parse)

2. メタデータの収集
   ├─ シート名、インデックスを取得
   ├─ 結合セル範囲を取得（merged_regions()）
   ├─ 非表示行・列のリストを取得
   └─ SheetMetadata を構築

3. セルデータの抽出（ストリーミング処理）
   ├─ range.rows() でイテレート
   ├─ 各行について
   │  ├─ config.include_hidden が false の場合、非表示行をスキップ
   │  └─ 各セルについて
   │     ├─ config.include_hidden が false の場合、非表示列をスキップ
   │     ├─ config.range が指定されている場合、範囲外をスキップ
   │     ├─ RawCellData を生成
   │     └─ Vec<RawCellData> に追加
   └─ セルデータのリストを返却

4. (SheetMetadata, Vec<RawCellData>) を返却
```

**実装:**
```rust
pub fn parse_sheet(
    &mut self,
    sheet_name: &str,
    config: &ConversionConfig,
) -> Result<(SheetMetadata, Vec<RawCellData>), XlsxToMdError> {
    // 1. シートの取得
    let range = self.workbook
        .worksheet_range(sheet_name)
        .ok_or_else(|| XlsxToMdError::Parse(
            calamine::Error::Msg(format!("Sheet '{}' not found", sheet_name))
        ))??;

    // 2. メタデータの収集
    let metadata = self.collect_metadata(sheet_name)?;

    // 3. セルデータの抽出
    let mut cells = Vec::new();

    for (row_idx, row) in range.rows().enumerate() {
        let row_idx = row_idx as u32;

        // 非表示行のスキップ
        if !config.include_hidden && metadata.hidden_rows.contains(&row_idx) {
            continue;
        }

        for (col_idx, cell) in row.iter().enumerate() {
            let col_idx = col_idx as u32;

            // 非表示列のスキップ
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
            let raw_cell = self.extract_cell_data(coord, cell)?;
            cells.push(raw_cell);
        }
    }

    Ok((metadata, cells))
}
```

---

#### **extract_cell_data()メソッド（内部ヘルパー）**

**シグネチャ:**
```rust
fn extract_cell_data(
    &self,
    coord: CellCoord,
    cell: &calamine::DataType,
) -> Result<RawCellData, XlsxToMdError>
```

**処理フロー:**
```
1. calamine::DataType から CellValue への変換
   ├─ DataType::Int(i) -> CellValue::Number(i as f64)
   ├─ DataType::Float(f) -> CellValue::Number(f)
   ├─ DataType::String(s) -> CellValue::String(s)
   ├─ DataType::Bool(b) -> CellValue::Bool(b)
   ├─ DataType::Error(e) -> CellValue::Error(e)
   └─ DataType::Empty -> CellValue::Empty

2. 書式情報の取得（可能な場合）
   ├─ format_id を取得
   └─ format_string を取得

3. 数式情報の取得（数式セルの場合）
   └─ formula を取得

4. RawCellData を生成して返却
```

**実装:**
```rust
fn extract_cell_data(
    &self,
    coord: CellCoord,
    cell: &calamine::DataType,
) -> Result<RawCellData, XlsxToMdError> {
    use calamine::DataType;

    // 1. 値の変換
    let value = match cell {
        DataType::Int(i) => CellValue::Number(*i as f64),
        DataType::Float(f) => CellValue::Number(*f),
        DataType::String(s) => CellValue::String(s.clone()),
        DataType::Bool(b) => CellValue::Bool(*b),
        DataType::Error(e) => CellValue::Error(format!("{:?}", e)),
        DataType::Empty => CellValue::Empty,
        _ => CellValue::Empty,
    };

    // 2. 書式情報の取得
    // Phase I: calamine APIでは書式情報が取得不可（公式未サポート）
    // Phase II: XlsxMetadataParserでxl/styles.xmlから取得予定
    let format_id = None; // Phase I: None固定
    let format_string = None; // Phase I: None固定

    // 3. 数式情報の取得
    // Phase I: calamine 0.32.0のworksheet_formula() APIで取得可能
    let formula = match workbook.worksheet_formula(sheet_name) {
        Ok(range) => range.get((coord.row, coord.col)).map(|s| s.to_string()),
        Err(_) => None,
    };

    Ok(RawCellData {
        coord,
        value,
        format_id,
        format_string,
        formula,
    })
}
```

---

### **3.3. メタデータ収集の実装**

#### **collect_metadata()メソッド**

**シグネチャ:**
```rust
fn collect_metadata(
    &mut self,
    sheet_name: &str,
) -> Result<SheetMetadata, XlsxToMdError>
```

**処理フロー:**
```
1. シートインデックスの取得
   └─ sheet_names から検索

2. 非表示フラグの取得
   └─ workbook のメタデータから取得（API依存）

3. 結合セル範囲の取得
   ├─ workbook.merged_regions(sheet_name) を呼び出し
   └─ Vec<MergedRegion> に変換

4. 非表示行・列のリストを取得
   └─ workbook のメタデータから取得（API依存）

5. SheetMetadata を返却
```

**実装:**
```rust
fn collect_metadata(
    &mut self,
    sheet_name: &str,
) -> Result<SheetMetadata, XlsxToMdError> {
    // 1. シートインデックスの取得
    let index = self.workbook
        .sheet_names()
        .iter()
        .position(|name| name == sheet_name)
        .ok_or_else(|| XlsxToMdError::Config(
            format!("Sheet '{}' not found", sheet_name)
        ))?;

    // 2. 非表示フラグの取得
    // Phase I: calamine APIで非表示シート情報は未サポート
    // Phase II: XlsxMetadataParserで取得予定
    let hidden = false; // Phase I: false固定

    // 3. 結合セル範囲の取得
    // Phase I: calamine 0.32.0で完全対応
    workbook.load_merged_regions()?;
    let merged_regions = workbook.worksheet_merge_cells(sheet_name)
        .unwrap_or_default()
        .iter()
        .map(|dims| MergedRegion {
            start: CellCoord::new(dims.start.0, dims.start.1),
            end: CellCoord::new(dims.end.0, dims.end.1),
        })
        .collect();

    // 4. 非表示行・列のリスト
    // Phase I: calamine APIで未サポート（GitHub Issue #237）
    // Phase II: XlsxMetadataParserでxl/worksheets/*.xmlから取得予定
    let hidden_rows = Vec::new(); // Phase I: 空リスト（全行を処理）
    let hidden_cols = Vec::new(); // Phase I: 空リスト（全列を処理）

    Ok(SheetMetadata {
        name: sheet_name.to_string(),
        index,
        hidden,
        merged_regions,
        hidden_rows,
        hidden_cols,
    })
}
```

**注意:** calamineの現在のバージョンでは、結合セル範囲や非表示要素へのアクセスが制限されている。将来的なAPI拡張を見越して、このような構造で設計する。

---

## **4. Formatter Module (`formatter.rs`)**

### **4.1. CellFormatter構造体**

**目的:** セル値のフォーマット処理のファサード。

```rust
/// セルフォーマッター
pub(crate) struct CellFormatter {
    /// 日付フォーマッター
    date_formatter: DateFormatter,

    /// 数値フォーマッター
    number_formatter: NumberFormatter,
}
```

---

### **4.2. 主要メソッドの実装**

#### **format_cell()メソッド**

**シグネチャ:**
```rust
pub fn format_cell(
    &self,
    raw_cell: &RawCellData,
    config: &ConversionConfig,
) -> Result<String, XlsxToMdError>
```

**処理フロー:**
```
1. 数式モードの処理
   ├─ FormulaMode::Formula かつ raw_cell.formula.is_some()
   │  └─ 数式文字列をそのまま返却
   └─ FormulaMode::CachedValue
      └─ 次のステップへ

2. 値の種類に応じてフォーマット処理を分岐
   ├─ CellValue::Number(n)
   │  ├─ 日付判定（is_date_value()）
   │  │  ├─ true: date_formatter.format() で日付変換
   │  │  └─ false: number_formatter.format() で数値フォーマット
   │  └─ フォーマット済み文字列を返却
   │
   ├─ CellValue::String(s)
   │  └─ そのまま返却（エスケープ処理を適用）
   │
   ├─ CellValue::Bool(b)
   │  └─ "TRUE" / "FALSE" として返却
   │
   ├─ CellValue::Error(e)
   │  └─ エラー文字列をそのまま返却
   │
   └─ CellValue::Empty
      └─ 空文字列を返却

3. フォーマット済み文字列を返却
```

**実装:**
```rust
pub fn format_cell(
    &self,
    raw_cell: &RawCellData,
    config: &ConversionConfig,
) -> Result<String, XlsxToMdError> {
    // 1. 数式モードの処理
    if config.formula_mode == FormulaMode::Formula {
        if let Some(ref formula) = raw_cell.formula {
            return Ok(formula.clone());
        }
    }

    // 2. 値の種類に応じてフォーマット
    match &raw_cell.value {
        CellValue::Number(n) => {
            // 日付判定
            if self.is_date_value(*n, &raw_cell.format_id, &raw_cell.format_string) {
                self.date_formatter.format(*n, config)
            } else {
                self.number_formatter.format(*n, &raw_cell.format_string)
            }
        }

        CellValue::String(s) => {
            Ok(self.escape_markdown(s))
        }

        CellValue::Bool(b) => {
            Ok(if *b { "TRUE" } else { "FALSE" }.to_string())
        }

        CellValue::Error(e) => {
            Ok(e.clone())
        }

        CellValue::Empty => {
            Ok(String::new())
        }
    }
}
```

---

#### **is_date_value()メソッド（内部ヘルパー）**

**シグネチャ:**
```rust
fn is_date_value(
    &self,
    value: f64,
    format_id: &Option<u16>,
    format_string: &Option<String>,
) -> bool
```

**処理フロー:**
```
1. format_id が存在する場合
   ├─ 組み込み日付書式ID（14-22, 45-47）をチェック
   └─ マッチした場合は true を返却

2. format_string が存在する場合
   ├─ 日付キーワード（"yy", "mm", "dd", "hh"）を検索
   └─ 見つかった場合は true を返却

3. 値の範囲チェック（ヒューリスティック）
   ├─ 0 < value < 60000 (1900年～2064年の範囲)
   └─ 範囲外の場合は false

4. デフォルトで false を返却
```

**実装:**
```rust
fn is_date_value(
    &self,
    value: f64,
    format_id: &Option<u16>,
    format_string: &Option<String>,
) -> bool {
    // 1. 組み込み日付書式IDのチェック
    if let Some(id) = format_id {
        // Excel組み込み日付書式ID
        // 14: "m/d/yy"
        // 15-17: 日付形式
        // 18-21: 時刻形式
        // 22: "m/d/yy h:mm"
        // 45-47: 追加の日付形式
        if matches!(id, 14..=22 | 45..=47) {
            return true;
        }
    }

    // 2. カスタム書式文字列のチェック
    if let Some(ref format_str) = format_string {
        let format_lower = format_str.to_lowercase();
        if format_lower.contains("yy")
            || format_lower.contains("mm")
            || format_lower.contains("dd")
            || format_lower.contains("hh") {
            return true;
        }
    }

    // 3. 値の範囲チェック
    if value < 0.0 || value > 60000.0 {
        return false;
    }

    false
}
```

---

#### **escape_markdown()メソッド（内部ヘルパー）**

**シグネチャ:**
```rust
fn escape_markdown(&self, s: &str) -> String
```

**処理フロー:**
```
1. Markdown特殊文字のエスケープ
   ├─ '|' -> '\\|'
   ├─ '\n' -> '<br>'
   └─ バックスラッシュ: '\\' -> '\\\\'

2. エスケープ済み文字列を返却
```

**実装:**
```rust
fn escape_markdown(&self, s: &str) -> String {
    s.replace('\\', "\\\\")
        .replace('|', "\\|")
        .replace('\n', "<br>")
}
```

---

### **4.3. DateFormatter構造体**

```rust
/// 日付フォーマッター
pub(crate) struct DateFormatter;
```

#### **format()メソッド**

**シグネチャ:**
```rust
pub fn format(
    &self,
    serial_value: f64,
    config: &ConversionConfig,
) -> Result<String, XlsxToMdError>
```

**処理フロー:**
```
1. Excelシリアル値をNaiveDateTimeに変換
   ├─ Excelエポック: 1900年1月1日
   ├─ serial_value を日数として解釈
   ├─ NaiveDate::from_ymd(1900, 1, 1) + Duration::days(serial_value as i64)
   └─ 変換失敗時は Err(XlsxToMdError::UnsupportedFeature)

2. DateFormat に応じてフォーマット
   ├─ DateFormat::Iso8601
   │  └─ "%Y-%m-%d" でフォーマット
   └─ DateFormat::Custom(format_str)
      └─ カスタムフォーマット文字列でフォーマット

3. フォーマット済み文字列を返却
```

**Phase I実装:**
```rust
pub fn format(
    &self,
    serial_value: f64,
    config: &ConversionConfig,
    // Phase II追加予定: is_1904: bool,
) -> Result<String, XlsxToMdError> {
    use chrono::{Duration, NaiveDate};

    // Phase Iでは常に1900年起算
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();

    // シリアル値からNaiveDateに変換
    let days = serial_value.floor() as i64;
    let date = epoch + Duration::days(days);

    // DateFormatに応じてフォーマット
    let formatted = match &config.date_format {
        DateFormat::Iso8601 => {
            date.format("%Y-%m-%d").to_string()
        }
        DateFormat::Custom(format_str) => {
            date.format(format_str).to_string()
        }
    };

    Ok(formatted)
}
```

**Phase I制限事項:**
- `DateFormatter::format()`は`is_1904`引数を**受け取らない**
- 常に1900年システム（1899年12月30日起算）として処理
- Phase IIで`SheetMetadata::is_1904`を追加し、引数に渡す予定

**Phase II移行計画:**
```rust
// Phase II実装 (Issue #14完了後)
pub fn format(
    &self,
    serial_value: f64,
    config: &ConversionConfig,
    is_1904: bool,  // XmlMetadataParserから取得
) -> Result<String, XlsxToMdError> {
    use chrono::{Duration, NaiveDate};

    // 1. エポックを決定
    // - 1900年システム: 1899年12月30日起算（Excelの1900年うるう年バグを考慮）
    // - 1904年システム: 1904年1月1日起算
    let epoch = if is_1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1).unwrap()
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30).unwrap()
    };

    // 2. シリアル値からNaiveDateに変換
    let days = serial_value.floor() as i64;
    let date = epoch + Duration::days(days);

    // 3. DateFormatに応じてフォーマット
    let formatted = match &config.date_format {
        DateFormat::Iso8601 => {
            date.format("%Y-%m-%d").to_string()
        }
        DateFormat::Custom(format_str) => {
            date.format(format_str).to_string()
        }
    };

    Ok(formatted)
}
```

**注意:**
- Excelには2つの日付システムが存在:
  - **1900年システム** (デフォルト): 1899年12月30日起算。「1900年2月29日問題」（1900年をうるう年として扱う誤り）が存在するため、エポックを1899年12月30日として調整。
  - **1904年システム**: 1904年1月1日起算。主にMac版Excelで使用。
- ワークブックの日付システムは`xl/workbook.xml`の`<workbookPr date1904="true"/>`属性で判定（Phase II: Issue #14で実装）。

---

### **4.4. NumberFormatter構造体**

```rust
/// 数値フォーマッター
pub(crate) struct NumberFormatter;
```

#### **format()メソッド**

**シグネチャ:**
```rust
pub fn format(
    &self,
    value: f64,
    format_string: &Option<String>,
) -> Result<String, XlsxToMdError>
```

**処理フロー:**
```
1. format_string が存在する場合
   ├─ Number Format Parser を使用して解析
   ├─ フォーマットトークンに基づいて数値を変換
   └─ フォーマット済み文字列を返却

2. format_string が存在しない場合
   ├─ デフォルトフォーマット（小数点以下を保持）
   └─ to_string() で変換

3. フォーマット済み文字列を返却
```

**実装:**
```rust
pub fn format(
    &self,
    value: f64,
    format_string: &Option<String>,
) -> Result<String, XlsxToMdError> {
    if let Some(ref format_str) = format_string {
        // Number Format Parser を使用
        let parser = crate::format::FormatParser::new(format_str);
        let formatted = parser.format_number(value)?;
        Ok(formatted)
    } else {
        // デフォルトフォーマット
        Ok(value.to_string())
    }
}
```

---

## **5. Grid Module (`grid.rs`)**

### **5.1. LogicalGrid構造体**

**目的:** スパースなセルデータから稠密なグリッド構造を構築。

```rust
/// 論理的なグリッド構造
pub(crate) struct LogicalGrid {
    /// グリッドデータ（行 × 列）
    cells: Vec<Vec<Cell>>,

    /// 行数
    rows: usize,

    /// 列数
    cols: usize,
}
```

---

### **5.2. Cell構造体**

```rust
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
```

**実装メソッド:**
```rust
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
```

---

### **5.3. 主要メソッドの実装**

#### **build()メソッド**

**シグネチャ:**
```rust
pub fn build(
    cells: Vec<RawCellData>,
    formatted_cells: Vec<(CellCoord, String)>,
    metadata: &SheetMetadata,
    merge_strategy: MergeStrategy,
) -> Result<Self, XlsxToMdError>
```

**処理フロー:**
```
1. グリッドサイズの決定
   ├─ すべてのセル座標から最大行・列を算出
   └─ (rows, cols) を決定

2. 空のグリッドを初期化
   └─ Vec<Vec<Cell>> を rows × cols で生成

3. フォーマット済みセルデータをグリッドに配置
   ├─ formatted_cells をイテレート
   └─ 各セルを対応する座標に配置

4. セル結合の処理
   ├─ merge_strategy に応じて処理を分岐
   ├─ MergeStrategy::DataDuplication
   │  └─ apply_data_duplication() を呼び出し
   └─ MergeStrategy::HtmlFallback
      └─ apply_html_fallback() を呼び出し

5. LogicalGrid インスタンスを返却
```

**実装:**
```rust
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
```

---

#### **determine_grid_size()メソッド（内部ヘルパー）**

**シグネチャ:**
```rust
fn determine_grid_size(cells: &[RawCellData]) -> (usize, usize)
```

**処理フロー:**
```
1. すべてのセル座標をイテレート
   ├─ 最大行インデックスを記録
   └─ 最大列インデックスを記録

2. (max_row + 1, max_col + 1) を返却
```

**実装:**
```rust
fn determine_grid_size(cells: &[RawCellData]) -> (usize, usize) {
    let mut max_row = 0;
    let mut max_col = 0;

    for cell in cells {
        max_row = max_row.max(cell.coord.row);
        max_col = max_col.max(cell.coord.col);
    }

    ((max_row + 1) as usize, (max_col + 1) as usize)
}
```

---

#### **apply_data_duplication()メソッド**

**シグネチャ:**
```rust
fn apply_data_duplication(
    &mut self,
    merged_regions: &[MergedRegion],
) -> Result<(), XlsxToMdError>
```

**処理フロー:**
```
1. すべての結合範囲をイテレート
   ├─ 親セル（左上セル）の内容を取得
   └─ 結合範囲内のすべてのセルに内容を複製
      ├─ Cell::new_merged() を使用
      └─ merge_parent に親座標を設定

2. 完了
```

**実装:**
```rust
fn apply_data_duplication(
    &mut self,
    merged_regions: &[MergedRegion],
) -> Result<(), XlsxToMdError> {
    for region in merged_regions {
        // 親セルの内容を取得
        let parent_content = self.cells
            [region.parent.row as usize]
            [region.parent.col as usize]
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
```

---

#### **render_markdown()メソッド**

**シグネチャ:**
```rust
pub fn render_markdown<W: Write>(&self, writer: &mut W) -> Result<(), XlsxToMdError>
```

**処理フロー:**
```
1. 列幅の計算
   ├─ 各列の最大文字列長を計算
   └─ Vec<usize> として保持

2. ヘッダー区切り行の生成
   └─ "|---|---|...|" 形式

3. 各行をイテレート
   ├─ 行の開始: '|' を出力
   ├─ 各セルについて
   │  ├─ セル内容を出力
   │  ├─ パディング（列幅に合わせて空白を追加）
   │  └─ '|' を出力
   └─ 行の終了: '\n' を出力

4. バッファをフラッシュ
```

**実装:**
```rust
pub fn render_markdown<W: Write>(&self, writer: &mut W) -> Result<(), XlsxToMdError> {
    use std::io::Write;

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
            write!(writer, " {:<width$} |", cell.content, width = width)?;
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
```

---

#### **calculate_column_widths()メソッド（内部ヘルパー）**

**シグネチャ:**
```rust
fn calculate_column_widths(&self) -> Vec<usize>
```

**処理フロー:**
```
1. 各列について、すべての行のセル内容の長さを計算
2. 列ごとの最大長を記録
3. Vec<usize> として返却
```

**実装:**
```rust
fn calculate_column_widths(&self) -> Vec<usize> {
    let mut widths = vec![0; self.cols];

    for row in &self.cells {
        for (col_idx, cell) in row.iter().enumerate() {
            widths[col_idx] = widths[col_idx].max(cell.content.len());
        }
    }

    widths
}
```

---

### **5.4. HTML出力の実装**

#### **render_html()メソッド**

**シグネチャ:**
```rust
pub fn render_html<W: Write>(
    &self,
    writer: &mut W,
    merged_regions: &[MergedRegion]
) -> Result<(), XlsxToMdError>
```

**概要:**
セル結合を含むテーブルをHTML形式で出力する。`MergeStrategy::HtmlFallback`が指定された場合に使用される。

**処理フロー:**
```
1. <table> 開始タグを出力
2. 各行について
   - 結合セルの親の場合、rowspan/colspan属性を計算
   - 結合セルの子の場合、出力をスキップ
   - 通常セルの場合、<td>セル内容</td>を出力
3. </table> 終了タグを出力
```

**実装:**
```rust
pub fn render_html<W: Write>(
    &self,
    writer: &mut W,
    merged_regions: &[MergedRegion]
) -> Result<(), XlsxToMdError> {
    use std::io::Write;

    writeln!(writer, "<table>")?;

    for (row_idx, row) in self.cells.iter().enumerate() {
        writeln!(writer, "  <tr>")?;

        for (col_idx, cell) in row.iter().enumerate() {
            let coord = CellCoord::new(row_idx as u32, col_idx as u32);

            // 結合セルの子かチェック
            if cell.is_merged && cell.merge_parent.is_some() {
                continue;  // スキップ
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

            write!(writer, "{}</td>\n", cell.content)?;
        }

        writeln!(writer, "  </tr>")?;
    }

    writeln!(writer, "</table>")?;
    writer.flush()?;
    Ok(())
}
```

---

#### **calculate_span()メソッド（内部ヘルパー）**

**シグネチャ:**
```rust
fn calculate_span(
    &self,
    coord: &CellCoord,
    merged_regions: &[MergedRegion]
) -> (u32, u32)
```

**処理フロー:**
```
1. 指定されたセル座標が結合セルの親かチェック
2. 親の場合、row_span()とcol_span()を返却
3. 親でない場合、(1, 1)を返却
```

**実装:**
```rust
fn calculate_span(
    &self,
    coord: &CellCoord,
    merged_regions: &[MergedRegion]
) -> (u32, u32) {
    for region in merged_regions {
        if region.parent == *coord {
            return (region.row_span(), region.col_span());
        }
    }
    (1, 1)
}
```

**出力例:**
```html
<table>
  <tr>
    <th colspan="3">Header</th>
  </tr>
  <tr>
    <td>Data1</td>
    <td>Data2</td>
    <td>Data3</td>
  </tr>
</table>
```

**Converterモジュールでの使用:**
```rust
// converter.rs内
match self.config.merge_strategy {
    MergeStrategy::DataDuplication => {
        grid.render_markdown(&mut writer)?;
    }
    MergeStrategy::HtmlFallback => {
        grid.render_html(&mut writer, &metadata.merged_regions)?;
    }
}
```

---

#### **generate_separator()メソッド（内部ヘルパー）**

**シグネチャ:**
```rust
fn generate_separator(&self, col_widths: &[usize]) -> String
```

**処理フロー:**
```
1. 各列幅に応じて "---" を生成
2. "|" で連結
3. 文字列として返却
```

**実装:**
```rust
fn generate_separator(&self, col_widths: &[usize]) -> String {
    let mut parts = vec!["|"];

    for &width in col_widths {
        parts.push(&"-".repeat(width + 2));
        parts.push("|");
    }

    parts.join("")
}
```

---

## **6. Converter Module (`converter.rs`)**

### **6.1. Converter構造体**

```rust
/// 変換処理のファサード
pub struct Converter {
    /// 変換設定
    config: ConversionConfig,

    /// セルフォーマッター
    formatter: CellFormatter,
}
```

---

### **6.2. 主要メソッドの実装**

#### **convert()メソッド**

**シグネチャ:**
```rust
pub fn convert<R: Read + Seek, W: Write>(
    &self,
    input: R,
    output: W,
) -> Result<(), XlsxToMdError>
```

**処理フロー:**
```
1. WorkbookParserの初期化
   └─ WorkbookParser::open(input)

2. シート選択
   └─ parser.select_sheets(&self.config.sheet_selector, self.config.include_hidden)

3. 各シートについて処理（ループ）
   ├─ シートのパース
   │  └─ parser.parse_sheet(sheet_name, &self.config)
   │
   ├─ セルのフォーマット
   │  ├─ raw_cells をイテレート
   │  ├─ self.formatter.format_cell() で各セルをフォーマット
   │  └─ formatted_cells: Vec<(CellCoord, String)> を生成
   │
   ├─ グリッドの構築
   │  └─ LogicalGrid::build(raw_cells, formatted_cells, metadata, merge_strategy)
   │
   ├─ Markdownの出力
   │  ├─ シート名をヘッダーとして出力（"# シート名"）
   │  └─ grid.render_markdown(writer)
   │
   └─ 次のシートへ

4. 出力バッファをフラッシュ
   └─ writer.flush()

5. 完了
```

**実装:**
```rust
pub fn convert<R: Read + Seek, W: Write>(
    &self,
    input: R,
    mut output: W,
) -> Result<(), XlsxToMdError> {
    use std::io::{BufWriter, Write};

    // 1. WorkbookParserの初期化
    let mut parser = WorkbookParser::open(input)?;

    // 2. シート選択
    let sheet_names = parser.select_sheets(
        &self.config.sheet_selector,
        self.config.include_hidden,
    )?;

    // バッファライター
    let mut writer = BufWriter::new(&mut output);

    // 3. 各シートの処理
    for (sheet_idx, sheet_name) in sheet_names.iter().enumerate() {
        // シート間の区切り
        if sheet_idx > 0 {
            writeln!(writer, "\n---\n")?;
        }

        // シート名をヘッダーとして出力
        writeln!(writer, "# {}\n", sheet_name)?;

        // シートのパース
        let (metadata, raw_cells) = parser.parse_sheet(sheet_name, &self.config)?;

        // セルのフォーマット
        let mut formatted_cells = Vec::new();
        for raw_cell in &raw_cells {
            let content = self.formatter.format_cell(raw_cell, &self.config)?;
            formatted_cells.push((raw_cell.coord, content));
        }

        // グリッドの構築
        let grid = LogicalGrid::build(
            raw_cells,
            formatted_cells,
            &metadata,
            self.config.merge_strategy,
        )?;

        // Markdownの出力
        grid.render_markdown(&mut writer)?;
    }

    // 4. フラッシュ
    writer.flush()?;

    Ok(())
}
```

---

#### **convert_to_string()メソッド**

**シグネチャ:**
```rust
pub fn convert_to_string<R: Read + Seek>(
    &self,
    input: R,
) -> Result<String, XlsxToMdError>
```

**処理フロー:**
```
1. メモリバッファを作成
   └─ Vec<u8>

2. convert() を呼び出し
   └─ self.convert(input, &mut buffer)

3. バッファを文字列に変換
   └─ String::from_utf8(buffer)

4. 文字列を返却
```

**実装:**
```rust
pub fn convert_to_string<R: Read + Seek>(
    &self,
    input: R,
) -> Result<String, XlsxToMdError> {
    let mut buffer = Vec::new();
    self.convert(input, &mut buffer)?;

    let result = String::from_utf8(buffer)
        .map_err(|e| XlsxToMdError::Io(
            std::io::Error::new(std::io::ErrorKind::InvalidData, e)
        ))?;

    Ok(result)
}
```

---

## **7. Number Format Parser Submodule (`format/mod.rs`)**

### **7.1. モジュール構成**

```
format/
├── mod.rs           // エントリーポイント
├── parser.rs        // FormatParserの実装
├── tokens.rs        // FormatTokenの定義
└── sections.rs      // FormatSectionの定義
```

---

### **7.2. FormatParser構造体**

**目的:** Excel Number Format Stringの構文解析。

```rust
/// Number Format Stringパーサー
pub(crate) struct FormatParser {
    /// 元のフォーマット文字列
    format_string: String,

    /// パースされたセクション
    sections: Vec<FormatSection>,
}
```

---

### **7.3. FormatSection構造体**

```rust
/// フォーマットのセクション（正数、負数、ゼロ、テキスト）
#[derive(Debug, Clone)]
pub(crate) struct FormatSection {
    /// セクションの種類
    pub kind: SectionKind,

    /// 条件（例: [>100]）
    pub condition: Option<Condition>,

    /// フォーマットトークン
    pub tokens: Vec<FormatToken>,
}
```

```rust
/// セクションの種類
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SectionKind {
    Positive,  // 正数
    Negative,  // 負数
    Zero,      // ゼロ
    Text,      // テキスト
}
```

---

### **7.4. FormatToken列挙型**

```rust
/// フォーマットトークン
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum FormatToken {
    /// 年（例: "yyyy" -> 4桁）
    Year(usize),

    /// 月（例: "mm" -> 2桁）
    Month(usize),

    /// 日（例: "dd" -> 2桁）
    Day(usize),

    /// 時（例: "hh" -> 2桁）
    Hour(usize),

    /// 分（例: "mm" -> 2桁）
    Minute(usize),

    /// 秒（例: "ss" -> 2桁）
    Second(usize),

    /// 整数部のゼロパディング（例: "0" -> 1桁）
    IntegerZero(usize),

    /// 整数部の空白パディング（例: "#"）
    IntegerHash,

    /// 小数点
    DecimalPoint,

    /// 小数部のゼロパディング（例: "0" -> 1桁）
    DecimalZero(usize),

    /// 千の位区切り
    ThousandSeparator,

    /// パーセント記号
    Percent,

    /// リテラル文字列
    Literal(String),

    /// 色指定（例: "[Red]"）
    Color(String),
}
```

---

### **7.5. Phase I実装範囲と制限事項**

#### **Phase Iでサポートする書式**

**日付・時刻書式:**
- `yyyy-mm-dd` (ISO 8601形式)
- `mm/dd/yyyy` (米国形式)
- `dd-mmm-yy` (短縮月名)
- `hh:mm:ss` (24時間制)
- `hh:mm AM/PM` (12時間制)
- 基本的な日付・時刻の組み合わせ

**数値書式:**
- `0` (整数)
- `0.00` (小数点以下2桁)
- `#,##0` (桁区切り)
- `#,##0.00` (桁区切り + 小数)
- `0.00%` (パーセント)
- `"$"#,##0.00` (通貨記号付き)

**特殊書式:**
- `@` (テキスト)
- `[Red]`, `[Blue]` などの基本色指定
- `General` (汎用書式)

**処理方針:**
```rust
// サポートされた書式の例
match format_token {
    FormatToken::Year(4) => /* yyyy */,
    FormatToken::Month(2) => /* mm */,
    FormatToken::Day(2) => /* dd */,
    FormatToken::Digit { zero_pad: true, .. } => /* 0 */,
    FormatToken::Digit { zero_pad: false, .. } => /* # */,
    FormatToken::DecimalPoint => /* . */,
    FormatToken::ThousandsSeparator => /* , */,
    FormatToken::Percent => /* % */,
    FormatToken::Color(name) => /* [Red] など */,
    FormatToken::Literal(s) => /* "$", "-" など */,
    _ => /* サポート外 */
}
```

#### **Phase Iでサポートしない書式**

**条件付き書式:**
- `[>1000]`, `[<0]`, `[=0]` などの条件
- Phase Iでは条件部分を無視し、デフォルトセクションのみを使用

**ロケール依存書式:**
- `[$-409]` (言語コード指定)
- `[$¥-411]` (日本円記号 + 日本語ロケール)
- Phase Iでは常に英語ロケールとして処理

**科学記法:**
- `0.00E+00`
- Phase Iでは `to_string()` へフォールバック

**分数:**
- `# ?/?`, `# ??/??`
- Phase Iでは `to_string()` へフォールバック

**高度な日付書式:**
- `[$-F800]dddd, mmmm dd, yyyy` (ロケール付き長い日付)
- `[DBNum1][$-804]yyyy"年"m"月"d"日"` (漢数字)
- Phase Iでは基本的な日付形式へフォールバック

**エラーハンドリング戦略:**
```rust
impl FormatParser {
    pub fn format(&self, value: &DataType) -> Result<String, XlsxToMdError> {
        match self.try_format(value) {
            Ok(formatted) => Ok(formatted),
            Err(XlsxToMdError::UnsupportedNumberFormat(_)) => {
                // Phase Iではフォールバック
                Ok(value.to_string())
            }
            Err(e) => Err(e),
        }
    }

    fn try_format(&self, value: &DataType) -> Result<String, XlsxToMdError> {
        // 書式トークンを順に処理
        for token in &self.tokens {
            match token {
                FormatToken::ConditionalOperator(_) => {
                    return Err(XlsxToMdError::UnsupportedNumberFormat(
                        "Conditional formats not supported in Phase I".to_string()
                    ));
                }
                FormatToken::ScientificNotation => {
                    return Err(XlsxToMdError::UnsupportedNumberFormat(
                        "Scientific notation not supported in Phase I".to_string()
                    ));
                }
                // ... サポートされたトークンの処理
                _ => { /* 処理 */ }
            }
        }
        Ok(formatted_string)
    }
}
```

#### **Phase IIでの拡張計画**

Phase IIでは以下をサポート予定:
- 条件付き書式の完全サポート
- ロケール依存書式（`xl/styles.xml` からの numFmtId マッピング）
- 科学記法
- 分数表示
- カスタム書式文字列の高度なパース

---

### **7.6. 主要メソッドの実装**

#### **parse()メソッド**

**シグネチャ:**
```rust
pub fn parse(format_string: &str) -> Result<Self, XlsxToMdError>
```

**処理フロー:**
```
1. セクション分割
   ├─ ';' で分割（ただし "[...]" 内は除外）
   └─ 最大4セクション（正数、負数、ゼロ、テキスト）

2. 各セクションのパース
   ├─ 条件のパース（"[>100]" など）
   ├─ トークンのパース
   └─ FormatSection を生成

3. FormatParser インスタンスを返却
```

**実装:**
```rust
pub fn parse(format_string: &str) -> Result<Self, XlsxToMdError> {
    // 1. セクション分割
    let section_strs = Self::split_sections(format_string);

    // 2. 各セクションのパース
    let mut sections = Vec::new();
    for (idx, section_str) in section_strs.iter().enumerate() {
        let kind = match idx {
            0 => SectionKind::Positive,
            1 => SectionKind::Negative,
            2 => SectionKind::Zero,
            3 => SectionKind::Text,
            _ => break,
        };

        let section = Self::parse_section(section_str, kind)?;
        sections.push(section);
    }

    Ok(Self {
        format_string: format_string.to_string(),
        sections,
    })
}
```

---

#### **format_number()メソッド**

**シグネチャ:**
```rust
pub fn format_number(&self, value: f64) -> Result<String, XlsxToMdError>
```

**処理フロー:**
```
1. 適切なセクションの選択
   ├─ value > 0 -> Positive
   ├─ value < 0 -> Negative
   └─ value == 0 -> Zero

2. 条件のチェック（存在する場合）
   └─ 条件に合致しない場合は別のセクションを選択

3. トークンに基づいてフォーマット
   ├─ FormatToken をイテレート
   ├─ 各トークンを処理して文字列を生成
   └─ 連結して返却

4. フォーマット済み文字列を返却
```

**実装:**
```rust
pub fn format_number(&self, value: f64) -> Result<String, XlsxToMdError> {
    // 1. セクションの選択
    let section = self.select_section(value);

    // 2. トークンに基づいてフォーマット
    let mut result = String::new();

    for token in &section.tokens {
        match token {
            FormatToken::IntegerZero(digits) => {
                let int_part = value.abs().floor() as u64;
                result.push_str(&format!("{:0width$}", int_part, width = digits));
            }

            FormatToken::DecimalPoint => {
                result.push('.');
            }

            FormatToken::DecimalZero(digits) => {
                let frac_part = (value.abs().fract() * 10f64.powi(*digits as i32)) as u64;
                result.push_str(&format!("{:0width$}", frac_part, width = digits));
            }

            FormatToken::ThousandSeparator => {
                // 千の位区切りの挿入（既存の整数部を修正）
                // TODO: 実装
            }

            FormatToken::Percent => {
                result.push('%');
            }

            FormatToken::Literal(s) => {
                result.push_str(s);
            }

            FormatToken::Color(_) => {
                // 色指定は無視
            }

            _ => {
                // その他のトークンは未実装
            }
        }
    }

    Ok(result)
}
```

---

## **8. エラーハンドリングの詳細**

### **8.1. エラー変換の自動化**

**thiserrorの`#[from]`属性:**

```rust
#[derive(Error, Debug)]
pub enum XlsxToMdError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),  // 自動変換

    #[error("Failed to parse Excel file: {0}")]
    Parse(#[from] calamine::Error),  // 自動変換

    // 明示的な変換が必要なエラー
    #[error("Configuration error: {0}")]
    Config(String),

    #[error("Unsupported feature at sheet '{sheet}', cell {cell}: {message}")]
    UnsupportedFeature {
        sheet: String,
        cell: String,
        message: String,
    },
}
```

**使用例:**
```rust
// std::io::Error は自動的に XlsxToMdError::Io に変換される
fn read_file(path: &str) -> Result<String, XlsxToMdError> {
    let content = std::fs::read_to_string(path)?;  // ? 演算子で自動変換
    Ok(content)
}
```

---

### **8.2. コンテキスト付きエラーの生成**

**UnsupportedFeatureエラーの生成パターン:**

```rust
fn handle_unsupported_feature(
    sheet_name: &str,
    coord: CellCoord,
    message: impl Into<String>,
) -> XlsxToMdError {
    XlsxToMdError::UnsupportedFeature {
        sheet: sheet_name.to_string(),
        cell: coord.to_a1_notation(),
        message: message.into(),
    }
}
```

**使用例:**
```rust
if complex_feature_detected {
    return Err(handle_unsupported_feature(
        "Sheet1",
        CellCoord::new(5, 10),
        "Pivot table is not supported"
    ));
}
```

---

### **8.3. エラー伝播のガイドライン**

**原則:**
1. 内部関数は詳細なエラーを返す
2. 公開APIは適切に変換・集約する
3. `?`演算子を積極的に使用
4. panic!は禁止（デバッグビルドのみ許可）

**実装例:**
```rust
// 内部関数: 詳細なエラーを返す
fn parse_internal(data: &[u8]) -> Result<Data, XlsxToMdError> {
    let value = parse_value(data)?;  // エラーを伝播
    Ok(value)
}

// 公開API: エラーを集約
pub fn convert<R: Read>(input: R) -> Result<(), XlsxToMdError> {
    let data = read_data(input)?;  // IO Error -> XlsxToMdError::Io
    parse_internal(&data)?;  // Parse Error -> XlsxToMdError::Parse
    Ok(())
}
```

---

## **9. パフォーマンス最適化の実装詳細**

### **9.1. メモリ効率の実現**

#### **ストリーミング処理の実装パターン**

```rust
pub fn process_large_file<R: Read, W: Write>(
    input: R,
    output: W,
) -> Result<(), XlsxToMdError> {
    use std::io::{BufReader, BufWriter};

    // バッファサイズ: 8KB（システムページサイズの倍数）
    let mut reader = BufReader::with_capacity(8192, input);
    let mut writer = BufWriter::with_capacity(8192, output);

    // 行単位で処理（全データをメモリに展開しない）
    for row in parse_rows(&mut reader)? {
        let formatted = format_row(row)?;
        writeln!(writer, "{}", formatted)?;
    }

    writer.flush()?;
    Ok(())
}
```

---

### **9.2. 文字列アロケーションの最適化**

#### **String::with_capacity()の使用**

```rust
fn format_cell_content(value: f64, precision: usize) -> String {
    // 事前にキャパシティを確保（再アロケーション回避）
    let mut result = String::with_capacity(32);

    // フォーマット処理
    result.push_str(&value.to_string());

    result
}
```

---

### **9.3. Copy-on-Writeの回避**

```rust
// 悪い例: 不要なクローン
fn process_bad(s: String) -> String {
    s.clone()  // 不要なコピー
}

// 良い例: 参照の使用
fn process_good(s: &str) -> &str {
    s  // コピーなし
}
```

---

## **10. テストガイドライン**

### **10.1. 単体テストの構造**

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_coord_to_a1_notation() {
        assert_eq!(CellCoord::new(0, 0).to_a1_notation(), "A1");
        assert_eq!(CellCoord::new(0, 25).to_a1_notation(), "Z1");
        assert_eq!(CellCoord::new(0, 26).to_a1_notation(), "AA1");
    }

    #[test]
    fn test_cell_range_contains() {
        let range = CellRange::new(
            CellCoord::new(0, 0),
            CellCoord::new(9, 9),
        );

        assert!(range.contains(CellCoord::new(5, 5)));
        assert!(!range.contains(CellCoord::new(10, 10)));
    }

    #[test]
    fn test_date_formatter() {
        let formatter = DateFormatter;
        let config = ConversionConfig {
            date_format: DateFormat::Iso8601,
            ..Default::default()
        };

        // 2025-01-01 のシリアル値
        let serial = 45658.0;
        let result = formatter.format(serial, &config).unwrap();

        assert_eq!(result, "2025-01-01");
    }
}
```

---

### **10.2. 統合テストの構造**

```rust
// tests/integration_test.rs

use xlsxzero::ConverterBuilder;
use std::fs::File;

#[test]
fn test_basic_conversion() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/simple.xlsx").unwrap();
    let mut output = Vec::new();

    converter.convert(input, &mut output).unwrap();

    let markdown = String::from_utf8(output).unwrap();

    assert!(markdown.contains("| Header1 |"));
    assert!(markdown.contains("| Data1   |"));
}
```

---

## **11. 今後の拡張ポイント**

### **11.1. プラグインアーキテクチャの基盤**

**Traitベースの戦略パターン:**

```rust
// 将来的な拡張を見越した設計
pub trait MergeStrategyTrait {
    fn apply(&self, grid: &mut LogicalGrid, regions: &[MergedRegion]) -> Result<(), XlsxToMdError>;
}

impl MergeStrategyTrait for DataDuplicationStrategy {
    fn apply(&self, grid: &mut LogicalGrid, regions: &[MergedRegion]) -> Result<(), XlsxToMdError> {
        // 実装
    }
}

// カスタム戦略の追加が容易
struct CustomMergeStrategy;
impl MergeStrategyTrait for CustomMergeStrategy {
    fn apply(&self, grid: &mut LogicalGrid, regions: &[MergedRegion]) -> Result<(), XlsxToMdError> {
        // カスタム実装
    }
}
```

---

### **11.2. WebAssembly対応の準備**

**stdへの依存最小化:**

```rust
// no_std対応の可能性を残す
#![cfg_attr(not(feature = "std"), no_std)]

#[cfg(not(feature = "std"))]
extern crate alloc;

#[cfg(not(feature = "std"))]
use alloc::string::String;
```

---

## **付録A: 内部データフロー図**

```
RawCellData (Parser層)
    │
    ├─ coord: CellCoord
    ├─ value: CellValue
    ├─ format_id: Option<u16>
    ├─ format_string: Option<String>
    └─ formula: Option<String>
    │
    ▼
CellFormatter.format_cell()
    │
    ├─ 数式モードチェック
    ├─ 日付判定
    ├─ 数値フォーマット
    └─ 文字列エスケープ
    │
    ▼
(CellCoord, String) (フォーマット済み)
    │
    ▼
LogicalGrid.build()
    │
    ├─ グリッド初期化
    ├─ セル配置
    └─ 結合処理
    │
    ▼
LogicalGrid
    │
    └─ cells: Vec<Vec<Cell>>
    │
    ▼
LogicalGrid.render_markdown()
    │
    └─ Markdown文字列
```

---

## **付録B: 重要なアルゴリズムの疑似コード**

### **B.1. A1記法変換アルゴリズム**

```
function col_index_to_letter(col: u32) -> String:
    result = ""
    loop:
        remainder = col mod 26
        result = char('A' + remainder) + result
        if col < 26:
            break
        col = col / 26 - 1
    return result
```

**例:**
- `col=0` → `'A'`
- `col=25` → `'Z'`
- `col=26` → `'AA'` (26 mod 26 = 0 → 'A', 26/26-1 = 0 → 'A')

---

### **B.2. Excelシリアル値から日付への変換**

```
function excel_serial_to_date(serial: f64) -> NaiveDate:
    # Excelのエポック（1900年問題の調整）
    epoch = NaiveDate(1899, 12, 30)

    days = floor(serial)
    date = epoch + Duration::days(days)

    return date
```

**例:**
- `serial=45658.0` → `2025-01-01`
- `serial=1.0` → `1899-12-31`

---

**文書管理情報:**
* 作成日: 2025-11-20
* バージョン: 1.0
* 関連文書: [requirements.md](requirements.md), [architecture.md](architecture.md), [interface.md](interface.md)
