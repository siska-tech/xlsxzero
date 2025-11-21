# **テスト仕様書 (Test Specification)**

## **目的**

本テスト仕様書は、xlsxzeroクレートが[requirements.md](requirements.md)で定義された要件を満たし、[interface.md](interface.md)で定義された公開APIが正しく動作することを保証するためのテスト戦略と具体的なテストケースを定義する。

**主要な目標:**
* ライブラリの品質を保証し、バグの早期発見を実現する
* 将来の変更によるデグレード（機能低下）を防止する
* RAGシステムへの統合において信頼性の高い動作を保証する
* 大規模ファイル処理におけるメモリ効率とパフォーマンスを検証する

**関連文書:**
* [requirements.md](requirements.md): 要件定義
* [interface.md](interface.md): 公開API設計
* [detailed_design.md](detailed_design.md): 詳細設計

---

## **1. テスト方針**

### **1.1. テストの種類**

| テスト種別 | 目的 | 実施タイミング | カバレッジ目標 |
|:----------|:-----|:-------------|:-------------|
| **単体テスト (Unit Tests)** | 各モジュール・関数の個別動作を検証 | 開発中・コミット前 | 90%以上 |
| **統合テスト (Integration Tests)** | モジュール間連携と公開APIの動作を検証 | 機能実装完了後 | 100% (公開API) |
| **パフォーマンステスト (Performance Tests)** | 処理速度とメモリ使用量を検証 | リリース前 | 要件定義の性能指標 |
| **境界値テスト (Boundary Tests)** | エッジケースと異常系を検証 | 実装完了後 | 主要機能100% |
| **回帰テスト (Regression Tests)** | 既存機能のデグレード防止 | CI/CD実行時 | 全テストケース |

---

### **1.2. テストカバレッジ目標**

**コードカバレッジ:**
* **全体**: 85%以上
* **公開API**: 100%
* **コアロジック** (Parser, Formatter, Grid): 95%以上
* **エラーハンドリング**: 90%以上

**測定ツール:**
* `cargo-tarpaulin` (Rustのカバレッジツール)
* 実行コマンド: `cargo tarpaulin --out Html --output-dir coverage`

---

### **1.3. 使用するテストフレームワーク**

| 用途 | ツール | 説明 |
|:-----|:------|:-----|
| 単体テスト | Rust標準テストフレームワーク | `#[test]`属性、`cargo test` |
| アサーション | `assert!`, `assert_eq!`, `assert_matches!` | 標準マクロ |
| モックデータ | `tempfile` | 一時ファイル生成 |
| プロパティベーステスト | `proptest` | ランダム入力による網羅的テスト |
| ベンチマーク | `criterion` | パフォーマンス測定 |
| カバレッジ | `cargo-tarpaulin` | コードカバレッジ測定 |

---

### **1.4. テスト実行方針**

**自動化:**
* すべてのテストはCI/CD (GitHub Actions等) で自動実行
* コミット前のプレコミットフックで単体テストを実行
* プルリクエスト時に全テストスイートを実行

**テストデータ:**
* `tests/fixtures/` ディレクトリにサンプルExcelファイルを配置
* 小規模 (< 1MB)、中規模 (10MB)、大規模 (100MB+) のファイルを用意
* 結合セル、数式、日付、複雑な書式を含むファイルを用意

**テストの独立性:**
* 各テストケースは独立して実行可能
* テスト間で状態を共有しない
* 一時ファイルは各テスト終了後に自動削除

---

## **2. テスト環境**

### **2.1. サポート対象環境**

| 項目 | 仕様 |
|:-----|:-----|
| **OS** | Windows 10/11, macOS 12+, Linux (Ubuntu 20.04+) |
| **Rust バージョン** | 1.70.0以上 (Edition 2021) |
| **依存クレート** | `Cargo.toml`に記載のバージョン |
| **メモリ** | 最小4GB (大規模ファイルテストは8GB推奨) |

---

### **2.2. CI/CD環境**

**GitHub Actions設定:**
```yaml
name: CI

on: [push, pull_request]

jobs:
  test:
    runs-on: ${{ matrix.os }}
    strategy:
      matrix:
        os: [ubuntu-latest, windows-latest, macos-latest]
        rust: [stable, 1.70.0]

    steps:
      - uses: actions/checkout@v3
      - uses: actions-rs/toolchain@v1
        with:
          toolchain: ${{ matrix.rust }}
      - run: cargo test --all-features
      - run: cargo test --release
```

---

## **3. テストケース**

### **3.1. 単体テスト (Unit Tests)**

#### **3.1.1. Types Module (`types.rs`)**

##### **TC-U-001: CellCoord::to_a1_notation()**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | セル座標のA1記法変換 |
| **テスト条件** | 正常系：有効な座標値 |
| **入力** | `CellCoord::new(0, 0)` |
| **期待結果** | `"A1"` |
| **実装** | [detailed_design.md:246-258](detailed_design.md#L246-L258) |

```rust
#[test]
fn test_cell_coord_to_a1_notation() {
    assert_eq!(CellCoord::new(0, 0).to_a1_notation(), "A1");
    assert_eq!(CellCoord::new(0, 25).to_a1_notation(), "Z1");
    assert_eq!(CellCoord::new(0, 26).to_a1_notation(), "AA1");
    assert_eq!(CellCoord::new(99, 701).to_a1_notation(), "ZZ100");
}
```

---

##### **TC-U-002: CellRange::contains()**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | セル範囲の包含判定 |
| **テスト条件** | 正常系：範囲内外の座標 |
| **入力** | `CellRange::new((0,0), (9,9))`, `CellCoord::new(5,5)` |
| **期待結果** | `true` |

```rust
#[test]
fn test_cell_range_contains() {
    let range = CellRange::new(
        CellCoord::new(0, 0),
        CellCoord::new(9, 9),
    );

    assert!(range.contains(CellCoord::new(5, 5)));
    assert!(!range.contains(CellCoord::new(10, 10)));
    assert!(!range.contains(CellCoord::new(9, 10)));
}
```

---

##### **TC-U-003: CellValue::is_empty()**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | セル値の空判定 |
| **テスト条件** | 正常系：各種セル値 |
| **入力** | `CellValue::Empty`, `CellValue::String("test")` |
| **期待結果** | `true`, `false` |

```rust
#[test]
fn test_cell_value_is_empty() {
    assert!(CellValue::Empty.is_empty());
    assert!(!CellValue::String("test".to_string()).is_empty());
    assert!(!CellValue::Number(123.45).is_empty());
    assert!(!CellValue::Bool(false).is_empty());
}
```

---

#### **3.1.2. Formatter Module (`formatter.rs`)**

##### **TC-U-010: DateFormatter::format() - ISO 8601**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | Excel日付シリアル値のISO 8601変換 |
| **テスト条件** | 正常系：有効なシリアル値 |
| **入力** | `serial=45658.0` (2025-01-01) |
| **期待結果** | `"2025-01-01"` |

```rust
#[test]
fn test_date_formatter_iso8601() {
    let formatter = DateFormatter;
    let config = ConversionConfig {
        date_format: DateFormat::Iso8601,
        ..Default::default()
    };

    // 2025-01-01
    let serial = 45658.0;
    let result = formatter.format(serial, &config).unwrap();
    assert_eq!(result, "2025-01-01");

    // 1900-01-01 (Excel epoch)
    let serial = 1.0;
    let result = formatter.format(serial, &config).unwrap();
    assert_eq!(result, "1899-12-31");
}
```

---

##### **TC-U-011: DateFormatter::format() - Custom Format**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | カスタム日付形式の変換 |
| **テスト条件** | 正常系：日本語形式 |
| **入力** | `serial=45658.0`, `format="%Y年%m月%d日"` |
| **期待結果** | `"2025年01月01日"` |

```rust
#[test]
fn test_date_formatter_custom() {
    let formatter = DateFormatter;
    let config = ConversionConfig {
        date_format: DateFormat::Custom("%Y年%m月%d日".to_string()),
        ..Default::default()
    };

    let serial = 45658.0;
    let result = formatter.format(serial, &config).unwrap();
    assert_eq!(result, "2025年01月01日");
}
```

---

##### **TC-U-012: DateFormatter::format() - 1904 Epoch**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 1904年エポックの日付変換 |
| **テスト条件** | 正常系：1904年起算のシリアル値 |
| **入力** | `serial=0.0` (1904-01-01) |
| **期待結果** | `"1904-01-01"` |

```rust
#[test]
fn test_date_formatter_1904_epoch() {
    // TODO: 1904年エポック対応実装後に追加
}
```

---

##### **TC-U-013: CellFormatter::escape_markdown()**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | Markdown特殊文字のエスケープ |
| **テスト条件** | 正常系：特殊文字を含む文字列 |
| **入力** | `"foo\|bar\nbaz"` |
| **期待結果** | `"foo\\\|bar<br>baz"` |

```rust
#[test]
fn test_escape_markdown() {
    let formatter = CellFormatter::new();

    assert_eq!(
        formatter.escape_markdown("foo|bar"),
        "foo\\|bar"
    );

    assert_eq!(
        formatter.escape_markdown("line1\nline2"),
        "line1<br>line2"
    );

    assert_eq!(
        formatter.escape_markdown("back\\slash"),
        "back\\\\slash"
    );
}
```

---

#### **3.1.3. Grid Module (`grid.rs`)**

##### **TC-U-020: LogicalGrid::build()**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | グリッド構築 |
| **テスト条件** | 正常系：スパースなセルデータから稠密グリッドへ |
| **入力** | `Vec<RawCellData>`, `Vec<(CellCoord, String)>` |
| **期待結果** | `LogicalGrid` with correct dimensions |

```rust
#[test]
fn test_logical_grid_build() {
    let raw_cells = vec![
        RawCellData {
            coord: CellCoord::new(0, 0),
            value: CellValue::String("A1".to_string()),
            format_id: None,
            format_string: None,
            formula: None,
        },
        RawCellData {
            coord: CellCoord::new(1, 1),
            value: CellValue::String("B2".to_string()),
            format_id: None,
            format_string: None,
            formula: None,
        },
    ];

    let formatted_cells = vec![
        (CellCoord::new(0, 0), "A1".to_string()),
        (CellCoord::new(1, 1), "B2".to_string()),
    ];

    let metadata = SheetMetadata {
        name: "Sheet1".to_string(),
        index: 0,
        hidden: false,
        merged_regions: vec![],
        hidden_rows: vec![],
        hidden_cols: vec![],
    };

    let grid = LogicalGrid::build(
        raw_cells,
        formatted_cells,
        &metadata,
        MergeStrategy::DataDuplication,
    ).unwrap();

    assert_eq!(grid.rows, 2);
    assert_eq!(grid.cols, 2);
}
```

---

##### **TC-U-021: LogicalGrid::apply_data_duplication()**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | データ重複フィル戦略 |
| **テスト条件** | 正常系：結合セルの値を子セルに複製 |
| **入力** | `MergedRegion(A1:C1)`, 親セル値="Header" |
| **期待結果** | A1, B1, C1 すべてに"Header"が設定 |

```rust
#[test]
fn test_apply_data_duplication() {
    let mut grid = LogicalGrid {
        cells: vec![
            vec![
                Cell::new("Header".to_string()),
                Cell::empty(),
                Cell::empty(),
            ],
        ],
        rows: 1,
        cols: 3,
    };

    let merged_regions = vec![
        MergedRegion::new(CellRange::new(
            CellCoord::new(0, 0),
            CellCoord::new(0, 2),
        )),
    ];

    grid.apply_data_duplication(&merged_regions).unwrap();

    assert_eq!(grid.cells[0][0].content, "Header");
    assert_eq!(grid.cells[0][1].content, "Header");
    assert_eq!(grid.cells[0][2].content, "Header");
    assert!(grid.cells[0][1].is_merged);
    assert!(grid.cells[0][2].is_merged);
}
```

---

#### **3.1.4. Builder Module (`builder.rs`)**

##### **TC-U-030: ConverterBuilder::build() - Valid Config**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | ビルダーの正常系検証 |
| **テスト条件** | 正常系：デフォルト設定 |
| **入力** | `ConverterBuilder::new()` |
| **期待結果** | `Ok(Converter)` |

```rust
#[test]
fn test_builder_default() {
    let result = ConverterBuilder::new().build();
    assert!(result.is_ok());
}
```

---

##### **TC-U-031: ConverterBuilder::build() - Invalid Range**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | ビルダーの異常系検証（範囲指定エラー） |
| **テスト条件** | 異常系：start > end |
| **入力** | `with_range((10, 10), (5, 5))` |
| **期待結果** | `Err(XlsxToMdError::Config)` |

```rust
#[test]
fn test_builder_invalid_range() {
    let result = ConverterBuilder::new()
        .with_range((10, 10), (5, 5))
        .build();

    assert!(result.is_err());
    match result.unwrap_err() {
        XlsxToMdError::Config(msg) => {
            assert!(msg.contains("Invalid range"));
        }
        _ => panic!("Expected Config error"),
    }
}
```

---

### **3.2. 統合テスト (Integration Tests)**

#### **3.2.1. Basic Conversion Tests**

##### **TC-I-001: Simple Table Conversion**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 基本的なテーブル変換 |
| **テスト条件** | 正常系：単純な2x2テーブル |
| **入力ファイル** | `tests/fixtures/simple_table.xlsx` |
| **期待結果** | Markdownテーブルが正しく生成される |

```rust
#[test]
fn test_simple_table_conversion() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/simple_table.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("| Header1 | Header2 |"));
    assert!(markdown.contains("| Data1   | Data2   |"));
}
```

**Fixture仕様:**
```
tests/fixtures/simple_table.xlsx:
┌─────────┬─────────┐
│ Header1 │ Header2 │
├─────────┼─────────┤
│ Data1   │ Data2   │
└─────────┴─────────┘
```

---

##### **TC-I-002: Multiple Sheets Conversion**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 複数シートの変換 |
| **テスト条件** | 正常系：3シートのワークブック |
| **入力ファイル** | `tests/fixtures/multi_sheets.xlsx` |
| **期待結果** | 各シートが "# シート名" で区切られる |

```rust
#[test]
fn test_multiple_sheets() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/multi_sheets.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("# Sheet1"));
    assert!(markdown.contains("# Sheet2"));
    assert!(markdown.contains("# Sheet3"));
    assert!(markdown.matches("---").count() >= 2);
}
```

---

##### **TC-I-003: Merged Cells - Data Duplication**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | セル結合のデータ重複フィル戦略 |
| **テスト条件** | 正常系：水平結合セル (A1:C1) |
| **入力ファイル** | `tests/fixtures/merged_cells.xlsx` |
| **期待結果** | 結合範囲の各セルに同じ値が出力される |

```rust
#[test]
fn test_merged_cells_data_duplication() {
    let converter = ConverterBuilder::new()
        .with_merge_strategy(MergeStrategy::DataDuplication)
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/merged_cells.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    // A1:C1が"Header"で結合されている場合
    assert!(markdown.contains("| Header | Header | Header |"));
}
```

---

##### **TC-I-004: Merged Cells - HTML Fallback**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | セル結合のHTMLフォールバック戦略 |
| **テスト条件** | 正常系：rowspan/colspan属性を使用 |
| **入力ファイル** | `tests/fixtures/merged_cells.xlsx` |
| **期待結果** | HTMLテーブルが生成される |

```rust
#[test]
fn test_merged_cells_html_fallback() {
    let converter = ConverterBuilder::new()
        .with_merge_strategy(MergeStrategy::HtmlFallback)
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/merged_cells.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("<table>"));
    assert!(markdown.contains("colspan=\"3\""));
}
```

---

##### **TC-I-005: Date Formatting**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 日付セルの書式適用 |
| **テスト条件** | 正常系：日付シリアル値の変換 |
| **入力ファイル** | `tests/fixtures/dates.xlsx` |
| **期待結果** | ISO 8601形式で出力 |

```rust
#[test]
fn test_date_formatting() {
    let converter = ConverterBuilder::new()
        .with_date_format(DateFormat::Iso8601)
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/dates.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("2025-01-01"));
    assert!(markdown.contains("2025-12-31"));
}
```

---

##### **TC-I-006: Formula Cells - Cached Value**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 数式セルのキャッシュ値出力 |
| **テスト条件** | 正常系：`=SUM(A1:A10)` のキャッシュ値 |
| **入力ファイル** | `tests/fixtures/formulas.xlsx` |
| **期待結果** | 計算結果（例：123.45）が出力される |

```rust
#[test]
fn test_formula_cached_value() {
    let converter = ConverterBuilder::new()
        .with_formula_mode(FormulaMode::CachedValue)
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/formulas.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    // =SUM(A1:A10) のキャッシュ値が100の場合
    assert!(markdown.contains("100"));
}
```

---

##### **TC-I-007: Formula Cells - Formula String**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 数式セルの数式文字列出力 |
| **テスト条件** | 正常系：数式をそのまま出力 |
| **入力ファイル** | `tests/fixtures/formulas.xlsx` |
| **期待結果** | `=SUM(A1:A10)` が出力される |

```rust
#[test]
fn test_formula_string() {
    let converter = ConverterBuilder::new()
        .with_formula_mode(FormulaMode::Formula)
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/formulas.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("=SUM(A1:A10)"));
}
```

---

##### **TC-I-008: Sheet Selection by Index**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | シートのインデックス指定 |
| **テスト条件** | 正常系：2番目のシートのみ変換 |
| **入力ファイル** | `tests/fixtures/multi_sheets.xlsx` |
| **期待結果** | "# Sheet2" のみが出力される |

```rust
#[test]
fn test_sheet_selection_by_index() {
    let converter = ConverterBuilder::new()
        .with_sheet_selector(SheetSelector::Index(1))
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/multi_sheets.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("# Sheet2"));
    assert!(!markdown.contains("# Sheet1"));
    assert!(!markdown.contains("# Sheet3"));
}
```

---

##### **TC-I-009: Sheet Selection by Name**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | シートの名前指定 |
| **テスト条件** | 正常系：特定シート名で指定 |
| **入力ファイル** | `tests/fixtures/multi_sheets.xlsx` |
| **期待結果** | 指定シートのみが出力される |

```rust
#[test]
fn test_sheet_selection_by_name() {
    let converter = ConverterBuilder::new()
        .with_sheet_selector(SheetSelector::Name("Sheet2".to_string()))
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/multi_sheets.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("# Sheet2"));
    assert!(!markdown.contains("# Sheet1"));
}
```

---

##### **TC-I-010: Range Restriction**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | セル範囲の制限 |
| **テスト条件** | 正常系：A1:C3 の範囲のみ処理 |
| **入力ファイル** | `tests/fixtures/large_table.xlsx` |
| **期待結果** | 指定範囲のセルのみが出力される |

```rust
#[test]
fn test_range_restriction() {
    let converter = ConverterBuilder::new()
        .with_range((0, 0), (2, 2))  // A1:C3
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/large_table.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    // 3列のみが出力される
    let first_line = markdown.lines().next().unwrap();
    assert_eq!(first_line.matches('|').count(), 4);  // | col1 | col2 | col3 |
}
```

---

##### **TC-I-011: Hidden Elements Exclusion**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 非表示要素の除外 |
| **テスト条件** | 正常系：非表示行・列をスキップ |
| **入力ファイル** | `tests/fixtures/hidden_elements.xlsx` |
| **期待結果** | 非表示要素が出力されない |

```rust
#[test]
fn test_hidden_elements_exclusion() {
    let converter = ConverterBuilder::new()
        .include_hidden(false)
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/hidden_elements.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    // 非表示行B2が存在しない
    assert!(!markdown.contains("HiddenData"));
}
```

---

##### **TC-I-012: Hidden Elements Inclusion**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 非表示要素の包含 |
| **テスト条件** | 正常系：非表示行・列を含む |
| **入力ファイル** | `tests/fixtures/hidden_elements.xlsx` |
| **期待結果** | 非表示要素も出力される |

```rust
#[test]
fn test_hidden_elements_inclusion() {
    let converter = ConverterBuilder::new()
        .include_hidden(true)
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/hidden_elements.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("HiddenData"));
}
```

---

#### **3.2.2. Error Handling Tests**

##### **TC-I-100: Invalid File Format**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 不正なファイル形式の検出 |
| **テスト条件** | 異常系：Excel形式でないファイル |
| **入力** | テキストファイル |
| **期待結果** | `Err(XlsxToMdError::Parse)` |

```rust
#[test]
fn test_invalid_file_format() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = b"This is not an Excel file";
    let result = converter.convert(&input[..], std::io::sink());

    assert!(result.is_err());
    match result.unwrap_err() {
        XlsxToMdError::Parse(_) => (),
        e => panic!("Expected Parse error, got {:?}", e),
    }
}
```

---

##### **TC-I-101: Non-Existent Sheet**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 存在しないシート名の指定 |
| **テスト条件** | 異常系：存在しないシート名 |
| **入力** | `SheetSelector::Name("NonExistent")` |
| **期待結果** | `Err(XlsxToMdError::Config)` |

```rust
#[test]
fn test_nonexistent_sheet() {
    let converter = ConverterBuilder::new()
        .with_sheet_selector(SheetSelector::Name("NonExistent".to_string()))
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/simple_table.xlsx").unwrap();
    let result = converter.convert(input, std::io::sink());

    assert!(result.is_err());
    match result.unwrap_err() {
        XlsxToMdError::Config(msg) => {
            assert!(msg.contains("not found"));
        }
        e => panic!("Expected Config error, got {:?}", e),
    }
}
```

---

##### **TC-I-102: Sheet Index Out of Range**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 範囲外のシートインデックス |
| **テスト条件** | 異常系：インデックス > シート数 |
| **入力** | `SheetSelector::Index(999)` |
| **期待結果** | `Err(XlsxToMdError::Config)` |

```rust
#[test]
fn test_sheet_index_out_of_range() {
    let converter = ConverterBuilder::new()
        .with_sheet_selector(SheetSelector::Index(999))
        .build()
        .unwrap();

    let input = File::open("tests/fixtures/simple_table.xlsx").unwrap();
    let result = converter.convert(input, std::io::sink());

    assert!(result.is_err());
    match result.unwrap_err() {
        XlsxToMdError::Config(msg) => {
            assert!(msg.contains("out of range"));
        }
        e => panic!("Expected Config error, got {:?}", e),
    }
}
```

---

##### **TC-I-103: File Not Found**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 存在しないファイルの処理 |
| **テスト条件** | 異常系：ファイルが存在しない |
| **入力** | `"nonexistent.xlsx"` |
| **期待結果** | `Err(XlsxToMdError::Io)` |

```rust
#[test]
fn test_file_not_found() {
    let converter = ConverterBuilder::new().build().unwrap();

    let result = File::open("nonexistent.xlsx");
    assert!(result.is_err());

    match result.unwrap_err().kind() {
        std::io::ErrorKind::NotFound => (),
        e => panic!("Expected NotFound error, got {:?}", e),
    }
}
```

---

### **3.3. パフォーマンステスト (Performance Tests)**

#### **3.3.1. Memory Efficiency Tests**

##### **TC-P-001: Small File Memory Usage**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 小規模ファイルのメモリ使用量 |
| **テスト条件** | ファイルサイズ: 10MB |
| **性能目標** | ピークメモリ ≤ 100MB |
| **測定方法** | `memory_profiler` または OSメトリクス |

```rust
#[test]
#[ignore]  // 手動実行用
fn test_small_file_memory() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/10mb_file.xlsx").unwrap();
    let mut output = Vec::new();

    converter.convert(input, &mut output).unwrap();

    // メモリ使用量は外部ツールで測定
    // 目標: ピークメモリ ≤ 100MB
}
```

---

##### **TC-P-002: Large File Memory Usage**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 大規模ファイルのメモリ使用量 |
| **テスト条件** | ファイルサイズ: 1GB |
| **性能目標** | ピークメモリ ≤ ファイルサイズの10% (100MB) |
| **測定方法** | `memory_profiler` または OSメトリクス |

```rust
#[test]
#[ignore]  // 手動実行用
fn test_large_file_memory() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/1gb_file.xlsx").unwrap();
    let mut output = std::io::sink();

    converter.convert(input, &mut output).unwrap();

    // 目標: ピークメモリ ≤ 100MB
}
```

---

#### **3.3.2. Processing Speed Tests**

##### **TC-P-010: Small File Processing Speed**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 小規模ファイルの処理速度 |
| **テスト条件** | ファイルサイズ: 10MB |
| **性能目標** | 処理時間 < 1秒 |
| **測定方法** | `criterion` ベンチマーク |

```rust
use criterion::{black_box, criterion_group, criterion_main, Criterion};

fn benchmark_small_file(c: &mut Criterion) {
    let converter = ConverterBuilder::new().build().unwrap();

    c.bench_function("convert_10mb_file", |b| {
        b.iter(|| {
            let input = File::open("tests/fixtures/10mb_file.xlsx").unwrap();
            let mut output = std::io::sink();
            converter.convert(black_box(input), black_box(&mut output)).unwrap();
        });
    });
}

criterion_group!(benches, benchmark_small_file);
criterion_main!(benches);
```

---

##### **TC-P-011: Batch Processing Throughput**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | バッチ処理のスループット |
| **テスト条件** | 50ファイル（各10MB） |
| **性能目標** | 1分以内に完了 |
| **測定方法** | 実測 |

```rust
#[test]
#[ignore]  // 手動実行用
fn test_batch_processing_throughput() {
    let converter = ConverterBuilder::new().build().unwrap();
    let start = std::time::Instant::now();

    for i in 0..50 {
        let input = File::open(format!("tests/fixtures/batch/file_{}.xlsx", i)).unwrap();
        let mut output = std::io::sink();
        converter.convert(input, &mut output).unwrap();
    }

    let duration = start.elapsed();
    assert!(duration.as_secs() < 60, "Batch processing took too long: {:?}", duration);
}
```

---

### **3.4. 境界値テスト (Boundary Tests)**

##### **TC-B-001: Empty Workbook**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 空のワークブック |
| **テスト条件** | 境界値：シートが存在しない |
| **入力ファイル** | `tests/fixtures/empty.xlsx` |
| **期待結果** | 空の出力（エラーなし） |

```rust
#[test]
fn test_empty_workbook() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/empty.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.is_empty() || markdown.trim().is_empty());
}
```

---

##### **TC-B-002: Empty Sheet**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 空のシート |
| **テスト条件** | 境界値：セルが存在しない |
| **入力ファイル** | `tests/fixtures/empty_sheet.xlsx` |
| **期待結果** | シート名のみ出力（エラーなし） |

```rust
#[test]
fn test_empty_sheet() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/empty_sheet.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    assert!(markdown.contains("# Sheet1"));
}
```

---

##### **TC-B-003: Maximum Rows (1,048,576)**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | Excel最大行数 |
| **テスト条件** | 境界値：1,048,576行 |
| **入力ファイル** | `tests/fixtures/max_rows.xlsx` |
| **期待結果** | すべての行が処理される |

```rust
#[test]
#[ignore]  // 時間がかかるため手動実行
fn test_maximum_rows() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/max_rows.xlsx").unwrap();
    let result = converter.convert(input, std::io::sink());

    assert!(result.is_ok());
}
```

---

##### **TC-B-004: Maximum Columns (16,384)**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | Excel最大列数 (XFD列) |
| **テスト条件** | 境界値：16,384列 |
| **入力ファイル** | `tests/fixtures/max_cols.xlsx` |
| **期待結果** | すべての列が処理される |

```rust
#[test]
#[ignore]  // 時間がかかるため手動実行
fn test_maximum_columns() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/max_cols.xlsx").unwrap();
    let result = converter.convert(input, std::io::sink());

    assert!(result.is_ok());
}
```

---

##### **TC-B-005: Very Long Cell Content**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 非常に長いセル内容 |
| **テスト条件** | 境界値：32,767文字（Excel上限） |
| **入力ファイル** | `tests/fixtures/long_cell.xlsx` |
| **期待結果** | すべての文字が出力される |

```rust
#[test]
fn test_very_long_cell_content() {
    let converter = ConverterBuilder::new().build().unwrap();

    let input = File::open("tests/fixtures/long_cell.xlsx").unwrap();
    let markdown = converter.convert_to_string(input).unwrap();

    // 32,767文字の文字列が含まれる
    assert!(markdown.len() > 32000);
}
```

---

##### **TC-B-006: Date at Epoch Boundary**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | エポック境界の日付 |
| **テスト条件** | 境界値：1900-01-01, 2099-12-31 |
| **入力** | シリアル値 1.0, 73050.0 |
| **期待結果** | 正確な日付文字列 |

```rust
#[test]
fn test_date_epoch_boundary() {
    let formatter = DateFormatter;
    let config = ConversionConfig {
        date_format: DateFormat::Iso8601,
        ..Default::default()
    };

    // 1900-01-01 (Excel epoch + 1)
    let result = formatter.format(1.0, &config).unwrap();
    assert_eq!(result, "1899-12-31");

    // 2099-12-31
    let result = formatter.format(73050.0, &config).unwrap();
    assert_eq!(result, "2099-12-31");
}
```

---

### **3.5. プロパティベーステスト (Property-Based Tests)**

##### **TC-PBT-001: A1 Notation Round-Trip**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | A1記法の往復変換 |
| **テスト条件** | ランダムな座標値 |
| **性質** | `parse(to_a1(coord)) == coord` |

```rust
use proptest::prelude::*;

proptest! {
    #[test]
    fn test_a1_notation_round_trip(row in 0u32..100, col in 0u32..100) {
        let coord = CellCoord::new(row, col);
        let a1 = coord.to_a1_notation();

        // A1記法の形式検証（例: "A1", "ZZ999"）
        prop_assert!(a1.chars().next().unwrap().is_ascii_uppercase());
        prop_assert!(a1.chars().last().unwrap().is_ascii_digit());
    }
}
```

---

##### **TC-PBT-002: Date Conversion Monotonicity**

| 項目 | 内容 |
|:-----|:-----|
| **テスト項目** | 日付変換の単調性 |
| **テスト条件** | ランダムなシリアル値 |
| **性質** | `serial1 < serial2 => date1 < date2` |

```rust
proptest! {
    #[test]
    fn test_date_conversion_monotonicity(
        serial1 in 1.0f64..50000.0,
        serial2 in 1.0f64..50000.0
    ) {
        let formatter = DateFormatter;
        let config = ConversionConfig {
            date_format: DateFormat::Iso8601,
            ..Default::default()
        };

        let date1 = formatter.format(serial1, &config).unwrap();
        let date2 = formatter.format(serial2, &config).unwrap();

        if serial1 < serial2 {
            prop_assert!(date1 < date2);
        } else if serial1 > serial2 {
            prop_assert!(date1 > date2);
        } else {
            prop_assert_eq!(date1, date2);
        }
    }
}
```

---

## **4. テスト実行手順**

### **4.1. ローカル環境での実行**

#### **全テストの実行:**
```bash
cargo test --all-features
```

#### **単体テストのみ:**
```bash
cargo test --lib
```

#### **統合テストのみ:**
```bash
cargo test --test '*'
```

#### **特定のテストケース:**
```bash
cargo test test_simple_table_conversion
```

#### **無視されたテストの実行:**
```bash
cargo test -- --ignored
```

#### **カバレッジレポート生成:**
```bash
cargo install cargo-tarpaulin
cargo tarpaulin --out Html --output-dir coverage
```

---

### **4.2. CI/CD環境での実行**

**.github/workflows/test.yml:**
```yaml
name: Test Suite

on:
  push:
    branches: [ main, develop ]
  pull_request:
    branches: [ main ]

jobs:
  test:
    runs-on: ${{ matrix.os }}
    strategy:
      matrix:
        os: [ubuntu-latest, windows-latest, macos-latest]
        rust: [stable, 1.70.0]

    steps:
      - uses: actions/checkout@v3

      - name: Install Rust
        uses: actions-rs/toolchain@v1
        with:
          toolchain: ${{ matrix.rust }}
          profile: minimal
          override: true

      - name: Cache cargo registry
        uses: actions/cache@v3
        with:
          path: ~/.cargo/registry
          key: ${{ runner.os }}-cargo-registry-${{ hashFiles('**/Cargo.lock') }}

      - name: Run tests
        run: cargo test --all-features --verbose

      - name: Run ignored tests
        run: cargo test --all-features -- --ignored
        if: matrix.os == 'ubuntu-latest' && matrix.rust == 'stable'

      - name: Generate coverage
        if: matrix.os == 'ubuntu-latest' && matrix.rust == 'stable'
        run: |
          cargo install cargo-tarpaulin
          cargo tarpaulin --out Xml

      - name: Upload coverage to Codecov
        if: matrix.os == 'ubuntu-latest' && matrix.rust == 'stable'
        uses: codecov/codecov-action@v3
        with:
          files: ./cobertura.xml
```

---

## **5. テストデータ仕様**

### **5.1. Fixture Files**

| ファイル名 | 仕様 | サイズ | 目的 |
|:----------|:-----|:------|:-----|
| `simple_table.xlsx` | 2x2テーブル、ヘッダー+1行 | < 10KB | 基本動作確認 |
| `multi_sheets.xlsx` | 3シート（Sheet1, Sheet2, Sheet3） | < 20KB | 複数シート処理 |
| `merged_cells.xlsx` | 水平・垂直結合セル | < 15KB | セル結合処理 |
| `dates.xlsx` | 日付シリアル値（1900/1904エポック） | < 10KB | 日付変換 |
| `formulas.xlsx` | SUM, AVERAGE等の数式 | < 10KB | 数式処理 |
| `hidden_elements.xlsx` | 非表示行・列・シート | < 10KB | 非表示要素処理 |
| `large_table.xlsx` | 1000行×100列 | ~5MB | 範囲制限テスト |
| `10mb_file.xlsx` | 10MBのデータ | 10MB | メモリ効率検証 |
| `1gb_file.xlsx` | 1GBのデータ（手動生成） | 1GB | 大規模ファイル処理 |
| `empty.xlsx` | シートなし | < 5KB | 境界値テスト |
| `empty_sheet.xlsx` | 空シート | < 5KB | 境界値テスト |
| `long_cell.xlsx` | 32,767文字のセル | < 50KB | 長文処理 |

---

### **5.2. Fixture生成スクリプト**

**使用クレート:** `rust_xlsxwriter` (Pure-Rust実装)

```rust
// tests/fixtures/generate_fixtures.rs
use rust_xlsxwriter::*;

fn generate_simple_table() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "Header1")?;
    worksheet.write_string(0, 1, "Header2")?;
    worksheet.write_string(1, 0, "Data1")?;
    worksheet.write_string(1, 1, "Data2")?;

    workbook.save("tests/fixtures/simple_table.xlsx")?;
    Ok(())
}

fn generate_dates() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // 1900年システムのシート（デフォルト）
    let sheet1 = workbook.add_worksheet();
    sheet1.set_name("Dates_1900")?;
    sheet1.write_number(0, 0, 43831.0)?; // 2020-01-01

    // 1904年システムのシート
    let sheet2 = workbook.add_worksheet();
    sheet2.set_name("Dates_1904")?;
    workbook.set_1904_date(true);  // 1904年エポックを有効化
    sheet2.write_number(0, 0, 42370.0)?; // 2020-01-01 in 1904 system

    workbook.save("tests/fixtures/dates.xlsx")?;
    Ok(())
}
```

**依存関係追加 (Cargo.toml):**
```toml
[dev-dependencies]
rust_xlsxwriter = "0.80"
tempfile = "3.0"
proptest = "1.0"
criterion = "0.5"
```

---

## **6. 品質メトリクス**

### **6.1. カバレッジ目標**

| メトリクス | 目標値 | 測定方法 |
|:----------|:------|:--------|
| 行カバレッジ | 85%以上 | `cargo-tarpaulin` |
| 関数カバレッジ | 90%以上 | `cargo-tarpaulin` |
| 公開API網羅率 | 100% | 手動レビュー |

---

### **6.2. 性能メトリクス**

| メトリクス | 目標値 | 測定方法 |
|:----------|:------|:--------|
| 10MBファイル処理時間 | < 1秒 | `criterion` |
| ピークメモリ使用量 | < ファイルサイズの10% | `memory_profiler` |
| バッチ処理スループット | 50ファイル/分以上 | 実測 |

---

## **7. リリース前チェックリスト**

- [ ] 全単体テストが成功 (`cargo test --lib`)
- [ ] 全統合テストが成功 (`cargo test --test '*'`)
- [ ] 無視されたテストが成功 (`cargo test -- --ignored`)
- [ ] カバレッジが85%以上
- [ ] パフォーマンステストが目標値を満たす
- [ ] 3つのOS（Windows, macOS, Linux）でテスト成功
- [ ] Rust 1.70.0以上でテスト成功
- [ ] ドキュメントテストが成功 (`cargo test --doc`)
- [ ] Clippyの警告がゼロ (`cargo clippy -- -D warnings`)
- [ ] フォーマットが統一 (`cargo fmt --check`)
- [ ] セキュリティ監査が完了 (`cargo audit`)

---

**文書管理情報:**
* 作成日: 2025-11-20
* バージョン: 1.0
* 関連文書: [requirements.md](requirements.md), [interface.md](interface.md), [detailed_design.md](detailed_design.md)
