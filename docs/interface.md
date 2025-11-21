# **インターフェース設計書 (API設計書)**

## **目的**

本設計書は、xlsxzeroクレートの公開APIを厳密に定義し、利用者が直接呼び出す「契約（Contract）」を明確化する。この設計書は**プログラム開発において最も重要な文書**であり、一度公開すると変更が非常に困難になるため、慎重に設計する。

**主要な目標:**
* 利用者が直感的かつ安全にプログラムを使用できるAPIを提供する
* 型システムを活用し、コンパイル時に多くのエラーを検出できる設計にする
* Rustエコシステムの慣習に従った、エルゴノミックなAPIを提供する

**関連文書:**
* [requirements.md](requirements.md): 要件定義
* [architecture.md](architecture.md): アーキテクチャ設計

---

## **1. 公開API一覧**

### **1.1. 名前空間・モジュール構成**

```rust
xlsxzero
├── ConverterBuilder          // ビルダーパターンによる変換設定の構築
├── Converter                 // 変換処理のエントリーポイント
├── XlsxToMdError             // エラー型
├── MergeStrategy             // セル結合戦略の列挙型
├── DateFormat                // 日付形式の列挙型
├── FormulaMode               // 数式出力モードの列挙型
└── SheetSelector             // シート選択方式の列挙型
```

### **1.2. 公開構造体・列挙型一覧**

| 名前               | 種類   | 概要                                       |
| :----------------- | :----- | :----------------------------------------- |
| `ConverterBuilder` | 構造体 | 変換設定を段階的に構築するビルダー         |
| `Converter`        | 構造体 | Excel→Markdown変換処理を実行するファサード |
| `XlsxToMdError`    | 列挙型 | クレート全体のエラー型                     |
| `MergeStrategy`    | 列挙型 | セル結合の処理戦略                         |
| `DateFormat`       | 列挙型 | 日付の出力形式                             |
| `FormulaMode`      | 列挙型 | 数式の出力モード                           |
| `SheetSelector`    | 列挙型 | シート選択方式                             |

---

## **2. 各APIの詳細定義**

### **2.1. ConverterBuilder構造体**

#### **概要**
Fluent Builder APIを提供し、`Converter`インスタンスを段階的に構築する。すべての設定項目にデフォルト値が設定されており、利用者は必要な設定のみをオーバーライド可能。

#### **コンストラクタ**

##### `ConverterBuilder::new() -> Self`

**シグネチャ:**
```rust
pub fn new() -> Self
```

**概要:**
デフォルト設定を持つビルダーインスタンスを生成する。

**戻り値:**
* `ConverterBuilder`: デフォルト設定済みのビルダー

**デフォルト設定:**
* シート選択: すべてのシート
* セル結合戦略: データ重複フィル
* 日付形式: ISO 8601 (YYYY-MM-DD)
* 非表示要素: スキップ
* 数式モード: キャッシュ値を出力

**使用例:**
```rust
use xlsxzero::ConverterBuilder;

let builder = ConverterBuilder::new();
```

---

#### **設定メソッド**

##### `with_sheet_selector(self, selector: SheetSelector) -> Self`

**シグネチャ:**
```rust
pub fn with_sheet_selector(self, selector: SheetSelector) -> Self
```

**概要:**
変換対象のシートを選択する。

**引数:**
* `selector: SheetSelector`: シート選択方式
  * `SheetSelector::All`: すべてのシートを変換（デフォルト）
  * `SheetSelector::Index(usize)`: インデックス指定（0始まり）
  * `SheetSelector::Name(String)`: シート名指定
  * `SheetSelector::Indices(Vec<usize>)`: 複数のインデックス指定
  * `SheetSelector::Names(Vec<String>)`: 複数のシート名指定

**戻り値:**
* `Self`: メソッドチェーン用のビルダー

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, SheetSelector};

// 単一シートをインデックスで指定
let builder = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Index(0));

// 単一シートを名前で指定
let builder = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Name("Sheet1".to_string()));

// 複数シートを指定
let builder = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Indices(vec![0, 2, 4]));
```

---

##### `with_merge_strategy(self, strategy: MergeStrategy) -> Self`

**シグネチャ:**
```rust
pub fn with_merge_strategy(self, strategy: MergeStrategy) -> Self
```

**概要:**
セル結合の処理戦略を指定する。

**引数:**
* `strategy: MergeStrategy`: セル結合戦略
  * `MergeStrategy::DataDuplication`: 結合セル範囲内に親セルの値を複製（デフォルト）
  * `MergeStrategy::HtmlFallback`: HTMLテーブル（rowspan/colspan）として出力

**戻り値:**
* `Self`: メソッドチェーン用のビルダー

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, MergeStrategy};

let builder = ConverterBuilder::new()
    .with_merge_strategy(MergeStrategy::HtmlFallback);
```

---

##### `with_date_format(self, format: DateFormat) -> Self`

**シグネチャ:**
```rust
pub fn with_date_format(self, format: DateFormat) -> Self
```

**概要:**
日付の出力形式を指定する。

**引数:**
* `format: DateFormat`: 日付形式
  * `DateFormat::Iso8601`: ISO 8601形式（YYYY-MM-DD）（デフォルト）
  * `DateFormat::Custom(String)`: カスタム形式（chrono互換フォーマット文字列）

**戻り値:**
* `Self`: メソッドチェーン用のビルダー

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, DateFormat};

// ISO 8601形式（デフォルト）
let builder = ConverterBuilder::new()
    .with_date_format(DateFormat::Iso8601);

// カスタム形式
let builder = ConverterBuilder::new()
    .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()));
```

---

##### `with_formula_mode(self, mode: FormulaMode) -> Self`

**シグネチャ:**
```rust
pub fn with_formula_mode(self, mode: FormulaMode) -> Self
```

**概要:**
数式セルの出力モードを指定する。

**引数:**
* `mode: FormulaMode`: 数式出力モード
  * `FormulaMode::CachedValue`: キャッシュされた結果値を出力（デフォルト）
  * `FormulaMode::Formula`: 数式文字列を出力

**戻り値:**
* `Self`: メソッドチェーン用のビルダー

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, FormulaMode};

let builder = ConverterBuilder::new()
    .with_formula_mode(FormulaMode::Formula);
```

---

##### `include_hidden(self, include: bool) -> Self`

**シグネチャ:**
```rust
pub fn include_hidden(self, include: bool) -> Self
```

**概要:**
非表示要素（非表示シート、行、列）を出力に含めるかを指定する。

> **⚠️ Phase I制限事項:**
> Phase Iでは `calamine` ライブラリの制限により、非表示行・非表示列の情報を取得できません。
> そのため、`include_hidden(false)` を指定しても**非表示行・列のフィルタリングは機能しません**。
> 非表示シートのみがフィルタリング対象となります。
> Phase IIで `xl/worksheets/sheet*.xml` から `hidden="1"` 属性を直接パースすることで完全対応予定です。

**引数:**
* `include: bool`:
  * `true`: 非表示要素を含める
  * `false`: 非表示要素をスキップ（デフォルト）
    * **Phase I**: 非表示シートのみスキップ
    * **Phase II**: 非表示行・列もスキップ

**戻り値:**
* `Self`: メソッドチェーン用のビルダー

**使用例:**
```rust
use xlsxzero::ConverterBuilder;

let builder = ConverterBuilder::new()
    .include_hidden(true);
```

---

##### `with_range(self, start: (u32, u32), end: (u32, u32)) -> Self`

**シグネチャ:**
```rust
pub fn with_range(self, start: (u32, u32), end: (u32, u32)) -> Self
```

**概要:**
処理対象のセル範囲を制限する。範囲外のセルは無視される。

**引数:**
* `start: (u32, u32)`: 開始セル座標 (row, col)（0始まり）
* `end: (u32, u32)`: 終了セル座標 (row, col)（0始まり）

**制約:**
* `start.0 <= end.0` かつ `start.1 <= end.1` でなければならない
* 制約違反の場合、`build()`時に`XlsxToMdError::Config`を返す

**戻り値:**
* `Self`: メソッドチェーン用のビルダー

**使用例:**
```rust
use xlsxzero::ConverterBuilder;

// A1:C10の範囲を処理（0始まりなので、row 0-9, col 0-2）
let builder = ConverterBuilder::new()
    .with_range((0, 0), (9, 2));
```

---

#### **ビルドメソッド**

##### `build(self) -> Result<Converter, XlsxToMdError>`

**シグネチャ:**
```rust
pub fn build(self) -> Result<Converter, XlsxToMdError>
```

**概要:**
設定を検証し、`Converter`インスタンスを生成する。

**戻り値:**
* `Ok(Converter)`: 設定が有効な場合、Converterインスタンス
* `Err(XlsxToMdError::Config)`: 設定が無効な場合（例: 範囲指定の開始 > 終了）

**発生し得るエラー:**
* `XlsxToMdError::Config(String)`: 設定の検証に失敗した場合
  * 範囲指定の開始座標が終了座標より大きい
  * カスタム日付形式が不正な書式文字列

**使用例:**
```rust
use xlsxzero::ConverterBuilder;

let converter = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Index(0))
    .build()?;
```

---

### **2.2. Converter構造体**

#### **概要**
Excel→Markdown変換処理を実行するファサード。`ConverterBuilder`から生成される。

#### **変換メソッド**

##### `convert<R: Read + Seek, W: Write>(&self, input: R, output: W) -> Result<(), XlsxToMdError>`

**シグネチャ:**
```rust
pub fn convert<R: Read + Seek, W: Write>(
    &self,
    input: R,
    output: W
) -> Result<(), XlsxToMdError>
```

**概要:**
Excelファイルを読み込み、Markdown形式に変換して出力する。ストリーミング処理により、大規模ファイルでもメモリ効率的に変換可能。

**型パラメータ:**
* `R: Read + Seek`: 入力ソース（`std::io::Read`および`std::io::Seek`トレイトを実装する型）
  * **Seekが必須な理由:** XLSXファイルはZIP形式で、ZIP Central Directoryがファイル末尾に配置されています。`calamine` は ZIP アーカイブを解析する際、まずファイル末尾の Central Directory を読み取り、その後個々のエントリ（`xl/workbook.xml`, `xl/worksheets/sheet1.xml` など）の位置へシークする必要があるため、`Seek` トレイトが必須です。順次読み取り専用（`Read` のみ）では ZIP 解析が不可能です。
  * 対応する型: `std::fs::File`, `std::io::Cursor<Vec<u8>>`, `std::io::Cursor<&[u8]>`
  * 非対応の型: `std::io::stdin()`, `TcpStream`, `BufReader<TcpStream>` など（Seekが実装されていない）
* `W: Write`: 出力先（`std::io::Write`トレイトを実装する型）

**引数:**
* `input: R`: Excelファイルの読み取りソース
  * `std::fs::File`: ファイルシステムから読み込む場合
  * `std::io::Cursor<Vec<u8>>`: カーソル付きバッファから読み込む場合（メモリバッファ）
  * `std::io::Cursor<&[u8]>`: 読み取り専用バッファから読み込む場合
* `output: W`: Markdown出力先
  * `std::fs::File`: ファイルシステムへ書き込む場合
  * `Vec<u8>`: メモリバッファへ書き込む場合
  * `std::io::stdout()`: 標準出力へ書き込む場合

**戻り値:**
* `Ok(())`: 変換が成功した場合
* `Err(XlsxToMdError)`: 変換中にエラーが発生した場合

**発生し得るエラー:**
* `XlsxToMdError::Io`: I/Oエラー（ファイル読み込み/書き込み失敗）
* `XlsxToMdError::Parse`: Excelファイルの解析に失敗
* `XlsxToMdError::Config`: 指定されたシート名が存在しない
* `XlsxToMdError::UnsupportedFeature`: 未サポートの機能に遭遇

**処理フロー:**
1. 入力ストリームからExcelファイルを解析（calamine使用）
2. 指定されたシートを選択
3. セルデータをストリーミング処理で抽出
4. 書式適用（日付変換、数値フォーマット）
5. セル結合処理（指定された戦略に基づく）
6. Markdownテーブルとして出力ストリームへ書き込み
7. 出力バッファをフラッシュ

**使用例:**
```rust
use xlsxzero::ConverterBuilder;
use std::fs::File;

// ファイル→ファイルの変換
let converter = ConverterBuilder::new().build()?;
let input = File::open("input.xlsx")?;
let output = File::create("output.md")?;
converter.convert(input, output)?;

// メモリバッファからの変換
use std::io::Cursor;
let excel_data: Vec<u8> = vec![/* ... */];
let mut markdown_output = Vec::new();
converter.convert(Cursor::new(excel_data), &mut markdown_output)?;

// 標準出力への変換
let input = File::open("input.xlsx")?;
converter.convert(input, std::io::stdout())?;
```

---

##### `convert_to_string<R: Read + Seek>(&self, input: R) -> Result<String, XlsxToMdError>`

**シグネチャ:**
```rust
pub fn convert_to_string<R: Read + Seek>(
    &self,
    input: R
) -> Result<String, XlsxToMdError>
```

**概要:**
Excelファイルを読み込み、Markdown形式の文字列として返す。内部的に`convert()`を呼び出し、メモリバッファに出力する便利メソッド。

**型パラメータ:**
* `R: Read + Seek`: 入力ソース（`std::io::Read`および`std::io::Seek`トレイトを実装する型）

**引数:**
* `input: R`: Excelファイルの読み取りソース

**戻り値:**
* `Ok(String)`: 変換されたMarkdown文字列
* `Err(XlsxToMdError)`: 変換中にエラーが発生した場合

**発生し得るエラー:**
* `convert()`と同じエラー（詳細は上記参照）

**使用例:**
```rust
use xlsxzero::ConverterBuilder;
use std::fs::File;

let converter = ConverterBuilder::new().build()?;
let input = File::open("input.xlsx")?;
let markdown = converter.convert_to_string(input)?;
println!("{}", markdown);
```

---

### **2.3. XlsxToMdError列挙型**

#### **概要**
クレート全体で使用される構造化エラー型。`thiserror`クレートを使用して実装され、自動的に`std::error::Error`トレイトを実装する。

#### **定義**

```rust
use thiserror::Error;

#[derive(Error, Debug)]
pub enum XlsxToMdError {
    /// I/Oエラー（ファイル操作や入出力ストリームの問題）
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// Excelファイルの解析エラー
    #[error("Failed to parse Excel file: {0}")]
    Parse(#[from] calamine::Error),

    /// 設定エラー（不正なシート名や範囲指定）
    #[error("Configuration error: {0}")]
    Config(String),

    /// 未サポート機能エラー
    #[error("Unsupported feature at sheet '{sheet}', cell {cell}: {message}")]
    UnsupportedFeature {
        sheet: String,
        cell: String,
        message: String,
    },
}
```

#### **エラーバリアントの詳細**

##### `XlsxToMdError::Io(std::io::Error)`

**発生条件:**
* 入力ファイルが存在しない
* 出力ファイルへの書き込み権限がない
* ディスク容量不足
* ネットワークストリームの接続エラー

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, XlsxToMdError};
use std::fs::File;

let converter = ConverterBuilder::new().build()?;

match converter.convert(
    File::open("nonexistent.xlsx")?,
    std::io::stdout()
) {
    Ok(_) => println!("変換成功"),
    Err(XlsxToMdError::Io(e)) => eprintln!("I/Oエラー: {}", e),
    Err(e) => eprintln!("その他のエラー: {}", e),
}
```

---

##### `XlsxToMdError::Parse(calamine::Error)`

**発生条件:**
* ファイルが有効なExcel形式ではない
* ファイルが破損している
* ZIPアーカイブ構造が不正
* XMLパースエラー

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, XlsxToMdError};

let converter = ConverterBuilder::new().build()?;
let broken_data = b"This is not an Excel file";

match converter.convert(&broken_data[..], std::io::stdout()) {
    Ok(_) => println!("変換成功"),
    Err(XlsxToMdError::Parse(e)) => eprintln!("解析エラー: {}", e),
    Err(e) => eprintln!("その他のエラー: {}", e),
}
```

---

##### `XlsxToMdError::Config(String)`

**発生条件:**
* 指定されたシート名が存在しない
* 指定されたシートインデックスが範囲外
* セル範囲の指定が不正（開始 > 終了）
* カスタム日付形式が不正

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, SheetSelector, XlsxToMdError};
use std::fs::File;

let converter = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Name("NonExistentSheet".to_string()))
    .build()?;

match converter.convert(File::open("input.xlsx")?, std::io::stdout()) {
    Ok(_) => println!("変換成功"),
    Err(XlsxToMdError::Config(msg)) => eprintln!("設定エラー: {}", msg),
    Err(e) => eprintln!("その他のエラー: {}", e),
}
```

---

##### `XlsxToMdError::UnsupportedFeature { sheet, cell, message }`

**発生条件:**
* サポート範囲外の複雑な書式に遭遇
* 未実装の機能（例: ピボットテーブル、マクロ）
* 深くネストされた構造（セキュリティ制限）

**フィールド:**
* `sheet: String`: エラーが発生したシート名
* `cell: String`: エラーが発生したセル座標（例: "A1"）
* `message: String`: エラーの詳細説明

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, XlsxToMdError};
use std::fs::File;

let converter = ConverterBuilder::new().build()?;

match converter.convert(File::open("complex.xlsx")?, std::io::stdout()) {
    Ok(_) => println!("変換成功"),
    Err(XlsxToMdError::UnsupportedFeature { sheet, cell, message }) => {
        eprintln!("未サポート機能: シート '{}', セル {}: {}", sheet, cell, message);
    },
    Err(e) => eprintln!("その他のエラー: {}", e),
}
```

---

### **2.4. MergeStrategy列挙型**

#### **概要**
セル結合の処理戦略を定義する列挙型。

#### **定義**

```rust
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeStrategy {
    /// 結合セル範囲内のすべてのセルに親セルの値を複製
    /// 純粋なMarkdownテーブルとして出力（デフォルト）
    DataDuplication,

    /// HTMLテーブル（rowspan/colspan属性）として出力
    /// 構造的忠実性を維持
    HtmlFallback,
}
```

#### **バリアントの詳細**

##### `MergeStrategy::DataDuplication`

**動作:**
* 結合セル範囲（例: A1:C1）を検出
* 親セル（A1）の値を取得
* 結合範囲内のすべてのセルに親セルの値を複製
* 純粋なMarkdownテーブルとして出力

**利点:**
* LLMが理解しやすい
* トークン効率が高い
* RAGシステムに最適

**欠点:**
* 視覚的な結合情報が失われる

**出力例:**
```markdown
| Header1 | Header1 | Header1 |
| ------- | ------- | ------- |
| Data1   | Data2   | Data3   |
```

---

##### `MergeStrategy::HtmlFallback`

**動作:**
* 結合セルを検出した場合、テーブル全体をHTMLとして出力
* `<td rowspan="...">`および`<td colspan="...">`属性を使用

**利点:**
* 構造的忠実性を完全に維持
* 視覚的な結合情報を保持

**欠点:**
* トークン数が増加
* LLMによる解析が複雑になる可能性

**出力例:**
```html
<table>
  <tr>
    <th colspan="3">Header1</th>
  </tr>
  <tr>
    <td>Data1</td>
    <td>Data2</td>
    <td>Data3</td>
  </tr>
</table>
```

---

### **2.5. DateFormat列挙型**

#### **概要**
日付の出力形式を定義する列挙型。

#### **定義**

```rust
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DateFormat {
    /// ISO 8601形式（YYYY-MM-DD）（デフォルト）
    Iso8601,

    /// カスタム形式（chrono互換フォーマット文字列）
    Custom(String),
}
```

#### **バリアントの詳細**

##### `DateFormat::Iso8601`

**出力形式:** `YYYY-MM-DD`（例: `2025-11-20`）

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, DateFormat};

let converter = ConverterBuilder::new()
    .with_date_format(DateFormat::Iso8601)
    .build()?;
```

---

##### `DateFormat::Custom(String)`

**概要:**
chrono互換のフォーマット文字列を使用して、カスタム日付形式を指定する。

**フォーマット指定子（主要なもの）:**
* `%Y`: 4桁の年（例: 2025）
* `%y`: 2桁の年（例: 25）
* `%m`: 2桁の月（01-12）
* `%d`: 2桁の日（01-31）
* `%B`: 月の完全名（例: January）
* `%b`: 月の省略名（例: Jan）

**詳細:** [chrono::format::strftime](https://docs.rs/chrono/latest/chrono/format/strftime/index.html) を参照。

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, DateFormat};

// 日本語形式
let converter = ConverterBuilder::new()
    .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()))
    .build()?;
// 出力例: 2025年11月20日

// US形式
let converter = ConverterBuilder::new()
    .with_date_format(DateFormat::Custom("%m/%d/%Y".to_string()))
    .build()?;
// 出力例: 11/20/2025
```

---

### **2.6. FormulaMode列挙型**

#### **概要**
数式セルの出力モードを定義する列挙型。

#### **定義**

```rust
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaMode {
    /// キャッシュされた結果値を出力（デフォルト）
    CachedValue,

    /// 数式文字列を出力
    Formula,
}
```

#### **バリアントの詳細**

##### `FormulaMode::CachedValue`

**動作:**
* Excelファイルにキャッシュされた数式の結果値を出力
* 数式を評価しない（評価エンジンを持たないため）

**利点:**
* 安定した処理
* 追加の依存関係不要

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, FormulaMode};

let converter = ConverterBuilder::new()
    .with_formula_mode(FormulaMode::CachedValue)
    .build()?;
```

**出力例:**
Excelセル: `=SUM(A1:A10)` → キャッシュ値: `123.45`
Markdown出力: `123.45`

---

##### `FormulaMode::Formula`

**動作:**
* 数式文字列をそのまま出力
* 数式を評価しない

**利点:**
* 数式の構造を保持
* デバッグや監査に有用

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, FormulaMode};

let converter = ConverterBuilder::new()
    .with_formula_mode(FormulaMode::Formula)
    .build()?;
```

**出力例:**
Excelセル: `=SUM(A1:A10)` → Markdown出力: `=SUM(A1:A10)`

---

### **2.7. SheetSelector列挙型**

#### **概要**
変換対象のシートを選択する方式を定義する列挙型。

#### **定義**

```rust
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetSelector {
    /// すべてのシートを変換（デフォルト）
    All,

    /// 単一シートをインデックスで指定（0始まり）
    Index(usize),

    /// 単一シートを名前で指定
    Name(String),

    /// 複数シートをインデックスで指定
    Indices(Vec<usize>),

    /// 複数シートを名前で指定
    Names(Vec<String>),
}
```

#### **バリアントの詳細**

##### `SheetSelector::All`

**動作:**
* ワークブック内のすべてのシートを順番に変換
* 非表示シートは`include_hidden()`設定に従う

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, SheetSelector};

let converter = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::All)
    .build()?;
```

---

##### `SheetSelector::Index(usize)`

**動作:**
* 指定されたインデックスのシートのみを変換
* インデックスは0始まり

**エラー:**
* インデックスが範囲外の場合、`XlsxToMdError::Config`を返す

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, SheetSelector};

// 最初のシートのみを変換
let converter = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Index(0))
    .build()?;
```

---

##### `SheetSelector::Name(String)`

**動作:**
* 指定された名前のシートのみを変換
* 名前は大文字小文字を区別

**エラー:**
* 指定された名前のシートが存在しない場合、`XlsxToMdError::Config`を返す

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, SheetSelector};

let converter = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Name("Sales Data".to_string()))
    .build()?;
```

---

##### `SheetSelector::Indices(Vec<usize>)`

**動作:**
* 指定された複数のインデックスのシートを順番に変換
* インデックスは0始まり

**エラー:**
* いずれかのインデックスが範囲外の場合、`XlsxToMdError::Config`を返す

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, SheetSelector};

// 1番目、3番目、5番目のシートを変換
let converter = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Indices(vec![0, 2, 4]))
    .build()?;
```

---

##### `SheetSelector::Names(Vec<String>)`

**動作:**
* 指定された複数の名前のシートを順番に変換
* 名前は大文字小文字を区別

**エラー:**
* いずれかの名前のシートが存在しない場合、`XlsxToMdError::Config`を返す

**使用例:**
```rust
use xlsxzero::{ConverterBuilder, SheetSelector};

let converter = ConverterBuilder::new()
    .with_sheet_selector(SheetSelector::Names(vec![
        "Sheet1".to_string(),
        "Sheet3".to_string(),
    ]))
    .build()?;
```

---

## **3. 使用例（Code Snippet）**

### **3.1. 基本的な使用例**

```rust
use xlsxzero::ConverterBuilder;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // デフォルト設定でConverterを構築
    let converter = ConverterBuilder::new().build()?;

    // ファイルからファイルへの変換
    let input = File::open("input.xlsx")?;
    let output = File::create("output.md")?;
    converter.convert(input, output)?;

    println!("変換完了!");
    Ok(())
}
```

---

### **3.2. カスタム設定を使用した変換**

```rust
use xlsxzero::{ConverterBuilder, SheetSelector, MergeStrategy, DateFormat};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // カスタム設定でConverterを構築
    let converter = ConverterBuilder::new()
        .with_sheet_selector(SheetSelector::Index(0))  // 最初のシートのみ
        .with_merge_strategy(MergeStrategy::DataDuplication)  // データ重複フィル
        .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()))  // 日本語形式
        .include_hidden(false)  // 非表示要素をスキップ
        .build()?;

    let input = File::open("sales_data.xlsx")?;
    let output = File::create("sales_data.md")?;
    converter.convert(input, output)?;

    println!("変換完了!");
    Ok(())
}
```

---

### **3.3. メモリバッファからの変換**

```rust
use xlsxzero::ConverterBuilder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Excelデータをメモリ上で取得（例: HTTPレスポンスから）
    let excel_data: Vec<u8> = fetch_excel_from_api()?;

    // Converterを構築
    let converter = ConverterBuilder::new().build()?;

    // Markdown文字列として変換
    let markdown = converter.convert_to_string(&excel_data[..])?;

    println!("変換結果:\n{}", markdown);
    Ok(())
}

fn fetch_excel_from_api() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    // HTTPリクエストでExcelファイルを取得する実装
    Ok(vec![/* ... */])
}
```

---

### **3.4. 複数シートの変換**

```rust
use xlsxzero::{ConverterBuilder, SheetSelector};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 複数のシートを名前で指定
    let converter = ConverterBuilder::new()
        .with_sheet_selector(SheetSelector::Names(vec![
            "Summary".to_string(),
            "Details".to_string(),
            "Appendix".to_string(),
        ]))
        .build()?;

    let input = File::open("report.xlsx")?;
    let output = File::create("report.md")?;
    converter.convert(input, output)?;

    println!("複数シートの変換完了!");
    Ok(())
}
```

---

### **3.5. セル範囲を限定した変換**

```rust
use xlsxzero::ConverterBuilder;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // A1:E20の範囲のみを変換
    let converter = ConverterBuilder::new()
        .with_range((0, 0), (19, 4))  // 0始まり（row 0-19, col 0-4）
        .build()?;

    let input = File::open("large_spreadsheet.xlsx")?;
    let output = File::create("summary.md")?;
    converter.convert(input, output)?;

    println!("指定範囲の変換完了!");
    Ok(())
}
```

---

### **3.6. HTMLフォールバック戦略を使用した変換**

```rust
use xlsxzero::{ConverterBuilder, MergeStrategy};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // HTMLフォールバック戦略で結合セルを正確に表現
    let converter = ConverterBuilder::new()
        .with_merge_strategy(MergeStrategy::HtmlFallback)
        .build()?;

    let input = File::open("complex_table.xlsx")?;
    let output = File::create("complex_table.md")?;
    converter.convert(input, output)?;

    println!("HTMLフォールバックで変換完了!");
    Ok(())
}
```

---

### **3.7. エラーハンドリングの実装例**

```rust
use xlsxzero::{ConverterBuilder, XlsxToMdError, SheetSelector};
use std::fs::File;

fn main() {
    if let Err(e) = convert_excel() {
        match e {
            XlsxToMdError::Io(io_err) => {
                eprintln!("I/Oエラー: {}", io_err);
                eprintln!("ファイルが存在するか、権限があるか確認してください。");
            }
            XlsxToMdError::Parse(parse_err) => {
                eprintln!("解析エラー: {}", parse_err);
                eprintln!("ファイルが有効なExcel形式か確認してください。");
            }
            XlsxToMdError::Config(msg) => {
                eprintln!("設定エラー: {}", msg);
                eprintln!("シート名や範囲指定を確認してください。");
            }
            XlsxToMdError::UnsupportedFeature { sheet, cell, message } => {
                eprintln!("未サポート機能: シート '{}', セル {}", sheet, cell);
                eprintln!("詳細: {}", message);
            }
        }
        std::process::exit(1);
    }
}

fn convert_excel() -> Result<(), XlsxToMdError> {
    let converter = ConverterBuilder::new()
        .with_sheet_selector(SheetSelector::Name("Data".to_string()))
        .build()?;

    let input = File::open("input.xlsx")?;
    let output = File::create("output.md")?;
    converter.convert(input, output)?;

    println!("変換成功!");
    Ok(())
}
```

---

### **3.8. 数式文字列を出力する例**

```rust
use xlsxzero::{ConverterBuilder, FormulaMode};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 数式文字列をそのまま出力
    let converter = ConverterBuilder::new()
        .with_formula_mode(FormulaMode::Formula)
        .build()?;

    let input = File::open("financial_model.xlsx")?;
    let output = File::create("financial_model.md")?;
    converter.convert(input, output)?;

    println!("数式を含む変換完了!");
    Ok(())
}
```

---

### **3.9. 非表示要素を含めた変換**

```rust
use xlsxzero::ConverterBuilder;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 非表示の行・列・シートも含めて変換
    let converter = ConverterBuilder::new()
        .include_hidden(true)
        .build()?;

    let input = File::open("complete_data.xlsx")?;
    let output = File::create("complete_data.md")?;
    converter.convert(input, output)?;

    println!("非表示要素を含む変換完了!");
    Ok(())
}
```

---

### **3.10. CLIツールでの使用例**

```rust
use xlsxzero::{ConverterBuilder, SheetSelector, XlsxToMdError};
use std::fs::File;
use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() < 3 {
        eprintln!("使い方: {} <input.xlsx> <output.md> [sheet_index]", args[0]);
        std::process::exit(1);
    }

    let input_path = &args[1];
    let output_path = &args[2];
    let sheet_index = args.get(3).and_then(|s| s.parse::<usize>().ok());

    match convert_excel(input_path, output_path, sheet_index) {
        Ok(_) => println!("変換完了: {} -> {}", input_path, output_path),
        Err(e) => {
            eprintln!("エラー: {}", e);
            std::process::exit(1);
        }
    }
}

fn convert_excel(
    input_path: &str,
    output_path: &str,
    sheet_index: Option<usize>,
) -> Result<(), XlsxToMdError> {
    let mut builder = ConverterBuilder::new();

    if let Some(index) = sheet_index {
        builder = builder.with_sheet_selector(SheetSelector::Index(index));
    }

    let converter = builder.build()?;

    let input = File::open(input_path)?;
    let output = File::create(output_path)?;
    converter.convert(input, output)?;

    Ok(())
}
```

---

## **4. API設計の方針と慣習**

### **4.1. Rust慣習への準拠**

* **命名規則:**
  * 構造体: UpperCamelCase（例: `ConverterBuilder`）
  * 列挙型: UpperCamelCase（例: `MergeStrategy`）
  * メソッド: snake_case（例: `with_sheet_selector`）
  * 定数: SCREAMING_SNAKE_CASE（必要に応じて）

* **所有権とライフタイム:**
  * ビルダーメソッドは`self`を消費し、新しい`Self`を返す（ムーブセマンティクス）
  * `convert()`メソッドは`&self`を受け取り、複数回呼び出し可能

* **エラーハンドリング:**
  * すべての公開APIは`Result<T, XlsxToMdError>`を返す
  * `thiserror`による構造化エラー
  * `#[from]`属性による自動変換

### **4.2. 型安全性**

* **コンパイル時検証:**
  * ビルダーパターンにより、設定の組み合わせをコンパイル時に検証
  * ジェネリクスを活用し、任意の`Read`/`Write`トレイト実装型に対応

* **ランタイム検証:**
  * `build()`時に設定の妥当性を検証
  * 不正な設定は`XlsxToMdError::Config`として早期に検出

### **4.3. 拡張性**

* **列挙型の拡張:**
  * 列挙型に`#[non_exhaustive]`属性を付与し、将来のバリアント追加に対応
  * セマンティックバージョニングに従い、破壊的変更を最小化

* **トレイトベースの設計:**
  * 将来的にカスタム戦略を実装可能な設計を検討

---

## **5. バージョニングとAPI安定性**

### **5.1. セマンティックバージョニング**

本クレートは[Semantic Versioning 2.0.0](https://semver.org/)に従う。

* **MAJOR（破壊的変更）:**
  * 公開APIのシグネチャ変更
  * 列挙型のバリアント削除
  * デフォルト動作の変更

* **MINOR（機能追加）:**
  * 新しい公開API追加
  * 列挙型への新しいバリアント追加（`#[non_exhaustive]`の場合）
  * 新しいオプション設定の追加

* **PATCH（バグ修正）:**
  * 内部実装のバグ修正
  * パフォーマンス改善
  * ドキュメント修正

### **5.2. 安定性保証**

* **v1.0.0以降:**
  * 公開APIは厳格な後方互換性を保証
  * 破壊的変更は次のメジャーバージョンでのみ実施

* **v0.x.x（開発段階）:**
  * APIは流動的であり、マイナーバージョンで破壊的変更が発生する可能性

---

## **6. パフォーマンス特性**

### **6.1. メモリ使用量**

* **目標:** ピークメモリ使用量をファイルサイズの10%以下に抑制
* **実現手段:** ストリーミング処理とBufWriterの使用

### **6.2. 処理速度**

* **目標:**
  * 10MB以下のファイル: 1秒以内
  * バッチ処理: 1分あたり50ファイル以上

### **6.3. I/O効率**

* **自動最適化:** `convert()`メソッドは内部的にBufWriterを使用し、システムコール数を最小化
* **柔軟性:** `Read`/`Write`トレイトにより、ファイル、メモリ、ネットワークストリームに対応

---

## **7. スレッドセーフティ**

### **7.1. Converterのスレッドセーフティ**

* `Converter`は`Send` + `Sync`を実装
* 複数のスレッドから同時に`convert()`を呼び出し可能（各スレッドは独立した入出力ストリームを使用）

**使用例:**
```rust
use xlsxzero::ConverterBuilder;
use std::fs::File;
use std::sync::Arc;
use std::thread;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let converter = Arc::new(ConverterBuilder::new().build()?);

    let mut handles = vec![];

    for i in 0..4 {
        let converter = Arc::clone(&converter);
        let handle = thread::spawn(move || {
            let input = File::open(format!("input{}.xlsx", i)).unwrap();
            let output = File::create(format!("output{}.md", i)).unwrap();
            converter.convert(input, output).unwrap();
        });
        handles.push(handle);
    }

    for handle in handles {
        handle.join().unwrap();
    }

    println!("並列変換完了!");
    Ok(())
}
```

---

## **付録A: エラーメッセージ一覧**

| エラーバリアント     | メッセージフォーマット                                                 | 発生条件                                     |
| :------------------- | :--------------------------------------------------------------------- | :------------------------------------------- |
| `Io`                 | `"IO error: {原因}"`                                                   | ファイル操作失敗、権限不足、ディスク容量不足 |
| `Parse`              | `"Failed to parse Excel file: {原因}"`                                 | 不正なExcel形式、破損ファイル                |
| `Config`             | `"Configuration error: {詳細}"`                                        | 不正なシート名、範囲指定エラー               |
| `UnsupportedFeature` | `"Unsupported feature at sheet '{シート名}', cell {セル座標}: {詳細}"` | 未サポート機能、複雑な書式                   |

---

## **付録B: 推奨される使用パターン**

### **B.1. RAGシステムへの統合**

```rust
use xlsxzero::{ConverterBuilder, MergeStrategy};

fn process_excel_for_rag(
    excel_path: &str,
) -> Result<String, Box<dyn std::error::Error>> {
    // RAG最適化された設定
    let converter = ConverterBuilder::new()
        .with_merge_strategy(MergeStrategy::DataDuplication)  // LLMフレンドリー
        .include_hidden(false)  // ノイズ削減
        .build()?;

    let input = std::fs::File::open(excel_path)?;
    let markdown = converter.convert_to_string(input)?;

    Ok(markdown)
}
```

### **B.2. バッチ処理パイプライン**

```rust
use xlsxzero::ConverterBuilder;
use std::fs;
use std::path::Path;

fn batch_convert_directory(
    input_dir: &Path,
    output_dir: &Path,
) -> Result<usize, Box<dyn std::error::Error>> {
    let converter = ConverterBuilder::new().build()?;
    let mut count = 0;

    for entry in fs::read_dir(input_dir)? {
        let entry = entry?;
        let path = entry.path();

        if path.extension().and_then(|s| s.to_str()) == Some("xlsx") {
            let output_path = output_dir.join(
                path.file_stem().unwrap()
            ).with_extension("md");

            let input = fs::File::open(&path)?;
            let output = fs::File::create(&output_path)?;
            converter.convert(input, output)?;

            count += 1;
        }
    }

    Ok(count)
}
```

---

**文書管理情報:**
* 作成日: 2025-11-20
* バージョン: 1.0
* 関連文書: [requirements.md](requirements.md), [architecture.md](architecture.md)
