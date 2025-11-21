# **純粋なRustによるExcel解析およびMarkdown変換クレートの要件定義書**

## **目的**

本要件定義書は、純粋なRust言語を用いてExcelファイル（XLSXを主対象とする）を解析し、構造化されたMarkdown形式に変換するクレートの開発に必要な要件を定義するものである。

* プログラム開発の**背景と目的（Why）**を明確にする。
* プログラムが提供すべき**機能（What）**を網羅的に定義する。
* 満たすべき**品質、性能、制約（非機能要件）**を定義する。

**本プロジェクトの主要な目的は、RAG（Retrieval-Augmented Generation）システムへの入力データを提供するために、Excelファイルを構造化されたMarkdown形式に変換することである。** RAGシステムは、大規模言語モデル（LLM）に外部知識を提供するための重要なアーキテクチャであり、Excelファイルに含まれる構造化データを効率的にベクトルデータベースに取り込むためには、LLMが理解しやすい形式（Markdown）への変換が不可欠である。

---

## **1. 開発背景と目的（Why）**

### **1.1. 解決しようとしている課題**

RAGシステムや大規模言語モデル（LLM）を活用したドキュメント解析パイプラインにおいて、Excelファイルを効率的に処理し、構造化されたテキストデータとして利用可能にする必要がある。既存のソリューションには以下の課題がある。

* **依存性の複雑性:** C/C++ライブラリに依存するツールは、ビルド環境の複雑化やクロスコンパイルの困難さを招く。
* **大規模ファイル処理の非効率性:** DOM方式の解析はメモリ消費が大きく、ギガバイト単位のExcelファイルを処理できない。
* **データ精度の問題:** Excelの日付シリアル値や複雑な書式情報を正確にMarkdownに変換できない既存ツールの存在。
* **セル結合の処理:** Markdownではネイティブにセル結合を表現できないため、データ整合性を保ちながら変換する高度なロジックが必要。

### **1.2. 想定される利用者（ターゲットユーザー）**

* **RAGシステム開発者:** Excelファイルをベクトルデータベースに取り込む前処理として使用。
* **データパイプラインエンジニア:** 大規模なExcelファイルをバッチ処理し、LLMが理解しやすい形式に変換する必要がある開発者。
* **ドキュメント変換ツール開発者:** ExcelからMarkdownへの変換機能を組み込む必要があるアプリケーション開発者。

### **1.3. このプログラムが提供する中核的な価値**

本プログラムは、RAGシステムへの入力データを提供することを主目的として設計されており、以下の価値を提供する。

1. **Pure-Rustによる堅牢性と移植性:** 
   * 外部のC/C++ライブラリへの依存を一切排除し、単一バイナリとして配布可能。
   * クロスコンパイルが容易で、WebAssembly環境への展開も可能。
   * RAGパイプラインを様々な環境（サーバーレス、エッジコンピューティング、ブラウザ）で実行可能。

2. **大規模ファイル処理能力:** 
   * ストリーミング解析アーキテクチャにより、数ギガバイト単位のExcelファイルでも、ピークメモリ使用量をファイルサイズの10%以下に抑制。
   * RAGパイプラインでのバッチ処理において、リソース効率的な処理を実現。

3. **RAG最適化された出力:** 
   * セル結合を「データ重複フィル」戦略で処理することで、LLMが理解しやすく、トークン効率の高いMarkdownテーブルを生成。
   * HTMLフォールバックも選択可能で、構造的な忠実性が必要な場合にも対応。
   * Markdown形式により、ベクトルデータベースへの埋め込みやチャンキング処理が容易。

4. **データ精度の保証:** 
   * ExcelのNumber Format Stringを正確に解析し、日付シリアル値を適切な文字列形式に変換することで、RAGシステムが正確な情報を取得可能。
   * データの意味を保持した変換により、LLMが正確な回答を生成できる基盤を提供。

---

## **2. 機能要件（What）**

本セクションでは、RAGシステムへの入力データを提供するために必要な機能を網羅的に定義する。

### **2.1. 提供する機能の一覧**

#### **2.1.1. Excelファイル解析機能**

* **ワークブック構造の解析:** シート一覧の取得、シート選択、非表示シートの検出。
* **ワークシートデータの抽出:** 行、列、セルの階層的走査とデータ抽出。
* **セル値の型判定と抽出:** f64（数値/日付）、String、Bool、Error型の識別と抽出。
* **書式情報の抽出:** Number Format String、セルの書式設定の取得。
* **セル結合範囲の検出:** 結合セルの開始位置と結合範囲（例: A1:C3）の正確な特定。
* **数式情報の抽出:** 数式文字列とキャッシュされた結果値の取得。
* **ハイパーリンクの抽出:** URLと表示テキスト情報の取得。

**入力:** Excelファイル（XLSX形式を主対象とする）への読み取りアクセス（std::io::Readトレイト）。

**期待される出力:** 構造化された中間データモデル（ワークブック構造、セルデータ、書式情報、結合範囲情報）。

#### **2.1.2. データ型変換機能**

* **日付・時刻の変換:** Excelシリアル値（f64）をNumber Format Stringに基づいてISO 8601形式（デフォルト）またはユーザー指定形式に変換。1900年起算と1904年起算の両エポックに対応。
* **数値の書式適用:** Number Format Stringを解析し、生の数値を人間が読める表示文字列に変換（例: 1234.567 → "1,234.57"）。
* **書式適用済み表示文字列の生成:** 生の値（Raw Value）と書式情報から、最終的な表示文字列を生成。Markdown変換時は常にこの表示文字列を参照。

**入力:** 生のセル値（f64、String、Bool）とNumber Format String、書式設定情報。

**期待される出力:** 書式適用済みの表示文字列。

#### **2.1.3. Markdown変換機能**

* **Markdownテーブルの生成:** ワークシートデータをGitHub Flavored Markdown (GFM) 準拠のテーブル形式に変換。
* **セル結合の処理:** 
  * **デフォルト戦略（データ重複フィル）:** 結合セル範囲内のすべての論理的な子セルに親セルの値を複製。純粋なMarkdownテーブル構造を維持。
  * **代替戦略（HTMLフォールバック）:** 結合セルを含むテーブルをHTMLテーブル（rowspan/colspan属性使用）として出力。
* **リッチテキストの変換:** 太字（**text**）、斜体（*text*）など、Markdownで表現可能な最小限の書式のみをサポート。
* **ハイパーリンクの変換:** 標準のMarkdownリンク構文 [表示テキスト](URL) へのマッピング。
* **数式の出力:** キャッシュされた結果値をデフォルトとして出力。オプションで数式文字列を出力可能。

**入力:** 構造化された中間データモデル、変換設定（ビルダーパターンで構築）。

**期待される出力:** Markdown形式のテキストデータ（std::io::Writeトレイトへ出力）。

#### **2.1.4. 設定・制御機能**

* **ビルダーパターンAPI:** 複雑な設定を安全かつ段階的に構築するFluent Builder API。
  * シート選択（特定シートのみ処理、シート名指定）。
  * 範囲指定（処理するセル範囲の制限）。
  * セル結合戦略の選択（データ重複フィル / HTMLフォールバック）。
  * 日付形式の指定（デフォルト: ISO 8601、カスタム形式指定可能）。
  * 非表示要素の包含制御（非表示シート、行、列をスキップ/含める）。
  * 数式出力モード（キャッシュ値 / 数式文字列）。

**入力:** ビルダーメソッドチェーンによる設定構築。

**期待される出力:** 設定完了済みのConverterインスタンス。

### **2.2. 各機能の入力と期待される出力の概要**

| 機能 | 入力 | 期待される出力 |
|:-----|:-----|:--------------|
| Excelファイル解析 | Excelファイル（XLSX）へのReadアクセス | ワークブック構造、セルデータ、書式情報、結合範囲の構造化データ |
| データ型変換 | 生のセル値（f64、String、Bool）とNumber Format String | 書式適用済みの表示文字列（ISO 8601日付、数値書式適用済み文字列など） |
| Markdown変換 | 構造化された中間データと変換設定 | Markdown形式のテキスト出力（Writeトレイトへ） |
| 設定・制御 | ビルダーメソッドによる設定パラメータ | 設定完了済みのConverterインスタンス |

---

## **3. 非機能要件**

本セクションでは、プログラムが満たすべき品質、性能、制約を定義する。

### **3.1. 性能（Performance）**

* **処理速度:** 
  * 大規模ファイル（1GB以上）を処理可能。
  * 具体的なベンチマーク要件として、既存のPythonベースの変換ツールと比較して、同等以上の処理速度を実現することを目標とする。
  * RAGパイプラインでの実用性を考慮し、バッチ処理でのスループットを最優先とする。
  * 10MB以下のExcelファイルを1秒以内に処理することを目標とする。
  * バッチ処理において、1分あたり少なくとも50ファイル以上の処理を目標とする。
* **メモリ使用量の許容範囲:** 
  * **目標:** ピークメモリ使用量をファイルサイズの10%以下に抑制。
  * **実現手段:** calamineによるストリーミング解析（SAXライクなイベント駆動型処理）と、出力バッファリング（std::io::BufWriter）の組み合わせ。
  * **大規模ファイル対応:** 数ギガバイト単位のファイルでもメモリフットプリントを低く抑える。必要に応じてMemory-Mapped I/O (Mmap) の採用を検討。
  * 10MBのExcelファイル処理時、最大メモリ使用量は100MB以下を目標とする。
* **I/O効率:** 
  * 入出力はstd::io::Read/Writeトレイトを利用することで、ファイルシステムだけでなく、メモリバッファやネットワークストリームにも対応。
  * 出力は自動的にBufWriterでラップし、システムコール数を最小化。

### **3.2. 互換性（Compatibility）**

* **サポートするOS:** 
  * Windows、macOS、Linux（主要ディストリビューション）
  * 将来的にWebAssembly環境への対応も検討可能な設計（stdライブラリへの依存を最小限に）
* **プログラミング言語のバージョン:** 
  * Rust 1.70.0以上（Edition 2021を推奨）
* **依存してはいけないプログラム:** 
  * **必須制約:** C/C++コンパイラや外部ネイティブライブラリ（libpoppler、libtesseractなど）への依存を一切禁止。
  * **Pure-Rustの保証:** すべての依存クレートはPure-Rustであること。calamineを主要な依存関係として採用（Pure-Rust製）。
  * **ZIP処理:** 必要に応じてflate2のrust_backend featureを明示的に指定し、Cバックエンドを回避。

### **3.3. 信頼性（Reliability）**

* **エラーハンドリングのポリシー:** 
  * **エラー返却方式:** Rustの`Result<T, E>`型を使用し、エラーを返す方式を採用します。例外（panic）は、回復不可能なエラーのみに限定します。
  * **構造化エラーの定義:** `thiserror`クレートの採用を推奨し、すべての処理層で発生し得るエラーを構造化します。以下のエラーカテゴリーを明確に区別するXlsxToMdError列挙型を定義します。
    * I/Oエラー（std::io::Error）: ファイル操作や入出力ストリームの問題。#[from]属性による自動変換を適用。
    * 解析エラー（calamine::Error）: 内部パーサーからのエラー。
    * 設定エラー: 不正なシート名や範囲指定。
    * 未サポート機能エラー: サポート範囲外の複雑な書式や構造に遭遇した場合。
  * **エラー報告の明確性:** エラーの原因と場所を正確に伝える（ファイルパス、シート名、セル位置、行番号などのコンテキスト情報を含む）。
  * **パニックの回避:** すべての潜在的なパニックを防ぎ、Result型を返すことでエラーを明示的に処理。std::unwrap()やstd::expect()の使用を最小限に。
  * **エラー処理の一貫性:** すべての公開APIはResult型を返し、エラーを明示的に処理する。パニックを発生させる可能性のある操作は内部実装に隠蔽する。
* **データ整合性の保証:** 
  * 数式評価は行わず、キャッシュされた結果値を使用することで、処理の安定性を確保。
  * 日付エポック（1900年起算/1904年起算）の自動検出と正確な変換。
  * セル結合処理において、データの欠落や重複を防止するロジックの実装。
* **ロバスト性:** 
  * 壊れたファイルや不正な構造への対応（可能な限りのデータを回復）。
  * 空のシート、非表示要素のみで構成されたファイルなどのエッジケースに対する安定した挙動。
  * 広範な境界条件テストの実施。

### **3.4. セキュリティ（Security）**

* **入力値のサニタイズ:** 
  * Excelファイルの解析において、悪意のあるファイルや破損したファイルに対する防御を実装。
  * ZIPアーカイブ構造の検証（ZIP bomb攻撃への対策）: 展開後のサイズやファイル数の上限を設定し、ZIP爆弾攻撃を防ぐ。
  * XML解析時のエンティティ展開攻撃やXXE（XML External Entity）攻撃への対策: quick-xmlやxmlparserのデフォルト設定を活用し、外部エンティティの展開を無効化する。
  * ファイルサイズの上限設定や、深くネストされた構造に対する制限を実装。
  * パストラバーサル攻撃の防止: ファイルパスの検証を行い、意図しないディレクトリへのアクセスを防止する。
* **メモリ安全性:** 
  * Rustの所有権システムにより、バッファオーバーフローやuse-after-freeなどの脆弱性を根本的に排除。
  * unsafeコードの使用を最小限に抑える設計原則。
* **依存関係の管理:** 
  * 依存クレートのセキュリティ監査を定期的に実施。
  * Cargo.lockファイルをコミットし、再現可能なビルドを保証。
  * 既知の脆弱性を持つ依存関係の使用を避け、セキュリティアップデートを迅速に反映。

---

## **4. 制約条件**

本セクションでは、開発における技術的制約と法的制約を定義する。

### **4.1. 使用するプログラミング言語、フレームワーク**

* **プログラミング言語:** 
  * Rust（エディション: 2021、Rust 1.70.0以上、Pure-Rust実装）
  * Pure-Rustの制約により、C/C++コンパイラや外部ネイティブライブラリへの依存を一切禁止。
* **主要な依存クレート:**
  * **calamine:** Excelファイル解析のコアパーサー（Pure-Rust、ストリーミング解析対応、Serdeサポート）
  * **thiserror:** カスタムエラー型の定義（エラーハンドリングの堅牢性向上）
  * **chrono（オプション）:** 日付・時刻の処理（タイムゾーン非依存のNaiveDateTime型を使用）
  * **formato（検討中）:** Number Format Stringの解析（書式適用ロジックの独立サブモジュール化）
  * **log（オプション）:** 標準ロギングファサード
* **標準ライブラリの活用:** 
  * std::io::Read、std::io::Writeトレイトによる柔軟なI/O抽象化
  * std::io::BufWriterによる出力バッファリング
* **フレームワーク:** 
  * 特定のWebフレームワークやアプリケーションフレームワークへの依存は不要（ライブラリクレートとして設計）。

### **4.2. ライセンス（非常に重要）**

* **推奨ライセンス:** 
  * **MIT License** または **Apache License 2.0**（デュアルライセンスも可）
* **選択理由:** 
  * **MIT License:** 
    * シンプルで広く採用されているライセンス
    * 商用利用も自由で、RAGシステムやエンタープライズ環境での利用を妨げない
    * 最小限の制約で最大限の互換性を提供
  * **Apache License 2.0:** 
    * パテント条項を含み、より包括的な法的保護を提供
    * Rustエコシステムで広く採用されている（Rustコンパイラ自体もApache 2.0/MITデュアルライセンス）
    * コントリビューターへの明示的なパテント付与を含む
* **依存クレートのライセンス互換性:** 
  * calamineはMIT/Apache 2.0デュアルライセンスのため、どちらのライセンスも互換性がある
  * すべての依存クレートのライセンスを確認し、互換性を保証する必要がある
* **ライセンスの明確化:** 
  * `LICENSE`ファイルをリポジトリのルートに配置する
  * `Cargo.toml`の`license`フィールドに適切なライセンスを指定（例: "MIT OR Apache-2.0"）
  * README.mdにライセンス情報を明記する
  * ソースコードの各ファイルにライセンスヘッダーを付与することを検討する

---

## **5. 実装上の技術的詳細（参考）**

### **5.1. アーキテクチャ概要**

本クレートは、以下の設計原則に基づいて構築される。

* **ストリーミング解析:** DOM方式ではなく、SAXライクなイベント駆動型のストリーミング処理を採用。メモリ効率を最優先。
* **中間データ構造:** スパースなセルデータとセル結合情報から、Markdownテーブルに必要な稠密なグリッド構造へ再構築する中間処理層を実装。
* **ビルダーパターン:** 複雑な設定を安全かつエルゴノミックに扱うFluent Builder APIを提供。
* **懸念の分離:** 書式適用ロジックは独立したサブモジュールとして設計し、テスト容易性とメンテナンス性を向上。

### **5.2. データ変換戦略の詳細**

#### **5.2.1. セル結合処理**

Markdownではセル結合（rowspan/colspan）をネイティブにサポートしていないため、以下の戦略を提供。

* **デフォルト戦略（データ重複フィル）:** 
  * 結合セル範囲（例: A1:C1）を検出後、親セル（A1）の値を結合範囲内のすべての論理的な子セル（A1, B1, C1）に複製。
  * 利点: 純粋なMarkdownテーブル構造を維持。LLMが理解しやすく、トークン効率が高い。
* **代替戦略（HTMLフォールバック）:** 
  * 結合セルを含むテーブル全体をHTMLテーブル（<table>タグ、rowspan/colspan属性）として出力。
  * 利点: 構造的な忠実性を完全に維持。
  * 欠点: トークン数が増加し、RAGシステムでのコスト増やコンテキスト減少の可能性。

#### **5.2.2. 日付・数値の書式適用**

* Excelのシリアル値（f64）をNumber Format Stringに基づいて正確な表示文字列に変換。
* デフォルトでISO 8601形式（YYYY-MM-DD）を使用。ユーザー指定形式も対応。
* 1900年起算と1904年起算の両エポックに対応するテストを必須とする。

### **5.3. 実装ロードマップ**

1. **フェーズ I: コア機能と性能PoC（60%）**
   * calamine統合による基本データ（文字列、数値）のストリーミング抽出。
   * ビルダーパターンとカスタムエラー型の確立。
   * セル結合情報抽出と、論理グリッド再構築によるデータ重複フィルロジックのPoC実装。

2. **フェーズ II: データインテグリティの強化（30%）**
   * Excel Number Format Parserのサブモジュール化と、日付/数値の正確な文字列変換の実装。
   * ハイパーリンクとリッチテキストの最小限の書式マッピング実装。

3. **フェーズ III: ロバスト性とベンチマーク（10%）**
   * 大規模データセット（1GB以上）に対するメモリ効率と速度のベンチマークを実施。
   * 広範な境界条件およびエッジケースに対するロバスト性の確保。

---

#### **引用文献**

1. GitHub - tafia/calamine: A pure Rust Excel/OpenDocument SpreadSheets file reader, 11月 19, 2025にアクセス、 [https://github.com/tafia/calamine](https://github.com/tafia/calamine)  
2. Builders in Rust - Shuttle.dev, 11月 19, 2025にアクセス、 [https://www.shuttle.dev/blog/2022/06/09/the-builder-pattern](https://www.shuttle.dev/blog/2022/06/09/the-builder-pattern)  
3. former - Rust - Docs.rs, 11月 19, 2025にアクセス、 [https://docs.rs/former](https://docs.rs/former)  
4. ExcelDateTime in rust_xlsxwriter - Rust - Docs.rs, 11月 19, 2025にアクセス、 [https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html](https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html)  
5. merged cell convert issue,both in excel and pptx · Issue #167 · microsoft/markitdown, 11月 19, 2025にアクセス、 [https://github.com/microsoft/markitdown/issues/167](https://github.com/microsoft/markitdown/issues/167)  
6. Can I merge table rows in markdown - github - Stack Overflow, 11月 19, 2025にアクセス、 [https://stackoverflow.com/questions/46621765/can-i-merge-table-rows-in-markdown](https://stackoverflow.com/questions/46621765/can-i-merge-table-rows-in-markdown)  
7. excel_reader — Rust parser // Lib.rs, 11月 19, 2025にアクセス、 [https://lib.rs/crates/excel_reader](https://lib.rs/crates/excel_reader)  
8. excel_reader - crates.io: Rust Package Registry, 11月 19, 2025にアクセス、 [https://crates.io/crates/excel_reader](https://crates.io/crates/excel_reader)  
9. excel - Keywords - crates.io: Rust Package Registry, 11月 19, 2025にアクセス、 [https://crates.io/keywords/excel](https://crates.io/keywords/excel)  
10. rust_xlsxwriter - Rust - Docs.rs, 11月 19, 2025にアクセス、 [https://docs.rs/rust_xlsxwriter](https://docs.rs/rust_xlsxwriter)  
11. BufWriter in std::io - Rust, 11月 19, 2025にアクセス、 [https://doc.rust-lang.org/std/io/struct.BufWriter.html](https://doc.rust-lang.org/std/io/struct.BufWriter.html)  
12. File in std::fs - Rust, 11月 19, 2025にアクセス、 [https://doc.rust-lang.org/std/fs/struct.File.html](https://doc.rust-lang.org/std/fs/struct.File.html)  
13. formato - crates.io: Rust Package Registry, 11月 19, 2025にアクセス、 [https://crates.io/crates/formato](https://crates.io/crates/formato)  
14. Processing large xlsx file - Stack Overflow, 11月 19, 2025にアクセス、 [https://stackoverflow.com/questions/4897766/processing-large-xlsx-file](https://stackoverflow.com/questions/4897766/processing-large-xlsx-file)  
15. Efficiently Handling Large Files and Long Connections with Streaming Responses in Rust Web Frameworks | Leapcell, 11月 19, 2025にアクセス、 [https://leapcell.io/blog/efficiently-handling-large-files-and-long-connections-with-streaming-responses-in-rust-web-frameworks](https://leapcell.io/blog/efficiently-handling-large-files-and-long-connections-with-streaming-responses-in-rust-web-frameworks)  
16. [Solved] How to scan efficiently big binary streams/files? - help - Rust Users Forum, 11月 19, 2025にアクセス、 [https://users.rust-lang.org/t/solved-how-to-scan-efficiently-big-binary-streams-files/6345](https://users.rust-lang.org/t/solved-how-to-scan-efficiently-big-binary-streams-files/6345)  
17. Managing Large Data between Memory and Disk - help - Rust Users Forum, 11月 19, 2025にアクセス、 [https://users.rust-lang.org/t/managing-large-data-between-memory-and-disk/63155](https://users.rust-lang.org/t/managing-large-data-between-memory-and-disk/63155)  
18. markdownify - crates.io: Rust Package Registry, 11月 19, 2025にアクセス、 [https://crates.io/crates/markdownify](https://crates.io/crates/markdownify)  
19. How do I extract data from a cell and order the cells alphabetically? - Stack Overflow, 11月 19, 2025にアクセス、 [https://stackoverflow.com/questions/34039387/how-do-i-extract-data-from-a-cell-and-order-the-cells-alphabetically](https://stackoverflow.com/questions/34039387/how-do-i-extract-data-from-a-cell-and-order-the-cells-alphabetically)  
20. How to convert Excel date format to proper date in R - Stack Overflow, 11月 19, 2025にアクセス、 [https://stackoverflow.com/questions/43230470/how-to-convert-excel-date-format-to-proper-date-in-r](https://stackoverflow.com/questions/43230470/how-to-convert-excel-date-format-to-proper-date-in-r)  
21. Formula in rust_xlsxwriter - Rust - Docs.rs, 11月 19, 2025にアクセス、 [https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Formula.html](https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Formula.html)  
22. Convert Spreadsheet - Table to Markdown, 11月 19, 2025にアクセス、 [https://tabletomarkdown.com/convert-spreadsheet-to-markdown/](https://tabletomarkdown.com/convert-spreadsheet-to-markdown/)  
23. formula - crates.io: Rust Package Registry, 11月 19, 2025にアクセス、 [https://crates.io/crates/formula](https://crates.io/crates/formula)  
24. How to decide if a cargo crate is "production ready"? - help - Rust Users Forum, 11月 19, 2025にアクセス、 [https://users.rust-lang.org/t/how-to-decide-if-a-cargo-crate-is-production-ready/68976](https://users.rust-lang.org/t/how-to-decide-if-a-cargo-crate-is-production-ready/68976)  
25. Conditional Formating with OLE Excel - PerlMonks, 11月 19, 2025にアクセス、 [https://www.perlmonks.org/?node_id=930945](https://www.perlmonks.org/?node_id=930945)  
26. How to keep value of merged cells in each cell? - Stack Overflow, 11月 19, 2025にアクセス、 [https://stackoverflow.com/questions/6464265/how-to-keep-value-of-merged-cells-in-each-cell](https://stackoverflow.com/questions/6464265/how-to-keep-value-of-merged-cells-in-each-cell)  
27. Concerns Regarding the tables in markdown output changes in 2024-07-31-preview, 11月 19, 2025にアクセス、 [https://learn.microsoft.com/en-au/answers/questions/2111913/concerns-regarding-the-tables-in-markdown-output-c](https://learn.microsoft.com/en-au/answers/questions/2111913/concerns-regarding-the-tables-in-markdown-output-c)  
28. fastest way to merge duplicate cells in without looping Excel - Stack Overflow, 11月 19, 2025にアクセス、 [https://stackoverflow.com/questions/45737497/fastest-way-to-merge-duplicate-cells-in-without-looping-excel](https://stackoverflow.com/questions/45737497/fastest-way-to-merge-duplicate-cells-in-without-looping-excel)  
29. Turning the visibility of chart series on/off using excel Macros/vba - Stack Overflow, 11月 19, 2025にアクセス、 [https://stackoverflow.com/questions/29703066/turning-the-visibility-of-chart-series-on-off-using-excel-macros-vba](https://stackoverflow.com/questions/29703066/turning-the-visibility-of-chart-series-on-off-using-excel-macros-vba)  
30. The rust_xlsxwriter crate for creating xlsx files : r/rust - Reddit, 11月 19, 2025にアクセス、 [https://www.reddit.com/r/rust/comments/1ezs66n/the_rust_xlsxwriter_crate_for_creating_xlsx_files/](https://www.reddit.com/r/rust/comments/1ezs66n/the_rust_xlsxwriter_crate_for_creating_xlsx_files/)  
31. Errors in Rust can now be handled more ergonomically, cleanly, and simply: introducing a new error crate - Rust Users Forum, 11月 19, 2025にアクセス、 [https://users.rust-lang.org/t/errors-in-rust-can-now-be-handled-more-ergonomically-cleanly-and-simply-introducing-a-new-error-crate/51527](https://users.rust-lang.org/t/errors-in-rust-can-now-be-handled-more-ergonomically-cleanly-and-simply-introducing-a-new-error-crate/51527)  
32. Practical guide to Error Handling in Rust - Dev State, 11月 19, 2025にアクセス、 [https://dev-state.com/posts/error_handling/](https://dev-state.com/posts/error_handling/)  
33. ErrorKind in std::io - Rust, 11月 19, 2025にアクセス、 [https://doc.rust-lang.org/std/io/enum.ErrorKind.html](https://doc.rust-lang.org/std/io/enum.ErrorKind.html)  
34. unsupported data type: &[] error on GORM field where custom Valuer returns nil?, 11月 19, 2025にアクセス、 [https://stackoverflow.com/questions/64035165/unsupported-data-type-error-on-gorm-field-where-custom-valuer-returns-nil](https://stackoverflow.com/questions/64035165/unsupported-data-type-error-on-gorm-field-where-custom-valuer-returns-nil)
