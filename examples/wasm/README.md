# xlsxzero WASM Example

このディレクトリには、ブラウザ上で動作する xlsxzero の WebAssembly (WASM) サンプルが含まれています。

## 概要

このサンプルでは、ブラウザ上で Excel ファイル (.xlsx, .xlsm, .xlsb) を Markdown 形式に変換する機能を提供します。

## 必要なツール

1. **Rust** (最新の stable 版)
2. **wasm-pack** - WebAssembly パッケージツール

### wasm-pack のインストール

```bash
curl https://rustwasm.github.io/wasm-pack/installer/init.sh -sSf | sh
```

または、cargo を使用してインストール:

```bash
cargo install wasm-pack
```

## ビルド手順

### 1. WASM モジュールのビルド

プロジェクトのルートディレクトリ（`xlsxzero`）で以下のコマンドを実行:

```bash
cd examples/wasm
wasm-pack build --target web --out-dir pkg
```

これにより、`examples/wasm/pkg/` ディレクトリに WASM モジュールと JavaScript バインディングが生成されます。

### 2. ローカルサーバーの起動

ブラウザのセキュリティ制約により、`file://` プロトコルでは WASM を読み込めません。
そのため、ローカルサーバーを起動する必要があります。

#### Python を使用する場合

```bash
# Python 3
python -m http.server 8000

# Python 2
python -m SimpleHTTPServer 8000
```

#### Node.js を使用する場合

```bash
# http-server をインストール
npm install -g http-server

# サーバーを起動
http-server -p 8000
```

#### PHP を使用する場合

```bash
php -S localhost:8000
```

### 3. ブラウザで開く

ブラウザで以下の URL を開いてください:

```text
http://localhost:8000
```

## 使用方法

1. **ファイルの選択**
   - 「ファイルを選択」ボタンをクリックして Excel ファイルを選択
   - または、ファイルをドラッグ＆ドロップ

2. **オプションの設定（オプション）**
   - **シート選択**: 特定のシートのみを変換したい場合、シート番号（0始まり）を入力
   - **結合セル戦略**: データ複製または HTML フォールバックを選択
   - **日付形式**: ISO8601（空白）またはカスタム形式を指定

3. **変換実行**
   - 「変換実行」ボタンをクリック

4. **結果の利用**
   - 変換結果が表示されます
   - 「コピー」ボタンでクリップボードにコピー
   - 「ダウンロード」ボタンで Markdown ファイルとしてダウンロード

## 機能

- ✅ Excel ファイル (.xlsx, .xlsm, .xlsb) の読み込み
- ✅ Markdown への変換
- ✅ ドラッグ＆ドロップ対応
- ✅ カスタムオプション設定（シート選択、結合セル戦略、日付形式）
- ✅ 変換結果のコピー・ダウンロード
- ✅ エラーハンドリング

## トラブルシューティング

### WASM モジュールが読み込めない

- ローカルサーバーを使用していることを確認してください（`file://` プロトコルでは動作しません）
- ブラウザのコンソールでエラーメッセージを確認してください

### ビルドエラー

- `wasm-bindgen` が正しくインストールされているか確認してください
- プロジェクトのルートで `cargo build` が成功することを確認してください

### 変換エラー

- 有効な Excel ファイルを選択しているか確認してください
- ファイルサイズが大きすぎる場合（50MB 以上）、エラーになる可能性があります

## ファイル構成

```text
examples/wasm/
├── Cargo.toml          # WASM プロジェクトの設定
├── src/
│   └── lib.rs          # WASM バインディングコード
├── index.html          # HTML UI
├── index.js            # JavaScript コード
├── README.md           # このファイル
└── pkg/                # ビルド出力（wasm-pack build で生成）
    ├── xlsxzero_wasm.js
    ├── xlsxzero_wasm_bg.wasm
    └── ...
```

## API

### `convert_excel_to_markdown(excel_bytes: &[u8]) -> Result<String, String>`

デフォルト設定で Excel ファイルを Markdown に変換します。

### `convert_excel_to_markdown_custom`

```rust
convert_excel_to_markdown_custom(
    excel_bytes: &[u8],
    sheet_index: Option<usize>,
    merge_strategy: Option<String>,
    date_format: Option<String>
) -> Result<String, String>
```

カスタムオプションを指定して Excel ファイルを Markdown に変換します。

- `sheet_index`: シート番号（0始まり）、`null` で全シート
- `merge_strategy`: `"data_duplication"` または `"html_fallback"`
- `date_format`: `"iso8601"` またはカスタム形式文字列（例: `"%Y年%m月%d日"`）

### `get_version() -> String`

xlsxzero のバージョン情報を返します。

## サーバーへのデプロイ

### クイックスタート

最も簡単な方法は **GitHub Pages** を使用することです：

1. **ビルドスクリプトを実行**:
   ```bash
   cd examples/wasm
   ./deploy.sh
   ```

2. **GitHub Pages 用にファイルをコピー**:
   ```bash
   # プロジェクトルートで実行
   mkdir -p docs/wasm
   cp examples/wasm/index.html docs/wasm/
   cp examples/wasm/index.js docs/wasm/
   cp -r examples/wasm/pkg docs/wasm/
   ```

3. **GitHub にプッシュして Pages を有効化**:
   - リポジトリの Settings → Pages
   - Source を `/docs` に設定

詳細は以下を参照してください。

### 方法1: GitHub Pages（推奨）

GitHub Pages を使用すると、GitHub リポジトリから自動的にデプロイできます。

#### 手動デプロイ

1. **ビルドを実行**:
   ```bash
   cd examples/wasm
   wasm-pack build --target web --out-dir pkg
   ```

2. **デプロイ用ディレクトリを作成**:
   ```bash
   # プロジェクトルートで実行
   mkdir -p docs/wasm
   cp examples/wasm/index.html docs/wasm/
   cp examples/wasm/index.js docs/wasm/
   cp -r examples/wasm/pkg docs/wasm/
   ```

3. **GitHub リポジトリにプッシュ**:
   ```bash
   git add docs/wasm
   git commit -m "Deploy WASM example to GitHub Pages"
   git push
   ```

4. **GitHub Pages を有効化**:
   - GitHub リポジトリの Settings → Pages
   - Source を "Deploy from a branch" に設定
   - Branch を `main` (または `master`)、フォルダを `/docs` に設定
   - Save をクリック

5. **アクセス**:
   - `https://<your-username>.github.io/xlsxzero/wasm/` でアクセス可能

#### 自動デプロイ（GitHub Actions）

`.github/workflows/deploy-wasm.yml` を作成すると、プッシュ時に自動デプロイされます（後述）。

### 方法2: Netlify

1. **Netlify アカウントを作成**: https://www.netlify.com/

2. **ビルドコマンドを設定**:
   ```bash
   cd examples/wasm && wasm-pack build --target web --out-dir pkg
   ```

3. **公開ディレクトリを設定**: `examples/wasm`

4. **環境変数**（必要に応じて）:
   - `RUST_VERSION`: `stable`

5. **デプロイ**: Netlify が自動的にビルドとデプロイを実行

### 方法3: Vercel

1. **Vercel アカウントを作成**: https://vercel.com/

2. **プロジェクトをインポート**: GitHub リポジトリを選択

3. **ビルド設定**:
   - Build Command: `cd examples/wasm && wasm-pack build --target web --out-dir pkg`
   - Output Directory: `examples/wasm`

4. **デプロイ**: Vercel が自動的にビルドとデプロイを実行

### 方法4: 手動デプロイ（任意のサーバー）

1. **ビルドを実行**:
   ```bash
   cd examples/wasm
   wasm-pack build --target web --out-dir pkg
   ```

2. **ファイルをアップロード**:
   - `index.html`
   - `index.js`
   - `pkg/` ディレクトリ全体
   
   を Web サーバーの公開ディレクトリにアップロード

3. **MIME タイプの設定**:
   - `.wasm` ファイルに対して `application/wasm` を設定
   - `.js` ファイルに対して `application/javascript` を設定

   **Apache (.htaccess)**:
   ```apache
   AddType application/wasm .wasm
   AddType application/javascript .js
   ```

   **Nginx (nginx.conf)**:
   ```nginx
   location ~ \.wasm$ {
       add_header Content-Type application/wasm;
   }
   ```

### デプロイ方法の比較

| 方法         | 難易度        | 自動デプロイ        | 無料プラン   | 推奨用途                   |
| ------------ | ------------- | ------------------- | ------------ | -------------------------- |
| GitHub Pages | ⭐ 簡単       | ✅ (GitHub Actions) | ✅           | オープンソースプロジェクト |
| Netlify      | ⭐⭐ 普通     | ✅                  | ✅           | 簡単なデプロイ             |
| Vercel       | ⭐⭐ 普通     | ✅                  | ✅           | モダンなデプロイ           |
| 手動デプロイ | ⭐⭐⭐ やや難 | ❌                  | サーバー次第 | カスタムサーバー           |

### 注意事項

- **CORS 設定**: 外部からアクセスする場合は、適切な CORS ヘッダーを設定してください
- **HTTPS**: 本番環境では HTTPS を使用することを推奨します
- **ファイルサイズ**: WASM ファイルは大きくなる可能性があるため、適切な圧縮設定を検討してください
- **GitHub Actions**: `.github/workflows/deploy-wasm.yml` が設定されていれば、プッシュ時に自動デプロイされます

