//! Security Module
//!
//! セキュリティ対策を実装するモジュール。
//! ZIP bomb攻撃、XXE攻撃、パストラバーサル攻撃などへの対策を提供します。

/// セキュリティ設定
///
/// ファイル処理時のセキュリティ制限を定義します。
#[derive(Debug, Clone)]
pub(crate) struct SecurityConfig {
    /// 展開後の最大サイズ（バイト）
    /// デフォルト: 1GB (1_073_741_824 bytes)
    pub max_decompressed_size: u64,
    /// ZIPアーカイブ内の最大ファイル数
    /// デフォルト: 10000
    pub max_file_count: usize,
    /// 単一ファイルの最大サイズ（バイト）
    /// デフォルト: 100MB (104_857_600 bytes)
    pub max_file_size: u64,
    /// 入力ファイルの最大サイズ（バイト）
    /// デフォルト: 2GB (2_147_483_648 bytes)
    pub max_input_file_size: u64,
}

impl Default for SecurityConfig {
    fn default() -> Self {
        Self {
            max_decompressed_size: 1_073_741_824, // 1GB
            max_file_count: 10_000,
            max_file_size: 104_857_600,         // 100MB
            max_input_file_size: 2_147_483_648, // 2GB
        }
    }
}

impl SecurityConfig {
    /// デフォルトのセキュリティ設定を作成
    #[allow(dead_code)]
    pub fn new() -> Self {
        Self::default()
    }
}

/// ファイルパスの検証
///
/// パストラバーサル攻撃を防ぐため、ファイルパスを検証します。
///
/// # 引数
///
/// * `path` - 検証するファイルパス
///
/// # 戻り値
///
/// * `Ok(())` - パスが安全な場合
/// * `Err(String)` - パスが危険な場合（`..`や絶対パスを含む）
pub(crate) fn validate_zip_path(path: &str) -> Result<(), String> {
    // 空のパスは拒否
    if path.is_empty() {
        return Err("Empty path is not allowed".to_string());
    }

    // 絶対パスを拒否（Windows形式の`C:\`やUnix形式の`/`で始まるパス）
    if path.starts_with('/') || path.starts_with("C:\\") || path.starts_with("c:\\") {
        return Err(format!("Absolute path is not allowed: {}", path));
    }

    // `..`を含むパスを拒否（ディレクトリトラバーサル攻撃）
    if path.contains("..") {
        return Err(format!("Path traversal detected: {}", path));
    }

    // `\`を含むパスを拒否（Windows形式のパスセパレータ）
    if path.contains('\\') {
        return Err(format!("Backslash in path is not allowed: {}", path));
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_validate_zip_path_valid() {
        assert!(validate_zip_path("xl/workbook.xml").is_ok());
        assert!(validate_zip_path("xl/worksheets/sheet1.xml").is_ok());
        assert!(validate_zip_path("xl/sharedStrings.xml").is_ok());
    }

    #[test]
    fn test_validate_zip_path_empty() {
        assert!(validate_zip_path("").is_err());
    }

    #[test]
    fn test_validate_zip_path_absolute_unix() {
        assert!(validate_zip_path("/etc/passwd").is_err());
        assert!(validate_zip_path("/xl/workbook.xml").is_err());
    }

    #[test]
    fn test_validate_zip_path_absolute_windows() {
        assert!(validate_zip_path("C:\\Windows\\system32").is_err());
        assert!(validate_zip_path("c:\\xl\\workbook.xml").is_err());
    }

    #[test]
    fn test_validate_zip_path_traversal() {
        assert!(validate_zip_path("../etc/passwd").is_err());
        assert!(validate_zip_path("xl/../../etc/passwd").is_err());
        assert!(validate_zip_path("xl/..").is_err());
        assert!(validate_zip_path("..").is_err());
    }

    #[test]
    fn test_validate_zip_path_backslash() {
        assert!(validate_zip_path("xl\\workbook.xml").is_err());
    }
}
