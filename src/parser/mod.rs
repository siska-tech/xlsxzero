//! Parser Module
//!
//! calamineを使用したExcelファイル解析の基礎実装。
//! ストリーミング処理により、メモリ効率的にセルデータを抽出します。

mod metadata;
mod workbook;

pub(crate) use metadata::XlsxMetadataParser;
pub(crate) use workbook::WorkbookParser;
