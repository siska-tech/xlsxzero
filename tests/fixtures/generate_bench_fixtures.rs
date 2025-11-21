//! ベンチマーク用テストフィクスチャ生成スクリプト
//!
//! このスクリプトは、ベンチマークテストで使用する大規模なExcelファイルを生成します。
//!
//! 生成するファイル:
//! - 10MBファイル: tests/fixtures/bench/10mb_file.xlsx
//! - 100MBファイル: tests/fixtures/bench/100mb_file.xlsx
//! - 1GBファイル: tests/fixtures/bench/1gb_file.xlsx（オプション）
//! - バッチファイル: tests/fixtures/bench/batch_00.xlsx ～ batch_49.xlsx（各10MB）

use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
use std::fs;
use std::path::PathBuf;

/// ベンチマークフィクスチャのディレクトリパス
fn bench_fixture_dir() -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("fixtures");
    path.push("bench");
    path
}

/// 指定されたサイズ（バイト）に近いExcelファイルを生成
///
/// セルにデータを書き込んで、目標サイズに近づけます。
fn generate_file_by_size(filename: &str, target_size: usize) -> Result<(), XlsxError> {
    let dir = bench_fixture_dir();
    fs::create_dir_all(&dir)?;

    let mut path = dir.clone();
    path.push(filename);

    println!("Generating {} (target: {} bytes)...", filename, target_size);

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // セルに書き込むデータ（約100バイト/セル）
    let cell_data = "This is a test cell with some data to make the file larger. ";
    let cell_size = cell_data.len();

    // 目標サイズに達するまでセルを書き込む
    let mut current_size = 0;
    let mut row = 0;
    let mut col = 0;

    // Excelファイルのオーバーヘッドを考慮（約10KB）
    let overhead = 10_000;
    let target_data_size = target_size.saturating_sub(overhead);

    while current_size < target_data_size {
        worksheet.write_string(row, col, cell_data)?;

        current_size += cell_size;
        col += 1;

        // 列が多すぎる場合は次の行へ
        if col >= 1000 {
            col = 0;
            row += 1;
        }
    }

    workbook.save(&path)?;

    // 実際のファイルサイズを確認
    let actual_size = fs::metadata(&path)?.len() as usize;
    println!(
        "Generated {}: {} bytes (target: {} bytes, ratio: {:.2}%)",
        filename,
        actual_size,
        target_size,
        (actual_size as f64 / target_size as f64) * 100.0
    );

    Ok(())
}

/// バッチファイルを生成（各10MB、50ファイル）
fn generate_batch_files() -> Result<(), XlsxError> {
    println!("Generating batch files (50 files, 10MB each)...");

    for i in 0..50 {
        let filename = format!("batch_{:02}.xlsx", i);
        generate_file_by_size(&filename, 10 * 1024 * 1024)?;
    }

    println!("Generated 50 batch files.");
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Generating benchmark fixtures...");
    println!("This may take a while, especially for large files.\n");

    // 10MBファイル
    generate_file_by_size("10mb_file.xlsx", 10 * 1024 * 1024)?;

    // 100MBファイル
    generate_file_by_size("100mb_file.xlsx", 100 * 1024 * 1024)?;

    // 1GBファイル（オプション、時間がかかるため）
    let generate_1gb = std::env::var("GENERATE_1GB").is_ok();
    if generate_1gb {
        println!("\nGenerating 1GB file (this will take a very long time)...");
        generate_file_by_size("1gb_file.xlsx", 1024 * 1024 * 1024)?;
    } else {
        println!("\nSkipping 1GB file generation (set GENERATE_1GB=true to generate).");
    }

    // バッチファイル
    generate_batch_files()?;

    println!("\nAll benchmark fixtures generated successfully!");
    Ok(())
}

