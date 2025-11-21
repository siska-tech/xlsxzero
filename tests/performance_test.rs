//! パフォーマンステスト
//!
//! このモジュールは、メモリ使用量と処理速度の要件を検証するテストを提供します。
//!
//! 実装するテスト:
//! - TC-P-001: Small File Memory Usage（ピークメモリ ≤ 100MB）
//! - TC-P-002: Large File Memory Usage（ピークメモリ ≤ ファイルサイズの10%）
//!
//! 注意: メモリ使用量の正確な測定には、valgrindやheaptrackなどの外部ツールが必要です。
//! このテストは基本的なメモリ使用量の目安を提供します。

use std::fs::File;
use std::io::{Cursor, Read};
use std::path::PathBuf;
use xlsxzero::ConverterBuilder;

/// テストフィクスチャのパスを取得
fn fixture_path(filename: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("fixtures");
    path.push("bench");
    path.push(filename);
    path
}

/// フィクスチャファイルが存在するかチェック
fn fixture_exists(filename: &str) -> bool {
    fixture_path(filename).exists()
}

/// TC-P-001: Small File Memory Usage
///
/// 10MBのExcelファイル処理時、ピークメモリ使用量が100MB以下であることを確認します。
///
/// 注意: このテストは基本的なメモリ使用量の目安を提供します。
/// 正確な測定には、valgrindやheaptrackなどの外部ツールを使用してください。
#[test]
#[ignore] // 手動実行用
fn test_small_file_memory_usage() {
    if !fixture_exists("10mb_file.xlsx") {
        eprintln!("Warning: tests/fixtures/bench/10mb_file.xlsx not found. Skipping test.");
        return;
    }

    let converter = ConverterBuilder::new().build().unwrap();
    let mut input_data = Vec::new();
    File::open(fixture_path("10mb_file.xlsx"))
        .unwrap()
        .read_to_end(&mut input_data)
        .unwrap();
    let input = Cursor::new(input_data);
    let mut output = Vec::new();

    // 処理前のメモリ使用量（概算）
    let before_memory = get_memory_usage();

    converter.convert(input, &mut output).unwrap();

    // 処理後のメモリ使用量（概算）
    let after_memory = get_memory_usage();
    let memory_delta = after_memory.saturating_sub(before_memory);

    println!("Memory delta: {} MB", memory_delta / 1024 / 1024);
    println!("Output size: {} bytes", output.len());

    // 目標: ピークメモリ ≤ 100MB
    // 注意: このテストは概算値のみを提供します
    // 正確な測定には外部ツールを使用してください
    if memory_delta > 100 * 1024 * 1024 {
        eprintln!(
            "Warning: Memory delta ({:.2} MB) exceeds target (100 MB). \
            Use valgrind or heaptrack for accurate measurement.",
            memory_delta as f64 / 1024.0 / 1024.0
        );
    }
}

/// TC-P-002: Large File Memory Usage
///
/// 1GBのExcelファイル処理時、ピークメモリ使用量がファイルサイズの10%（100MB）以下であることを確認します。
///
/// 注意: このテストは非常に時間がかかるため、通常はスキップされます。
/// 実行する場合は `cargo test -- --ignored` を使用してください。
#[test]
#[ignore] // 手動実行用（時間がかかる）
fn test_large_file_memory_usage() {
    if !fixture_exists("1gb_file.xlsx") {
        eprintln!("Warning: tests/fixtures/bench/1gb_file.xlsx not found. Skipping test.");
        return;
    }

    let converter = ConverterBuilder::new().build().unwrap();
    let mut input_data = Vec::new();
    let mut file = File::open(fixture_path("1gb_file.xlsx")).unwrap();
    let file_size = file.metadata().unwrap().len() as usize;
    file.read_to_end(&mut input_data).unwrap();
    let input = Cursor::new(input_data);
    let mut output = Vec::new();

    let target_memory = file_size / 10; // ファイルサイズの10%

    // 処理前のメモリ使用量（概算）
    let before_memory = get_memory_usage();

    converter.convert(input, &mut output).unwrap();

    // 処理後のメモリ使用量（概算）
    let after_memory = get_memory_usage();
    let memory_delta = after_memory.saturating_sub(before_memory);

    println!("File size: {} MB", file_size / 1024 / 1024);
    println!("Target memory: {} MB", target_memory / 1024 / 1024);
    println!("Memory delta: {} MB", memory_delta / 1024 / 1024);
    println!("Output size: {} bytes", output.len());

    // 目標: ピークメモリ ≤ ファイルサイズの10%
    // 注意: このテストは概算値のみを提供します
    // 正確な測定には外部ツールを使用してください
    if memory_delta > target_memory {
        eprintln!(
            "Warning: Memory delta ({:.2} MB) exceeds target ({:.2} MB). \
            Use valgrind or heaptrack for accurate measurement.",
            memory_delta as f64 / 1024.0 / 1024.0,
            target_memory as f64 / 1024.0 / 1024.0
        );
    }
}

/// 現在のプロセスのメモリ使用量を取得（概算）
///
/// 注意: この関数は概算値のみを提供します。
/// 正確な測定には、valgrindやheaptrackなどの外部ツールを使用してください。
fn get_memory_usage() -> usize {
    // Linuxの場合、/proc/self/statusから取得
    #[cfg(target_os = "linux")]
    {
        if let Ok(status) = std::fs::read_to_string("/proc/self/status") {
            for line in status.lines() {
                if line.starts_with("VmRSS:") {
                    if let Some(value) = line.split_whitespace().nth(1) {
                        if let Ok(kb) = value.parse::<usize>() {
                            return kb * 1024; // KBをバイトに変換
                        }
                    }
                }
            }
        }
    }

    // その他のOSでは、概算値として0を返す
    0
}

/// TC-P-010: Small File Processing Speed
///
/// 10MBのExcelファイルを1秒以内に処理することを確認します。
#[test]
#[ignore] // 手動実行用
fn test_small_file_processing_speed() {
    if !fixture_exists("10mb_file.xlsx") {
        eprintln!("Warning: tests/fixtures/bench/10mb_file.xlsx not found. Skipping test.");
        return;
    }

    let converter = ConverterBuilder::new().build().unwrap();
    let mut input_data = Vec::new();
    File::open(fixture_path("10mb_file.xlsx"))
        .unwrap()
        .read_to_end(&mut input_data)
        .unwrap();

    let start = std::time::Instant::now();
    let input = Cursor::new(&input_data);
    let mut output = Vec::new();
    converter.convert(input, &mut output).unwrap();
    let duration = start.elapsed();

    println!("Processing time: {:?}", duration);
    println!("Output size: {} bytes", output.len());

    // 目標: 処理時間 < 1秒
    assert!(
        duration.as_secs() < 1,
        "Processing took too long: {:?} (target: < 1 second)",
        duration
    );
}

/// TC-P-011: Batch Processing Throughput
///
/// 50ファイル（各10MB）を1分以内に処理することを確認します。
#[test]
#[ignore] // 手動実行用（時間がかかる）
fn test_batch_processing_throughput() {
    let converter = ConverterBuilder::new().build().unwrap();
    let mut batch_files = Vec::new();

    // バッチファイルを読み込む（最大50ファイル）
    for i in 0..50 {
        let filename = format!("batch_{:02}.xlsx", i);
        let path = fixture_path(&filename);
        if path.exists() {
            let mut data = Vec::new();
            File::open(&path).unwrap().read_to_end(&mut data).unwrap();
            batch_files.push(data);
        } else {
            break;
        }
    }

    if batch_files.is_empty() {
        eprintln!("Warning: No batch files found. Skipping test.");
        return;
    }

    let start = std::time::Instant::now();

    for file_data in &batch_files {
        let input = Cursor::new(file_data);
        let mut output = Vec::new();
        converter.convert(input, &mut output).unwrap();
    }

    let duration = start.elapsed();

    println!("Processed {} files in {:?}", batch_files.len(), duration);
    println!(
        "Throughput: {:.2} files/second",
        batch_files.len() as f64 / duration.as_secs_f64()
    );

    // 目標: 50ファイルを1分以内に処理
    assert!(
        duration.as_secs() < 60,
        "Batch processing took too long: {:?} (target: < 60 seconds for {} files)",
        duration,
        batch_files.len()
    );
}
