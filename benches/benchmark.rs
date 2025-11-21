//! パフォーマンスベンチマーク
//!
//! このモジュールは、xlsxzeroクレートのパフォーマンスを測定するためのベンチマークを提供します。
//!
//! 実装するベンチマーク:
//! - TC-P-010: Small File Processing Speed（< 1秒）
//! - TC-P-011: Batch Processing Throughput（50ファイル/分以上）
//!
//! メモリ使用量の測定は別途、valgrindやheaptrackなどのツールを使用してください。

use criterion::{black_box, criterion_group, criterion_main, Criterion, Throughput};
use std::fs::File;
use std::io::{Cursor, Read, Seek, Write};
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

/// フィクスチャファイルを読み込む（メモリに読み込む）
fn load_fixture(filename: &str) -> Result<Vec<u8>, std::io::Error> {
    let path = fixture_path(filename);
    let mut file = File::open(&path)?;
    let mut buffer = Vec::new();
    file.read_to_end(&mut buffer)?;
    Ok(buffer)
}

/// TC-P-010: Small File Processing Speed
///
/// 10MB以下のExcelファイルを1秒以内に処理することを目標とする。
fn benchmark_small_file(c: &mut Criterion) {
    // 10MBファイルが存在しない場合はスキップ
    if !fixture_exists("10mb_file.xlsx") {
        eprintln!("Warning: tests/fixtures/bench/10mb_file.xlsx not found. Skipping benchmark.");
        return;
    }

    let data = match load_fixture("10mb_file.xlsx") {
        Ok(d) => d,
        Err(e) => {
            eprintln!(
                "Warning: Failed to load 10mb_file.xlsx: {}. Skipping benchmark.",
                e
            );
            return;
        }
    };

    let converter = ConverterBuilder::new().build().unwrap();

    let mut group = c.benchmark_group("small_file");
    group.throughput(Throughput::Bytes(data.len() as u64));
    group.sample_size(10); // 10回のサンプルで平均を取る

    group.bench_function("convert_10mb_file", |b| {
        b.iter(|| {
            let input = Cursor::new(black_box(&data));
            let mut output = Vec::new();
            converter
                .convert(black_box(input), black_box(&mut output))
                .unwrap();
            black_box(output)
        });
    });

    group.finish();
}

/// TC-P-011: Batch Processing Throughput
///
/// 50ファイル（各10MB）を1分以内に処理することを目標とする。
fn benchmark_batch_processing(c: &mut Criterion) {
    // バッチファイルが存在しない場合はスキップ
    if !fixture_exists("batch_00.xlsx") {
        eprintln!("Warning: tests/fixtures/bench/batch_*.xlsx not found. Skipping benchmark.");
        return;
    }

    // バッチファイルを読み込む（最大50ファイル）
    let mut batch_files = Vec::new();
    for i in 0..50 {
        let filename = format!("batch_{:02}.xlsx", i);
        if let Ok(data) = load_fixture(&filename) {
            batch_files.push(data);
        } else {
            break;
        }
    }

    if batch_files.is_empty() {
        eprintln!("Warning: No batch files found. Skipping benchmark.");
        return;
    }

    let converter = ConverterBuilder::new().build().unwrap();

    let mut group = c.benchmark_group("batch_processing");
    group.sample_size(5); // バッチ処理は時間がかかるため、5回のサンプル

    group.bench_function("convert_50_files", |b| {
        b.iter(|| {
            for file_data in &batch_files {
                let input = Cursor::new(black_box(file_data));
                let mut output = Vec::new();
                converter
                    .convert(black_box(input), black_box(&mut output))
                    .unwrap();
                black_box(output);
            }
        });
    });

    group.finish();
}

/// 中規模ファイル（100MB）のベンチマーク
fn benchmark_medium_file(c: &mut Criterion) {
    if !fixture_exists("100mb_file.xlsx") {
        eprintln!("Warning: tests/fixtures/bench/100mb_file.xlsx not found. Skipping benchmark.");
        return;
    }

    let data = match load_fixture("100mb_file.xlsx") {
        Ok(d) => d,
        Err(e) => {
            eprintln!(
                "Warning: Failed to load 100mb_file.xlsx: {}. Skipping benchmark.",
                e
            );
            return;
        }
    };

    let converter = ConverterBuilder::new().build().unwrap();

    let mut group = c.benchmark_group("medium_file");
    group.throughput(Throughput::Bytes(data.len() as u64));
    group.sample_size(5); // 中規模ファイルは時間がかかるため、5回のサンプル

    group.bench_function("convert_100mb_file", |b| {
        b.iter(|| {
            let input = Cursor::new(black_box(&data));
            let mut output = Vec::new();
            converter
                .convert(black_box(input), black_box(&mut output))
                .unwrap();
            black_box(output)
        });
    });

    group.finish();
}

/// 大規模ファイル（1GB）のベンチマーク
///
/// 注意: このベンチマークは非常に時間がかかるため、通常はスキップされる。
/// 実行する場合は `cargo bench -- --ignored` を使用するか、
/// 環境変数 `BENCH_LARGE_FILE=true` を設定してください。
fn benchmark_large_file(c: &mut Criterion) {
    // 環境変数で有効化されていない場合はスキップ
    if std::env::var("BENCH_LARGE_FILE").is_err() {
        eprintln!("Info: Large file benchmark skipped. Set BENCH_LARGE_FILE=true to enable.");
        return;
    }

    if !fixture_exists("1gb_file.xlsx") {
        eprintln!("Warning: tests/fixtures/bench/1gb_file.xlsx not found. Skipping benchmark.");
        return;
    }

    let data = match load_fixture("1gb_file.xlsx") {
        Ok(d) => d,
        Err(e) => {
            eprintln!(
                "Warning: Failed to load 1gb_file.xlsx: {}. Skipping benchmark.",
                e
            );
            return;
        }
    };

    let converter = ConverterBuilder::new().build().unwrap();

    let mut group = c.benchmark_group("large_file");
    group.throughput(Throughput::Bytes(data.len() as u64));
    group.sample_size(3); // 大規模ファイルは時間がかかるため、3回のサンプル

    group.bench_function("convert_1gb_file", |b| {
        b.iter(|| {
            let input = Cursor::new(black_box(&data));
            let mut output = Vec::new();
            converter
                .convert(black_box(input), black_box(&mut output))
                .unwrap();
            black_box(output)
        });
    });

    group.finish();
}

criterion_group! {
    name = benches;
    config = Criterion::default()
        .measurement_time(std::time::Duration::from_secs(30))
        .warm_up_time(std::time::Duration::from_secs(5));
    targets = benchmark_small_file, benchmark_medium_file, benchmark_batch_processing
}

// 大規模ファイルのベンチマークは別グループとして定義
criterion_group! {
    name = large_benches;
    config = Criterion::default()
        .measurement_time(std::time::Duration::from_secs(300)) // 5分
        .warm_up_time(std::time::Duration::from_secs(10));
    targets = benchmark_large_file
}

criterion_main!(benches, large_benches);
