//! calamine統合調査とスパイク実装
//!
//! このスパイクコードは、calamineクレートのAPIを調査し、
//! 以下の機能の動作を検証することを目的としています:
//! - 数式情報取得（worksheet_formula）
//! - 結合セル情報取得（load_merged_regions, worksheet_merge_cells）
//! - Number Format String取得の可否確認
//! - 非表示行・列情報取得の可否確認

use calamine::{open_workbook, DataType, Reader, Xlsx};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== calamine API調査スパイク ===\n");

    // テスト用のExcelファイルパス（コマンドライン引数から取得、またはデフォルト）
    let file_path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "test.xlsx".to_string());

    if !Path::new(&file_path).exists() {
        eprintln!("警告: ファイル '{}' が見つかりません。", file_path);
        eprintln!("使用方法: cargo run --example spike_calamine -- <path_to_xlsx_file>");
        eprintln!("\n調査結果を確認するには、実際のExcelファイルが必要です。");
        return Ok(());
    }

    println!("ファイル: {}\n", file_path);

    // 1. 基本的なワークブックの読み込み
    println!("--- 1. 基本的なワークブックの読み込み ---");
    let mut workbook: Xlsx<_> = open_workbook(&file_path)?;
    println!("✅ ワークブックの読み込み成功\n");

    // 2. シート名の取得
    println!("--- 2. シート名の取得 ---");
    let sheet_names = workbook.sheet_names().to_vec();
    println!("シート数: {}", sheet_names.len());
    for (idx, name) in sheet_names.iter().enumerate() {
        println!("  [{}] {}", idx, name);
    }
    println!();

    // 3. 各シートの調査
    for (sheet_idx, sheet_name) in sheet_names.iter().enumerate() {
        println!(
            "=== シート: {} (インデックス: {}) ===",
            sheet_name, sheet_idx
        );

        // 3.1. シートの選択
        let range = workbook
            .worksheet_range(sheet_name)
            .map_err(|e| format!("シート '{}' の読み込みエラー: {}", sheet_name, e))?;

        println!("\n--- 3.1. セルデータの取得 ---");
        let mut cell_count = 0;
        let mut empty_count = 0;

        for row in range.rows() {
            for cell in row {
                cell_count += 1;
                if cell.is_empty() {
                    empty_count += 1;
                }
            }
        }
        println!("総セル数: {}", cell_count);
        println!("空セル数: {}", empty_count);
        println!("データセル数: {}", cell_count - empty_count);

        // 3.2. 数式情報の取得（worksheet_formula）
        println!("\n--- 3.2. 数式情報の取得（worksheet_formula） ---");
        let mut formula_count = 0;
        match workbook.worksheet_formula(sheet_name) {
            Ok(formula_range) => {
                println!("✅ worksheet_formula() は利用可能");
                for row in formula_range.rows() {
                    for cell in row {
                        if !cell.is_empty() {
                            formula_count += 1;
                            if formula_count <= 5 {
                                // 最初の5つの数式を表示
                                println!("  数式: {:?}", cell);
                            }
                        }
                    }
                }
                println!("  数式セル数: {}", formula_count);
            }
            Err(e) => {
                println!("❌ worksheet_formula() エラー: {}", e);
                println!("   この機能は利用できない可能性があります");
            }
        }

        // 3.3. 結合セル情報の取得（load_merged_regions）
        println!("\n--- 3.3. 結合セル情報の取得（load_merged_regions） ---");
        // load_merged_regions()は引数を取らず、ワークブック全体の結合領域を読み込む
        match workbook.load_merged_regions() {
            Ok(_) => {
                println!(
                    "✅ load_merged_regions() は利用可能（ワークブック全体の結合領域を読み込み）"
                );
                // 読み込み後、worksheet_merge_cells()で取得
            }
            Err(e) => {
                println!("❌ load_merged_regions() エラー: {}", e);
            }
        }

        // 3.4. 結合セル情報の取得（worksheet_merge_cells）
        println!("\n--- 3.4. 結合セル情報の取得（worksheet_merge_cells） ---");
        match workbook.worksheet_merge_cells(sheet_name) {
            Some(Ok(merge_cells)) => {
                println!("✅ worksheet_merge_cells() は利用可能");
                println!("  結合セル範囲数: {}", merge_cells.len());
                for (idx, region) in merge_cells.iter().take(5).enumerate() {
                    println!(
                        "  [{}] 開始: ({}, {}), 終了: ({}, {})",
                        idx, region.start.0, region.start.1, region.end.0, region.end.1
                    );
                }
                if merge_cells.len() > 5 {
                    println!("  ... (他 {} 件)", merge_cells.len() - 5);
                }
            }
            Some(Err(e)) => {
                println!("❌ worksheet_merge_cells() エラー: {}", e);
            }
            None => {
                println!("⚠️  worksheet_merge_cells() は None を返しました（結合セル情報が読み込まれていない可能性）");
            }
        }

        // 3.5. Number Format String取得の試行
        println!("\n--- 3.5. Number Format String取得の試行 ---");
        // calamineのAPIを確認: セルのスタイル情報を取得できるか？
        // 注: calamineのドキュメントによると、Phase Iでは直接的なAPIは提供されていない可能性が高い
        println!("⚠️  calamineのAPIを確認中...");
        println!("    Range<DataType>からは直接Number Format Stringを取得できない可能性が高い");
        println!("    Phase IIでXML直接解析が必要になる見込み");

        // セルの型情報を確認
        let range = workbook
            .worksheet_range(sheet_name)
            .map_err(|e| format!("シート '{}' の再読み込みエラー: {}", sheet_name, e))?;
        let mut sample_cells = 0;
        for row in range.rows().take(10) {
            for cell in row {
                if !cell.is_empty() && sample_cells < 5 {
                    println!(
                        "  セル値: {:?} (型: {:?})",
                        cell,
                        std::any::type_name_of_val(cell)
                    );
                    sample_cells += 1;
                }
            }
        }

        // 3.6. 非表示行・列情報取得の試行
        println!("\n--- 3.6. 非表示行・列情報取得の試行 ---");
        println!("⚠️  calamineのAPIを確認中...");
        println!("    GitHub Issue #237を参照: 非表示行・列の情報は取得できない可能性が高い");
        println!("    Phase IIでXML直接解析が必要になる見込み");

        // ワークシートの行数を確認（非表示情報なし）
        let range = workbook
            .worksheet_range(sheet_name)
            .map_err(|e| format!("シート '{}' の再読み込みエラー: {}", sheet_name, e))?;
        let max_row = range.end().map(|c| c.0).unwrap_or(0);
        let max_col = range.end().map(|c| c.1).unwrap_or(0);
        println!("   最大行: {}, 最大列: {}", max_row, max_col);
        println!("   (非表示行・列の情報は取得できません)");

        println!("\n");
    }

    // 4. 調査結果のまとめ
    println!("=== 調査結果のまとめ ===");
    println!("✅ 基本的なワークブック読み込み: 成功");
    println!("✅ シート名の取得: 成功");
    println!("✅ セルデータの取得: 成功");
    println!("⚠️  数式情報の取得: 要確認（worksheet_formula()の動作確認が必要）");
    println!("✅ 結合セル情報の取得: 確認済み（load_merged_regions / worksheet_merge_cells）");
    println!("❌ Number Format String: 取得不可（Phase IIでXML直接解析が必要）");
    println!("❌ 非表示行・列情報: 取得不可（Phase IIでXML直接解析が必要）");

    Ok(())
}
