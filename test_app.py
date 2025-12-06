#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDFコンバーターの動作テスト用スクリプト
"""

import os
import sys
import tempfile
import subprocess
from pathlib import Path

def create_test_files():
    """テスト用のWordとExcelファイルを作成"""
    try:
        import win32com.client
        
        # テスト用ディレクトリ
        test_dir = Path("test_files")
        test_dir.mkdir(exist_ok=True)
        
        # Wordファイル作成
        word_path = test_dir / "test_document.docx"
        if not word_path.exists():
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Add()
            doc.Content.Text = "これはテスト用のWordドキュメントです。\nPDF変換のテストに使用します。"
            doc.SaveAs(str(word_path))
            doc.Close()
            word.Quit()
            print(f"✓ Wordテストファイル作成: {word_path}")
        
        # Excelファイル作成
        excel_path = test_dir / "test_workbook.xlsx"
        if not excel_path.exists():
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Cells(1, 1).Value = "テストデータ"
            ws.Cells(2, 1).Value = "PDF変換テスト"
            ws.Cells(3, 1).Value = "成功を確認"
            wb.SaveAs(str(excel_path))
            wb.Close()
            excel.Quit()
            print(f"✓ Excelテストファイル作成: {excel_path}")
            
        return str(word_path), str(excel_path)
        
    except Exception as e:
        print(f"テストファイル作成エラー: {e}")
        return None, None

def test_converter():
    """PDFコンバーターの動作テスト"""
    print("=== PDFコンバーター動作テスト ===\n")
    
    # テストファイル作成
    word_file, excel_file = create_test_files()
    
    if not word_file or not excel_file:
        print("❌ テストファイル作成に失敗しました")
        return False
    
    # テスト実行
    test_files = [word_file, excel_file]
    success_count = 0
    
    for test_file in test_files:
        print(f"\n--- {os.path.basename(test_file)} の変換テスト ---")
        
        try:
            # PDF変換実行
            result = subprocess.run([
                sys.executable, "main.py", test_file
            ], capture_output=True, text=True, timeout=30)
            
            # 結果確認
            pdf_file = test_file.replace('.docx', '.pdf').replace('.xlsx', '.pdf')
            
            if os.path.exists(pdf_file):
                file_size = os.path.getsize(pdf_file)
                print(f"✅ 成功 - PDF作成: {pdf_file} ({file_size} bytes)")
                success_count += 1
                
                # テスト後クリーンアップ
                os.remove(pdf_file)
            else:
                print(f"❌ 失敗 - PDFファイルが作成されませんでした")
                if result.stderr:
                    print(f"エラー: {result.stderr}")
                    
        except subprocess.TimeoutExpired:
            print("❌ タイムアウト - 30秒以上かかっています")
        except Exception as e:
            print(f"❌ エラー: {e}")
    
    print(f"\n=== テスト結果 ===")
    print(f"成功: {success_count}/{len(test_files)} ファイル")
    
    if success_count == len(test_files):
        print("🎉 すべてのテストに成功しました！")
        return True
    else:
        print("⚠️ 一部のテストに失敗しました")
        return False

def test_edge_cases():
    """エッジケースのテスト"""
    print("\n=== エッジケーステスト ===\n")
    
    test_cases = [
        ("存在しないファイル", "nonexistent.docx"),
        ("空のファイルパス", ""),
        ("サポート外の拡張子", "test.txt"),
    ]
    
    for case_name, test_file in test_cases:
        print(f"--- {case_name} ---")
        
        try:
            result = subprocess.run([
                sys.executable, "main.py", test_file
            ], capture_output=True, text=True, timeout=10)
            
            print(f"終了コード: {result.returncode}")
            if result.stdout:
                print(f"出力: {result.stdout.strip()}")
            if result.stderr:
                print(f"エラー: {result.stderr.strip()}")
                
        except Exception as e:
            print(f"例外: {e}")
        
        print()

if __name__ == "__main__":
    # 現在のディレクトリをPDFコンバーターに変更
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    # 基本テスト
    basic_success = test_converter()
    
    # エッジケーステスト
    test_edge_cases()
    
    # 結果サマリー
    print("=== テスト完了 ===")
    if basic_success:
        print("✅ 基本機能は正常に動作しています")
    else:
        print("❌ 基本機能に問題があります")