#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
簡易PDFコンバーターテスト
"""

import os
import sys
from pathlib import Path

def test_basic_functionality():
    """基本的な機能テスト"""
    print("=== PDFコンバーター簡易テスト ===\n")
    
    # 現在のディレクトリ確認
    current_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"現在のディレクトリ: {current_dir}")
    
    # モジュールの存在確認
    try:
        from converters import ExcelConverter, WordConverter
        print("✅ コンバーターモジュールのインポート成功")
    except Exception as e:
        print(f"❌ コンバーターモジュールのインポート失敗: {e}")
        return False
    
    # ファイル拡張子チェック
    test_cases = [
        ("test.xlsx", ExcelConverter),
        ("test.xls", ExcelConverter),
        ("test.docx", WordConverter),
        ("test.doc", WordConverter),
        ("test.txt", None),
        ("", None)
    ]
    
    print("\n--- 拡張子判定テスト ---")
    for filename, expected_converter in test_cases:
        _, ext = os.path.splitext(filename)
        ext = ext.lower()
        
        converter = None
        if ext in [".xlsx", ".xls", ".xlsm"]:
            converter = ExcelConverter
        elif ext in [".docx", ".doc"]:
            converter = WordConverter
            
        if converter == expected_converter:
            print(f"✅ {filename}: 正しく判定")
        else:
            print(f"❌ {filename}: 判定エラー (期待: {expected_converter}, 実際: {converter})")
    
    # パス処理テスト
    print("\n--- パス処理テスト ---")
    test_paths = [
        "test.xlsx",
        "./test.xlsx",
        "C:\\test.xlsx",
        "  test.xlsx  ",
        ""
    ]
    
    for path in test_paths:
        normalized = os.path.abspath(path.strip()) if path and path.strip() else None
        print(f"'{path}' -> {normalized}")
    
    print("\n✅ 基本的な機能テスト完了")
    return True

if __name__ == "__main__":
    test_basic_functionality()