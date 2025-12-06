# -*- coding: utf-8 -*-
import sys
import os

from converters import ExcelConverter, WordConverter


def main():
    print("準備をしています...")

    # 引数チェック
    if len(sys.argv) < 2:
        print("\nWord、Excelファイルをこの実行ファイルにドラッグアンドドロップしてください。")
        input("\nEnterキーを押して終了してください。")
        return

    success_files = []
    error_files = []
    unsupported_files = []

    # ファイルごとに処理
    for file_path in sys.argv[1:]:
        # ファイルパスが空でないことを確認
        if not file_path or not file_path.strip():
            continue
            
        # パスを正規化
        file_path = os.path.abspath(file_path.strip())
        basename = os.path.basename(file_path)
        
        if not os.path.exists(file_path):
            error_files.append(f"{basename} (ファイルが見つかりません)")
            continue

        print(f"Converting to PDF: {basename}")

        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        converter = None

        try:
            if ext in [".xlsx", ".xls", ".xlsm"]:
                converter = ExcelConverter(file_path)
            elif ext in [".docx", ".doc"]:
                converter = WordConverter(file_path)

            if converter:
                converter.convert()
                success_files.append(basename)
            else:
                unsupported_files.append(basename)
        except Exception as e:
            error_message = str(e).strip()
            if not error_message:
                error_message = "不明なエラーが発生しました"
            error_files.append(f"{basename} (エラー: {error_message})")

    # 結果メッセージの生成
    result_message_parts = []
    if success_files:
        result_message_parts.append("--- 成功 ---\n" + "\n".join(success_files))
    if error_files:
        result_message_parts.append("--- 失敗 ---\n" + "\n".join(error_files))
    if unsupported_files:
        result_message_parts.append("--- 未対応 ---\n" + "\n".join(unsupported_files))

    # 何かしらの結果がある場合のみメッセージを標準出力
    if result_message_parts:
        final_message = "\n\n".join(result_message_parts)
        print(final_message)
    else:
        print("\n処理するファイルがありませんでした。")
    
    input("\n処理が完了しました。Enterキーを押して終了してください。")


if __name__ == "__main__":
    main()
