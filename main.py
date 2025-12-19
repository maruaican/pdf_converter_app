# -*- coding: utf-8 -*-
import sys
import os

# 最初に「準備をしています...」を表示（インポート等で時間がかかる前に行う）
print("準備をしています...")
sys.stdout.flush()

# pythoncomの初期化をメインスレッドで行う
import pythoncom
from converters import ExcelConverter, WordConverter


def main():
    # 標準出力のエンコーディングをUTF-8に設定（Windows環境での文字化け対策）
    if sys.platform == "win32":
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

    # 引数チェック
    if len(sys.argv) < 2:
        print("\nWord、Excelファイルをこの実行ファイルにドラッグアンドドロップしてください。")
        # 実行ファイルとして実行されている場合のみ入力を待つ
        if getattr(sys, 'frozen', False):
            input("\nEnterキーを押して終了してください。")
        return

    success_info = []
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
        dirname = os.path.dirname(file_path)
        
        if not os.path.exists(file_path):
            error_files.append(f"{basename} (ファイルが見つかりません)")
            continue

        print(f"PDFに変換中: {basename}")

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
                # 変換後のファイル名（拡張子をpdfに変更）
                pdf_name = os.path.splitext(basename)[0] + ".pdf"
                success_info.append({
                    'dir': dirname,
                    'file': pdf_name
                })
            else:
                unsupported_files.append(basename)
        except Exception as e:
            error_message = str(e).strip()
            if not error_message:
                error_message = "不明なエラーが発生しました"
            error_files.append(f"{basename} (エラー: {error_message})")

    # 結果メッセージの生成
    result_message_parts = []
    if success_info:
        success_lines = ["--- 成功 ---"]
        for info in success_info:
            success_lines.append(f"場所: {info['dir']}")
            success_lines.append(f"ファイル: {info['file']}")
            success_lines.append("") # 空行
        result_message_parts.append("\n".join(success_lines))
        
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
    
    # 実行ファイルとして実行されている場合、または標準入力が端末の場合のみ入力を待つ
    if getattr(sys, 'frozen', False) or sys.stdin.isatty():
        input("\n処理が完了しました。Enterキーを押して終了してください。")


if __name__ == "__main__":
    main()
