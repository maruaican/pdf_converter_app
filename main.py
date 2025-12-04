# -*- coding: utf-8 -*-
import sys
import os
import tkinter as tk
from tkinter import messagebox

from converters import ExcelConverter, WordConverter


def main():
    # 引数チェック
    if len(sys.argv) < 2:
        # 引数がない場合は使用方法を表示
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(
            "PDF Converter", "Word、Excelファイルを\nこのアイコンにドラッグアンドドロップしてください。"
        )
        root.destroy()
        return

    success_files = []
    error_files = []
    unsupported_files = []

    # ファイルごとに処理
    for file_path in sys.argv[1:]:
        basename = os.path.basename(file_path)
        if not os.path.exists(file_path):
            error_files.append(f"{basename} (ファイルが見つかりません)")
            continue

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
            error_files.append(f"{basename} (エラー: {e})")

    # 結果メッセージの生成
    result_message_parts = []
    if success_files:
        result_message_parts.append("--- 成功 ---\n" + "\n".join(success_files))
    if error_files:
        result_message_parts.append("--- 失敗 ---\n" + "\n".join(error_files))
    if unsupported_files:
        result_message_parts.append("--- 未対応 ---\n" + "\n".join(unsupported_files))

    # 何かしらの結果がある場合のみメッセージボックスを表示
    if result_message_parts:
        final_message = "\n\n".join(result_message_parts)
        final_message += "\n\n(このウィンドウは5秒後に閉じます)"

        root = tk.Tk()
        root.withdraw()

        # 5秒で消えるカスタムメッセージボックス
        dialog = tk.Toplevel(root)
        dialog.title("変換結果")
        dialog.attributes("-topmost", True)
        tk.Label(dialog, text=final_message, padx=20, pady=20).pack()

        # 5000ミリ秒（5秒）後にウィンドウを閉じる
        dialog.after(5000, dialog.destroy)

        dialog.mainloop()


if __name__ == "__main__":
    main()
