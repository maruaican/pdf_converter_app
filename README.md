# PDFConverter
Word、ExcelファイルをPDFに変換するWindows用アプリケーションです。
実行ファイル（EXE）にファイル（複数ファイル対応可）をドラッグ＆ドロップするだけで、簡単にPDF変換が可能です。

## 主な機能
- **Word変換**: `.docx`, `.doc` ファイルをPDFに変換します。
- **Excel変換**: `.xlsx`, `.xls`, `.xlsm` ファイルをPDFに変換します。
  - **全シート対応**: ワークブックに含まれるすべてのシートを1つのPDFにまとめて出力します。
- **一括処理**: 複数のファイルをまとめてドラッグ＆ドロップして一括変換できます。
- **詳細レポート**: 変換成功時に、保存先のフォルダパスとファイル名を分かりやすく表示します。

## 使用方法
1. `PDFConverter.exe` をデスクトップなどの使いやすい場所に配置します。
2. 変換したいファイルを `PDFConverter.exe` にドラッグ＆ドロップします。
3. 変換が完了すると、元のファイルと同じフォルダにPDFが作成されます。


## 動作環境
- **OS**: Windows 10 / 11
- **ソフトウェア**: Microsoft Office (Word, Excel) がインストールされている必要があります。
- **Python**: 3.10以上（開発・ビルド時）

## 開発者向け情報

### 依存ライブラリのインストール
```bash
pip install -r requirements.txt
```
### 実行ファイルのビルド方法
PyInstallerを使用してビルドします。
```bash
cd pdf_converter_app
pyinstaller --onefile --name PDFConverter main.py
```

## ファイル構成
```
pdf_converter_app/
├── main.py              # メインエントリーポイント
├── converters/          # 変換ロジック
│   ├── base_converter.py
│   ├── excel_converter.py
│   └── word_converter.py
├── dist/                # ビルド済み実行ファイル格納先
├── build/               # ビルド用一時フォルダ
├── test_files/          # テスト用サンプルファイル
├── requirements.txt     # 依存ライブラリ一覧
└── README.md            # 本ファイル
```

## 更新履歴

- **v1.2**: 
  - Excelの全シートPDF化に対応
  - 変換成功時の保存先パス・ファイル名表示機能を追加
- **v1.1**: 
  - COM初期化（pythoncom）の安定化
  - `DispatchEx` による独立プロセス実行の実装
  - 文字化け対策（UTF-8出力）
- **v1.0**: 初期リリース
