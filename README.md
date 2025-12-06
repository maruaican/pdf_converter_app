# PDFコンコンバーターアプリ

Word、ExcelファイルをPDFに変換するCLIアプリケーションです。

## 機能

- **Wordファイル** (.docx, .doc) → PDF変換
- **Excelファイル** (.xlsx, .xls, .xlsm) → PDF変換
- **ドラッグ&ドロップ対応**: ファイルを実行ファイルにドラッグ&ドロップするだけ
- **エラーーハンドリング**: 詳細なエラーーメッセージ表示
- **リソース管理**: Officeプロセスの適切なクリーンアップ

## 使用方法

### 1. 実行方法

#### 方法1: ドラッグ&ドロップ
WordまたはExcelファイルを `main.exe` にドラッグ&ドロップしてください。

#### 方法2: コマンドライン
```bash
python main.py ファイル1.docx ファイル2.xlsx
```

### 2. 出力
変換されたPDFファイルは元のファイルと同じディレクトリに作成されます。

## 動作環境

- **OS**: Windows
- **Python**: 3.7以上
- **必要なライブラリ**:
  - pywin32
  - comtypes

## インストール

```bash
pip install -r requirements.txt
```

## トラブルシューーティング

### よくある問題

1. **「ファイルが見つかりません」エラー**
   - ファイルパスに特殊文字が含まれていないか確認
   - ファイルが存在することを確認

2. **Officeが開かない**
   - Microsoft Officeがインストールされていることを確認
   - Officeのライセンスが有効であることを確認

3. **プロセスが残る**
   - アプリケーションは自動的にOfficeプロセスを終了します
   - 異常終了時はタスクマネージャーで手動終了

## テスト

### 簡易テスト
```bash
python simple_test.py
```

### 詳細テスト
```bash
python test_app.py
```

## ファイル構成

```
pdf_converter_app/
├── main.py              # メイン実行ファイル
├── converters/          # コンコンバーターーモジュール
│   ├── __init__.py
│   ├── base_converter.py
│   ├── excel_converter.py
│   └── word_converter.py
├── test_app.py          # 詳細テストスクリプト
├── simple_test.py       # 簡易テストスクリプト
├── requirements.txt     # 依存関係
└── README.md           # このファイル
```

## 更新履歴

- **v1.1**: エラーーハンドリング強化、パス処理改善
- **v1.0**: 初期リリース