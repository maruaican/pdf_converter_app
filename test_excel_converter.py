import unittest
import os
import tempfile
from converters.excel_converter import ExcelConverter

class TestExcelConverter(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # テスト用一時ディレクトリ作成
        cls.test_dir = tempfile.mkdtemp()
        cls.excel_file = os.path.join(cls.test_dir, "test.xlsx")
        
        # テスト用Excelファイル作成
        import openpyxl
        wb = openpyxl.Workbook()
        wb.save(cls.excel_file)

    def test_conversion_success(self):
        # 正常系テスト
        converter = ExcelConverter(self.excel_file)
        pdf_path = converter.convert()
        
        self.assertTrue(os.path.exists(pdf_path))
        self.assertEqual(os.path.splitext(pdf_path)[1], ".pdf")

    def test_file_not_found(self):
        # ファイルが存在しない場合
        with self.assertRaises(RuntimeError):
            converter = ExcelConverter("nonexistent.xlsx")
            converter.convert()

    def test_is_available(self):
        # 利用可否チェック
        self.assertTrue(ExcelConverter.is_available())

    def test_output_dir_creation(self):
        # 出力ディレクトリが作成されるかテスト
        new_dir = os.path.join(self.test_dir, "newdir")
        test_file = os.path.join(new_dir, "test.xlsx")
        
        # ディレクトリが存在しない状態で試す
        converter = ExcelConverter(test_file)
        with self.assertRaises(RuntimeError):
            converter.convert()

    @classmethod
    def tearDownClass(cls):
        # 一時ファイル削除
        for root, dirs, files in os.walk(cls.test_dir, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(cls.test_dir)

if __name__ == "__main__":
    unittest.main()