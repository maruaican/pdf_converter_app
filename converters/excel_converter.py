
import os
import win32com.client
from .base_converter import BaseConverter

class ExcelConverter(BaseConverter):
    def convert(self):
        excel = None
        wb = None
        try:
            # COMオブジェクト初期化
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.Interactive = False
            excel.ScreenUpdating = False

            # ファイルパスを絶対パスに変換
            abs_file_path = os.path.abspath(self.file_path)
            
            # ワークブックを開く
            wb = excel.Workbooks.Open(
                abs_file_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True
            )

            # PDF出力パスを準備
            base, ext = os.path.splitext(abs_file_path)
            pdf_path = base + ".pdf"
            
            # 出力ディレクトリが存在するか確認
            output_dir = os.path.dirname(pdf_path)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # PDFエクスポート
            wb.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=pdf_path,
                OpenAfterPublish=False,
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False
            )

            return pdf_path

        except Exception as e:
            raise RuntimeError(f"Excel to PDF変換に失敗しました: {str(e)}")
        finally:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.ScreenUpdating = True
                excel.Interactive = True
                excel.Quit()

    @staticmethod
    def is_available():
        try:
            win32com.client.Dispatch("Excel.Application")
            return True
        except Exception:
            return False
