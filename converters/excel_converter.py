import os
import win32com.client
import pythoncom
from .base_converter import BaseConverter

class ExcelConverter(BaseConverter):
    def convert(self):
        excel = None
        wb = None
        pythoncom.CoInitialize()
        try:
            # DispatchEx を使用
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.Interactive = False

            # ファイルパスを絶対パスに変換
            abs_file_path = os.path.abspath(self.file_path)
            
            # PDF出力パスを準備
            base, ext = os.path.splitext(abs_file_path)
            pdf_path = base + ".pdf"
            
            # 既存のPDFがあれば削除を試みる
            if os.path.exists(pdf_path):
                try:
                    os.remove(pdf_path)
                except:
                    pass

            # ワークブックを開く
            wb = excel.Workbooks.Open(
                abs_file_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True
            )

            # 全シートをPDF化するために、ワークブック全体をエクスポート対象にする
            # ExportAsFixedFormat の引数で Type=0 (xlTypePDF) を指定
            # ワークブックオブジェクトに対して呼び出すことで全シートが対象となる
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
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    @staticmethod
    def is_available():
        pythoncom.CoInitialize()
        try:
            win32com.client.Dispatch("Excel.Application")
            return True
        except Exception:
            return False
        finally:
            pythoncom.CoUninitialize()
