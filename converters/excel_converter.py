import os
import win32com.client
from .base_converter import BaseConverter

class ExcelConverter(BaseConverter):
    def convert(self):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        try:
            # UpdateLinks=0 でリンク更新を無効化し、シート間参照などによるダイアログやエラーを回避
            wb = excel.Workbooks.Open(self.file_path, UpdateLinks=0)
            
            # PDFのパスを作成
            pdf_path = os.path.splitext(self.file_path)[0] + ".pdf"
            
            # PDFとして保存 (0=xlTypePDF)
            wb.ExportAsFixedFormat(0, pdf_path)
            
            wb.Close()
        except Exception as e:
            raise e
        finally:
            excel.Quit()