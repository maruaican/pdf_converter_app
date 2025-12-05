
import os
import win32com.client
from .base_converter import BaseConverter

class ExcelConverter(BaseConverter):
    def convert(self):
        excel = None
        wb = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.Interactive = False
            excel.ScreenUpdating = False
            
            # UpdateLinks=0 でリンク更新を無効化
            # ReadOnly=True で読み取り専用として開き、変更確認ダイアログなどを抑制
            wb = excel.Workbooks.Open(self.file_path, UpdateLinks=0, ReadOnly=True)
            
            # PDFのパスを作成 (os.path.splitextはタプルを返すため、でファイル名部分を取得)
            pdf_path = os.path.splitext(self.file_path) + ".pdf"
            
            # PDFとして保存 (0=xlTypePDF)
            # OpenAfterPublish=False: 変換後にPDFを開かない
            wb.ExportAsFixedFormat(Type=0, Filename=pdf_path, OpenAfterPublish=False)
            
        except Exception as e:
            raise e
        finally:
            if wb:
                wb.Close(SaveChanges=0)
            if excel:
                # 設定を元に戻してから終了（プロセスが残るのを防ぐため念のため）
                excel.ScreenUpdating = True
                excel.Interactive = True
                excel.Quit()
