import os
import win32com.client
from .base_converter import BaseConverter

class WordConverter(BaseConverter):
    def convert(self):
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            word.ScreenUpdating = False
            
            # ReadOnly=True で開き、変更確認などを抑制
            doc = word.Documents.Open(self.file_path, ReadOnly=True)
            
            # PDFのパスを作成
            pdf_path = os.path.splitext(self.file_path) + ".pdf"
            
            # PDFとして保存 (17=wdExportFormatPDF)
            # OpenAfterExport=False: 変換後にPDFを開かない
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0, # 0=wdExportOptimizeForPrint
                Item=0, # 0=wdExportDocumentContent
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=0, # 0=wdExportCreateNoBookmarks
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )
            
        except Exception as e:
            raise e
        finally:
            if doc:
                # SaveChanges=0 (wdDoNotSaveChanges)
                doc.Close(SaveChanges=0)
            if word:
                word.ScreenUpdating = True
                word.Quit()