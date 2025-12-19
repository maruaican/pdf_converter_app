import os
import win32com.client
import pythoncom
from .base_converter import BaseConverter

class WordConverter(BaseConverter):
    def convert(self):
        word = None
        doc = None
        pythoncom.CoInitialize()
        try:
            # DispatchEx を使用して新しいインスタンスを強制する
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0 # wdAlertsNone
            
            # 絶対パスを使用
            abs_path = os.path.abspath(self.file_path)
            
            # PDFのパスを作成
            base, ext = os.path.splitext(abs_path)
            pdf_path = base + ".pdf"
            
            # 既存のPDFがあれば削除を試みる
            if os.path.exists(pdf_path):
                try:
                    os.remove(pdf_path)
                except:
                    pass

            # ReadOnly=True で開き、変更確認などを抑制
            doc = word.Documents.Open(abs_path, ReadOnly=True, ConfirmConversions=False)
            
            # PDFとして保存 (17=wdExportFormatPDF)
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0, # 0=wdExportOptimizeForPrint
                Range=0, # 0=wdExportAllDocument
                From=1,
                To=1,
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
                try:
                    doc.Close(SaveChanges=0) # 0=wdDoNotSaveChanges
                except:
                    pass
            if word:
                try:
                    word.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
