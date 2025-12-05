import os
import win32com.client
from .base_converter import BaseConverter

class WordConverter(BaseConverter):
    def convert(self):
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        doc = None
        try:
            doc = word.Documents.Open(self.file_path)
            
            # PDFのパスを作成
            pdf_path = os.path.splitext(self.file_path)[0] + ".pdf"
            
            # PDFとして保存 (17=wdFormatPDF)
            doc.SaveAs2(pdf_path, FileFormat=17)
            
        except Exception as e:
            raise e
        finally:
            if doc:
                doc.Close()
            word.Quit()