import win32com.client
import os

def create_excel_test_file(filename):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Add()
    
    # シート1: データあり
    ws1 = wb.Sheets(1)
    ws1.Name = "DataSheet"
    ws1.Cells(1, 1).Value = "Test Data"
    
    # シート2: 白紙
    ws2 = wb.Sheets.Add(After=ws1)
    ws2.Name = "BlankSheet"
    
    # シート3: データあり
    ws3 = wb.Sheets.Add(After=ws2)
    ws3.Name = "DataSheet2"
    ws3.Cells(2, 2).Value = "More Data"

    wb.SaveAs(os.path.abspath(filename))
    wb.Close()
    excel.Quit()
    print(f"Created: {filename}")

def create_word_test_file(filename):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Add()
    
    # ページ1: テキストあり
    doc.Content.Text = "Page 1 Content"
    doc.Content.InsertBreak(7) # wdPageBreak
    
    # ページ2: 白紙（改ページのみ）
    doc.Content.InsertBreak(7)
    
    # ページ3: テキストあり
    selection = word.Selection
    selection.EndKey(6) # wdStory
    selection.TypeText("Page 3 Content")

    doc.SaveAs(os.path.abspath(filename))
    doc.Close()
    word.Quit()
    print(f"Created: {filename}")

def create_ppt_test_file(filename):
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    # PowerPointはウィンドウ表示が必要な場合が多いが、作成時は非表示でもいけるか試す
    # エラーになる場合は WithWindow=True にする
    try:
        pres = ppt.Presentations.Add(WithWindow=False)
    except:
        pres = ppt.Presentations.Add(WithWindow=True)

    # スライド1: タイトルあり
    slide1 = pres.Slides.Add(1, 1) # ppLayoutTitle
    slide1.Shapes.Title.TextFrame.TextRange.Text = "Slide 1 Title"
    
    # スライド2: 白紙
    pres.Slides.Add(2, 12) # ppLayoutBlank
    
    # スライド3: コンテンツあり
    slide3 = pres.Slides.Add(3, 2) # ppLayoutText
    slide3.Shapes(1).TextFrame.TextRange.Text = "Slide 3 Title"
    slide3.Shapes(2).TextFrame.TextRange.Text = "Content"

    pres.SaveAs(os.path.abspath(filename))
    pres.Close()
    ppt.Quit()
    print(f"Created: {filename}")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    create_excel_test_file(os.path.join(base_dir, "test_excel.xlsx"))
    create_word_test_file(os.path.join(base_dir, "test_word.docx"))
    create_ppt_test_file(os.path.join(base_dir, "test_ppt.pptx"))