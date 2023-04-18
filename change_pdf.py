import win32com.client


excel = win32com.client.Dispatch("Excel.Application")

# Excelファイルの　フルパス
# file = excel.Workbooks.Open(r"C:\Users\natsume\Documents\python_ocr_excel\sample.xlsx")
file = excel.Workbooks.Open('C:/Users/natsume/Documents/python_ocr_excel/sample.xlsx')

file.WorkSheets("Sheet1").Select()

# 保存先の絶対パス
# file.ActiveSheet.ExportAsFixedFormat(0,r"C:\Users\natsume\Documents\python_ocr_excel\go_pdf\change.pdf")
file.ActiveSheet.ExportAsFixedFormat(0,'C:/Users/natsume/Documents/python_ocr_excel/go_pdf/change.pdf')

#エクセルを閉じる
file.Close()
excel.Quit()