# import openpyxl
 
# #ファイルのパスを指定
# file_path = r"D:\prg_workspace\RPA_DDCM_send_convertFiles\sample.xlsx"
 
# #ファイルを開く
# excelBook = openpyxl.load_workbook(file_path)

# #ファイルに名前を付けて保存
# excelBook.save(r"https://lixilgroup.sharepoint.com/sites/JPFS0144/02_Users/16_加藤恭子/実験/sample.xlsx")

import win32com.client
from pathlib import Path

# 起動する
app = win32com.client.Dispatch("Excel.Application")

# 開く
# abspath = str(Path(r"data/sample.xlsx").resolve())
wb  = app.Workbooks.Open(r"D:\prg_workspace\RPA_DDCM_send_convertFiles\sample.xlsx")

# 保存する
wb.SaveAs(r"https://lixilgroup.sharepoint.com/sites/JPFS0144/02_Users/16_加藤恭子/実験/sample_list.xlsx")

# 閉じる
wb .Close()
# # 終了する
# # app.Quit()