import openpyxl
 
#ファイルのパスを指定
file_path = r"D:\prg_workspace\RPA_DDCM_send_convertFiles\sample_list.xlsx"
 
#ファイルを開く
excelBook = openpyxl.load_workbook(file_path)

#Excelファイルを保存
wb.Save()
#Excelを閉じる
excel.Quit()