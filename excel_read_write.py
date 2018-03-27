#-*- coding: utf-8 -*- 
import xlrd, xlwt  ## 匯入讀取寫入 xls檔案必須的模組 xlrd跟xlwt 要使用pip另外安裝

open_xls = xlrd.open_workbook("python_test.xls")  # 打開xls文件

sheet_name = open_xls.sheet_names()    # 取用各工作表名稱

sheet_num  = open_xls.nsheets          # 取得檔案中的工作表總數量

sheet_0 = open_xls.sheet_by_index(0)   # 將sheet_0 設定為第一個工作表 index為0 
 
print(sheet_name)
print("工作表總數量 : " + str(sheet_num))
print(sheet_0.nrows, sheet_0.ncols)

# 取得第一個工作表中特定區域的儲存格資料
# index都要從0開始，所以cols欄是第0欄 而rows是第1列開始

for y in range(1, sheet_0.nrows,1):
    print("儲存格資料:"+ str(sheet_0.cell_value(y,0))) 
    # 因為輸出，會直接將資料型態一起，所以建議轉型為str 再輸出