#-*- coding: utf-8 -*- 
import xlrd, xlwt  ## 匯入讀取寫入 xls檔案必須的模組 xlrd跟xlwt 要使用pip另外安裝

open_xls = xlrd.open_workbook("python_test.xls")  # 打開xls文件

sheet_name = open_xls.sheet_names()    # 取用各工作表名稱

sheet_num  = open_xls.nsheets          # 取得檔案中的工作表總數量

sheet_0 = open_xls.sheet_by_index(0)   # 將sheet_0 設定為第一個工作表 index為0 
 
print(sheet_0.nrows)

print(sheet_name, sheet_num)