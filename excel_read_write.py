#-*- coding: utf-8 -*- 
import xlrd, xlwt, xlutils  ## 匯入讀取寫入 xls檔案必須的模組 xlrd跟xlwt 要使用pip另外安裝
## 要修改已存在的xls檔案，要使用xlutils
from xlutils.copy import copy # 一定要匯入，才知道是匯入哪一個copy

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

##### 修改原本Excel的檔案內容，將修改後內容，寫到第二個工作表 ####

write_xls = copy(open_xls) ## 使用xlutils中的copy() 來達到修改xls檔案內容的功能

write_content = write_xls.get_sheet(1) ## 使用get_sheet(index值) 來取得第二張工作表

for i in range (2, sheet_0.nrows,1):   ##  內容的數字是從index = 2 開始
    data = sheet_0.cell_value(i,0)
    data1000 = (data * 10 + 4) /4
    formula = "=(" + str(data) + "*10+4)/4"
    if data >= 3.0 :
        write_content.write(i,0, data1000)        ## 前面第一個數字表示第幾列，第二個數字是欄，第三個為內容
        write_content.write(i,1, formula)         ## 可以多重運算後，再寫入另外一欄的資料列中
    else: 
        write_content.write(i,0, "資料過小")
        write_content.write(i,1, formula)
    write_xls.save("python_test.xls")             ## 最後，存檔離開，亦可選擇其他路徑，另存新檔