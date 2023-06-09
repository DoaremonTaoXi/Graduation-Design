import xlwings as xw  # 导入xlwings模块

import myFct

# 宏定义
SHEETNUMBER = 23  # 工作表序号

with xw.App(visible = False, add_book = False) as app:  # 启动Excel程序
    workbook = app.books.open('(公开)S3低频电缆网20190920.xls')
    worksheets = workbook.sheets  # 获取工作簿中所有的工作表
    worksheet = worksheets[SHEETNUMBER]
    print(worksheet.name)


    # 节点序号检查
    ErrorNumber_flag = 0; ErrorNumber_cell = []; ErrorNumber_value = []
    ErrorNumber_flag, ErrorNumber_cell, ErrorNumber_value = myFct.CheckNumber(worksheet)
    
    if ErrorNumber_flag:
            for i in range(len(ErrorNumber_cell)):
                  print(f"Error Cell: {ErrorNumber_cell[i]}")
                  print(f"Error Value: {ErrorNumber_value[i]}")