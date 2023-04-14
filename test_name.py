import xlwings as xw  # 导入xlwings模块

import myFct

# 宏定义
SHEETNUMBER = 0  # 工作表序号

with xw.App(visible = False, add_book = False) as app:  # 启动Excel程序
    workbook = app.books.open('(公开)低频电缆网20190920（简化版）.xls')
    worksheets = workbook.sheets  # 获取工作簿中所有的工作表
    worksheet = worksheets[SHEETNUMBER]
    print(worksheet.name)

    # 节点名称检查
    ErrorName_flag = 0; ErrorName_cell = []; ErrorName_value = []
    ErrorName_flag, ErrorName_cell, ErrorName_value = myFct.CheckName(worksheet)
    
    if ErrorName_flag:
            for i in range(len(ErrorName_cell)):
                  print(f"Error Cell: {ErrorName_cell[i]}")
                  print(f"Error Value: {ErrorName_value[i]}")