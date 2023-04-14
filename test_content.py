import xlwings as xw  # 导入xlwings模块
import myFct

# 宏定义
SHEETNUMBER = 1  # 工作表序号

with xw.App(visible = False, add_book = False) as app:  # 启动Excel程序
    
    workbook = app.books.open('(公开)低频电缆网20190920（简化版）.xls')
    worksheets = workbook.sheets  # 获取工作簿中所有的工作表
    worksheet = worksheets[SHEETNUMBER]
    print(worksheet.name)

    # 节点内容检查
    ErrorContent_flag = 0; ErrorContent_cell = []; ErrorContent_value = []
    ErrorContent_flag, ErrorContent_cell, ErrorContent_value = myFct.CheckContent(worksheet)

    if ErrorContent_flag:
        for i in range(len(ErrorContent_cell)):
                print(f"Error Cell: {ErrorContent_cell[i]}")
                print(f"Error Cell: {ErrorContent_value[i]}")