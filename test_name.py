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
    ErrorName_flag = 0; ErrorName_index = []; ErrorName_value = []
    ErrorName_flag, ErrorName_index, ErrorName_value = myFct.CheckName(worksheet)
    
    if ErrorName_flag:
            for i in range(len(ErrorName_index)):
                  print(f"Error Cell Index:{ErrorName_index[i]}")
                  print(f"Error Cell Value:{ErrorName_value[i]}")