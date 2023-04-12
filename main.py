import xlwings as xw  # 导入xlwings模块

import myFct

with xw.App(visible = False, add_book = False) as app:  # 启动Excel程序
    
    workbook = app.books.open('(公开)低频电缆网20190920（简化版）.xls')
    worksheets = workbook.sheets  # 获取工作簿中所有的工作表


    for worksheet in worksheets:

        print(worksheet.name)
        

        # 节点名称检查
        ErrorName_flag = 0; ErrorName_index = []; ErrorName_value = []
        ErrorName_flag, ErrorName_index, ErrorName_value = myFct.CheckName(worksheet)

        if ErrorName_flag:
            for i in range(len(ErrorName_index)):
                  print(f"Error Cell Index: {ErrorName_index[i]}")
                  print(f"Error Cell Value: {ErrorName_value[i]}")

        # 节点序号检查
        ErrorNumber_flag = 0; ErrorNumber_index = []; ErrorNumber_value = []
        ErrorNumber_flag, ErrorNumber_index, ErrorNumber_value = myFct.CheckNumber(worksheet)

        if ErrorNumber_flag:
            for i in range(len(ErrorNumber_index)):
                  print(f"Error Cell Index: {ErrorNumber_index[i]}")
                  print(f"Error Cell Value: {ErrorNumber_value[i]}")

        # 节点内容检查
        ErrorContent_flag = 0; ErrorContent_cell = []; ErrorContent_value = []
        ErrorContent_flag, ErrorContent_cell, ErrorContent_value = myFct.CheckContent(worksheet)

        if ErrorContent_flag:
            for i in range(len(ErrorContent_cell)):
                    print(f"Error Cell Index: {ErrorContent_cell[i]}")
                    print(f"Error Cell Value: {ErrorContent_value[i]}")

