import xlwings as xw  # 导入xlwings模块
import time
import myFct


start_time = time.time()
Error_flag = 0; Error_sheet = []; Error_cell = []; Error_value = []

with xw.App(visible = False, add_book = False) as app:  # 启动Excel程序

    workbook = app.books.open('(公开)低频电缆网20190920（简化版）.xls')
    worksheets = workbook.sheets  # 获取工作簿中所有的工作表

    for worksheet in worksheets:

        print(worksheet.name)

        # 节点名称检查
        ErrorName_flag = 0; ErrorName_cell = []; ErrorName_value = []
        ErrorName_flag, ErrorName_cell, ErrorName_value = myFct.CheckName(worksheet)

        # if ErrorName_flag:
        #     for i in range(len(ErrorName_cell)):
        #           print(f"Error Cell: {ErrorName_cell[i]}")
        #           print(f"Error Value: {ErrorName_value[i]}")

        # 节点序号检查
        ErrorNumber_flag = 0; ErrorNumber_cell = []; ErrorNumber_value = []
        ErrorNumber_flag, ErrorNumber_cell, ErrorNumber_value = myFct.CheckNumber(worksheet)

        # if ErrorNumber_flag:
        #     for i in range(len(ErrorNumber_cell)):
        #           print(f"Error Cell: {ErrorNumber_cell[i]}")
        #           print(f"Error Value: {ErrorNumber_value[i]}")

        # 节点内容检查
        ErrorContent_flag = 0; ErrorContent_cell = []; ErrorContent_value = []
        ErrorContent_flag, ErrorContent_cell, ErrorContent_value = myFct.CheckContent(worksheet)

        # if ErrorContent_flag:
        #     for i in range(len(ErrorContent_cell)):
        #             print(f"Error Cell: {ErrorContent_cell[i]}")
        #             print(f"Error Value: {ErrorContent_value[i]}")
        
        Error_flag = ErrorName_flag or ErrorNumber_flag or ErrorContent_flag or Error_flag
        Error_cell.extend(ErrorName_cell + ErrorNumber_cell + ErrorContent_cell)
        Error_value.extend(ErrorName_value + ErrorNumber_value + ErrorContent_value)
        for i in range(len(ErrorName_cell + ErrorNumber_cell + ErrorContent_cell)):
            Error_sheet.append(worksheet.name)

    


if Error_flag:
    for i in range(len(Error_cell)):
        print(f"Error Sheet: {Error_sheet[i]}     Error Cell: {Error_cell[i]}    Error Value: {Error_value[i]}")


end_time = time.time()
elapsed_time = end_time - start_time
print(f"程序运行时间为: {elapsed_time:.2f} 秒")