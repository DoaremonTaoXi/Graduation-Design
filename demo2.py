import os
import xlwings as xw  # 导入xlwings模块
import pandas as pd
import numpy as np

app = xw.App(visible = False, add_book = False)  # 启动Excel程序
workbook = app.books.open('(公开)低频电缆网20190920（简化版）.xls')
worksheets = workbook.sheets  # 获取工作簿中所有的工作表


for worksheet in worksheets:
    # print(worksheet.name)  # 打印工作表的名称
    rng = worksheet.used_range
    Ncells = rng.count
    Nrows = rng.rows.count
    Ncols = rng.columns.count
    '''
    分析内容一：节点名称是否一致
    '''
    startnode0 = worksheet.range('B3').value  # 记录起始节点0的内容
    # 移除","和回车的影响
    if '、' in startnode0:
        startnode0 = startnode0.split('、\n')
    elif '\n' in startnode0:
        startnode0 = startnode0.split('\n')
    else:
        startnode0 = [startnode0]  # 将字符串以列表的形式储存
    # print(startnode0)

    startnode1 = worksheet.range('A6:A' + str(Nrows)).value  # 记录起始节点1的内容
    startnode1 = list(filter(None, startnode1))  # 去除列表中的空值
    startnode1 = list(set(startnode1))  # 去除重复元素
    # 移除回车的影响
    for i in range(len(startnode1)):
        if "\n" in startnode1[i]:
            startnode1[i] = startnode1[i].replace("\n", "")
    # print(startnode1)

    endnode0 = worksheet.range('E3').value  # 记录终止节点0的内容
    # 移除","和回车的影响
    if '、' in endnode0:
        endnode0 = endnode0.split('、\n')
    elif '\n' in endnode0:
        endnode0 = endnode0.split('\n')
    else:
        endnode0 = [endnode0]  # 将字符串以列表的形式储存
    # print(endnode0)

    endnode1 = worksheet.range('D6:D' + str(Nrows)).value  # 记录起始节点1的内容
    endnode1 = list(filter(None, endnode1))  # 去除列表中的空值
    endnode1 = list(set(endnode1))  # 去除重复元素
    # 移除回车的影响
    for i in range(len(endnode1)):
        if "\n" in endnode1[i]:
            endnode1[i] = endnode1[i].replace("\n", "")
    # print(endnode1)


    # 比较内容是否一致，并输出结果
    if set(startnode0)==set(startnode1) and set(endnode0)==set(endnode1):
        print(f"{worksheet.name} NODE NAME CHECK SUCCESSFULLY")
    else:
        print(f"{worksheet.name} NODE NAME CHECK FAILED")








    





workbook.close()  # 关闭工作簿
app.quit()  # 关闭Excel程序

