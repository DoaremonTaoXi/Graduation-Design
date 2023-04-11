import os
import xlwings as xw  # 导入xlwings模块
import pandas as pd
import numpy as np
from tabulate import tabulate

# 宏定义
SHEETNUMBER = 6  # 工作表序号



report = open("report.txt", "w")  # 新建检查报告
report.truncate(0)  # 检查报告初始化


app = xw.App(visible = False, add_book = False)  # 启动Excel程序
workbook = app.books.open('(公开)低频电缆网20190920（简化版）.xls')
worksheets = workbook.sheets  # 获取工作簿中所有的工作表


'''
分析内容一：节点名称是否一致
'''
print('---------检查节点名称---------')
report.write('---------检查节点名称---------\n')
worksheet = worksheets[SHEETNUMBER]  # 选中工作表
rng = worksheet.used_range
Ncells = rng.count
Nrows = rng.rows.count
Ncols = rng.columns.count

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
if set(startnode0)==set(startnode1):
    print("起始节点名称检查通过")
    report.write('起始节点名称检查通过\n')
else:
    print("起始节点名称检查不通过")
    report.write("起始节点名称检查不通过\n")
    report.write("错误内容：\n")
    report.write("   SHEETNAME     NODE0     NODE1\n")
    report.write(f"{worksheet.name} {startnode0} {startnode1}\n")

if set(endnode0)==set(endnode1):
    print("终止节点名称检查通过")
    report.write('终止节点名称检查通过\n')
else:
    print("终止节点名称检查不通过")
    report.write('终止节点名称检查不通过\n')
    report.write("错误内容：\n")
    report.write("   SHEETNAME     NODE0     NODE1\n")
    report.write(f"   {worksheet.name}     {endnode0}     {endnode1}\n")


'''
分析内容二：节点序号是否重复或连续
'''
print('---------检查节点序号---------')
report.write('\n---------检查节点序号---------\n')

# 起始节点序号检查
start_number_error_flag = 0  # 记录错误标志 出错为'1' 无错为'0'
startindex = 6  # 记录序号起点行数
start_errorindex = []  # 记录错误节点序号行数

for i in range(7,Nrows+1):
    if worksheet.range('A'+str(i)).value is not None:
        endindex = i-1  # 记录序号终点行数
        old_number = worksheet.range('B'+str(startindex)+':B'+str(endindex)).value
        new_number = []  # new_dic初始化
        new_number_index = []
        # 移除","和"~"的影响
        for index, item in enumerate(old_number):
            if isinstance(item, str):
                temp = item.split(',')
                for element in temp:
                    if '~' in element:
                        start, end = element.split('~')
                        for j in range(int(start), int(end)+1):
                            new_number.append(j)
                            new_number_index.append(index)
                    else:
                        new_number.append(int(element))
                        new_number_index.append(index)
            else:
                new_number.append(item)
                new_number_index.append(index)
        # 去除列表中为None的部分
        for index, item in enumerate(new_number):
            if item == None:
                del new_number[index]
                del new_number_index[index]

        # 判断序号是否重复或连续
        sorted_number = sorted(new_number)  # 序号重排 按从小到大的顺序
        for j in range(len(sorted_number)-1):
            if sorted_number[j+1] - sorted_number[j] != 1:
                start_number_error_flag = 1
                # 记录错误序号原先的位置
                for index, item in enumerate(new_number):
                    if item == sorted_number[j+1]:
                        start_errorindex.append(new_number_index[index] + startindex)

        startindex = endindex + 1  # 更新序号起点
    else:
        if i == Nrows:  # 处理最后一行的情况
            endindex = Nrows
            old_number = worksheet.range('B'+str(startindex)+':B'+str(endindex)).value
            new_number = []  # new_dic初始化
            new_number_index = []
            # 移除","和"~"的影响
            for index, item in enumerate(old_number):
                if isinstance(item, str):
                    temp = item.split(',')
                    for element in temp:
                        if '~' in element:
                            start, end = element.split('~')
                            for j in range(int(start), int(end)+1):
                                new_number.append(j)
                                new_number_index.append(index)
                        else:
                            new_number.append(int(element))
                            new_number_index.append(index)
                else:
                    new_number.append(item)
                    new_number_index.append(index)
            # 去除列表中为None的部分
            for index, item in enumerate(new_number):
                if item == None:
                    del new_number[index]
                    del new_number_index[index]

            # 判断序号是否重复或连续
            sorted_number = sorted(new_number)  # 序号重排 按从小到大的顺序
            for j in range(len(sorted_number)-1):
                if sorted_number[j+1] - sorted_number[j] != 1:
                    start_number_error_flag = 1
                    # 记录错误序号原先的位置
                    for index, item in enumerate(new_number):
                        if item == sorted_number[j+1]:
                            start_errorindex.append(new_number_index[index] + startindex)
                    
        else:
            continue  # 不为最后一行，继续往下寻找内容不为NONE的单元格

if start_number_error_flag == 1:
    print("起始节点序号检查不通过")
    report.write("起始节点序号检查不通过\n")
    report.write("错误内容：\n")
    report.write("   SHEETNAME     CELL     CONTENT\n")
    start_errorindex = list(set(start_errorindex))  # 去除重复元素
    for index in start_errorindex:
        report.write(f"   {worksheet.name}     B{index}     {worksheet.range('B'+str(index)).value}\n")
else:
    print("起始节点序号检查通过")
    report.write("起始节点序号检查通过\n")

# 终止节点序号检查
end_number_error_flag = 0  # 记录错误标志 出错为'1' 无错为'0'
startindex = 6  # 记录序号起点行数
end_errorindex = []  # 记录错误节点序号行数

for i in range(7,Nrows+1):
    if worksheet.range('D'+str(i)).value is not None:
        endindex = i-1  # 记录序号终点行数
        old_number = worksheet.range('E'+str(startindex)+':E'+str(endindex)).value
        new_number = []  # new_dic初始化
        new_number_index = []
        # 移除","和"~"的影响
        for index, item in enumerate(old_number):
            if isinstance(item, str):
                temp = item.split(',')
                for element in temp:
                    if '~' in element:
                        start, end = element.split('~')
                        for j in range(int(start), int(end)+1):
                            new_number.append(j)
                            new_number_index.append(index)
                    else:
                        new_number.append(int(element))
                        new_number_index.append(index)
            else:
                new_number.append(item)
                new_number_index.append(index)
        # 去除列表中为None的部分
        for index, item in enumerate(new_number):
            if item == None:
                del new_number[index]
                del new_number_index[index]

        # 判断序号是否重复或连续
        sorted_number = sorted(new_number)  # 序号重排 按从小到大的顺序
        for j in range(len(sorted_number)-1):
            if sorted_number[j+1] - sorted_number[j] != 1:
                end_number_error_flag = 1
                # 记录错误序号原先的位置
                for index, item in enumerate(new_number):
                    if item == sorted_number[j+1]:
                        end_errorindex.append(new_number_index[index] + startindex)

        startindex = endindex + 1  # 更新序号起点
    else:
        if i == Nrows:  # 处理最后一行的情况
            endindex = Nrows
            old_number = worksheet.range('E'+str(startindex)+':E'+str(endindex)).value
            new_number = []  # new_dic初始化
            new_number_index = []
            # 移除","和"~"的影响
            for index, item in enumerate(old_number):
                if isinstance(item, str):
                    temp = item.split(',')
                    for element in temp:
                        if '~' in element:
                            start, end = element.split('~')
                            for j in range(int(start), int(end)+1):
                                new_number.append(j)
                                new_number_index.append(index)
                        else:
                            new_number.append(int(element))
                            new_number_index.append(index)
                else:
                    new_number.append(item)
                    new_number_index.append(index)
            # 去除列表中为None的部分
            for index, item in enumerate(new_number):
                if item == None:
                    del new_number[index]
                    del new_number_index[index]

            # 判断序号是否重复或连续
            sorted_number = sorted(new_number)  # 序号重排 按从小到大的顺序
            for j in range(len(sorted_number)-1):
                if sorted_number[j+1] - sorted_number[j] != 1:
                    end_number_error_flag = 1
                    # 记录错误序号原先的位置
                    for index, item in enumerate(new_number):
                        if item == sorted_number[j+1]:
                            end_errorindex.append(new_number_index[index] + startindex)
        else:
            continue  # 不为最后一行，继续往下寻找内容不为NONE的单元格

if end_number_error_flag == 1:
    print("终止节点序号检查不通过")
    report.write("终止节点序号检查不通过\n")
    report.write("错误内容：\n")
    report.write("   SHEETNAME     CELL     CONTENT\n")
    end_errorindex = list(set(end_errorindex))  # 去除重复元素
    for index in end_errorindex:
        report.write(f"   {worksheet.name}     E{index}     {worksheet.range('E'+str(index)).value}\n")
else:
    print("终止节点序号检查通过")
    report.write("终止节点序号检查通过\n")

'''
分析内容三：节点内容是否一致
'''
print('---------检查节点内容---------')
report.write('\n---------检查节点内容---------\n')
startnode_content = worksheet.range('C6:C'+str(Nrows)).value  # 记录起始节点的内容
endnode_content = worksheet.range('F6:F'+str(Nrows)).value  # 记录终止节点的内容

content_error_flag = 0  # 记录错误标志 出错为'1' 无错为'0'

# 比较内容是否一致
for index in range(len(startnode_content)):
    if startnode_content[index] != endnode_content[index]:
        content_error_flag = 1
        report.write(f"{worksheet.name} 'C{index+6}':{startnode_content[index]} 'F{index+6}':{endnode_content[index]}\n")

if content_error_flag == 0:
    print("节点内容检查通过")
    report.write("节点内容检查通过\n")
else:
    print("节点内容检查不通过")
    report.write("节点内容检查不通过\n")


workbook.close()  # 关闭工作簿
app.quit()  # 关闭Excel程序

report.close()  # 关闭检查报告

