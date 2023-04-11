# file:demo3.py
# fuction:调试同义词识别

# 头文件导入
import xlwings as xw  # 导入xlwings模块
import pandas as pd
import numpy as np
from tabulate import tabulate
import jieba
import os

# 宏定义
SHEETNUMBER = 0  # 工作表序号

# 程序主体

# 加载同义词词典
dictionary = open("chinese synonym.txt","r",encoding='utf-8')  # 读取txt文件
synonyms = {}  # 定义字典
for line in dictionary:
    word = line.strip().split("\t")
    num = len(word)
    for i in range(0, num):
        synonyms[word[i]] = word[0] # synonyms的每个键的值是列表的第一个内容
# print(synonyms)


app = xw.App(visible = False, add_book = False)  # 启动Excel程序
workbook = app.books.open('(公开)低频电缆网20190920（简化版）.xls')
worksheets = workbook.sheets  # 获取工作簿中所有的工作表

worksheet = worksheets[SHEETNUMBER]  # 选中工作表
rng = worksheet.used_range
Ncells = rng.count
Nrows = rng.rows.count
Ncols = rng.columns.count

startnode_content = worksheet.range('C6:C'+str(Nrows)).value  # 记录起始节点的内容
endnode_content = worksheet.range('F6:F'+str(Nrows)).value  # 记录终止节点的内容

# 将内容中的“空”变为 None
for i in range(len(startnode_content)):
    if startnode_content[i] == "空":
        startnode_content[i] = None
for i in range(len(endnode_content)):
    if endnode_content[i] == "空":
        endnode_content[i] = None

content_error_flag = 0  # 记录错误标志 出错为'1' 无错为'0'

# 比较内容是否一致
for index in range(len(startnode_content)):
    if startnode_content[index] != endnode_content[index]:
        if synonyms.get(startnode_content[index]) == None or synonyms.get(endnode_content[index]) == None:
            print(f"error row = {index+6}")
            print(f"{startnode_content[index]}     {endnode_content[index]}  ")
            content_error_flag = 1
        elif synonyms[startnode_content[index]] != synonyms[endnode_content[index]]:
            print(f"error row = {index+6}")
            content_error_flag = 1
    

if content_error_flag == 0:
    print("节点内容检查通过")
else:
    print("节点内容检查不通过")


workbook.close()  # 关闭工作簿
app.quit()  # 关闭Excel程序