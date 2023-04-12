import xlwings as xw
import os

content0 = []
content1 = []

with xw.App(visible = False, add_book = False) as app:  # 启动Excel程序
    workbook = app.books.open('(公开)低频电缆网20190920（简化版）.xls')
    worksheets = workbook.sheets  # 获取工作簿中所有的工作表

    for worksheet in worksheets:
        rng = worksheet.used_range
        Ncells = rng.count
        Nrows = rng.rows.count
        Ncols = rng.columns.count

        for row in range(6, Nrows+1):
            temp0 = worksheet.range('C'+str(row)).value
            if temp0 == None:
                temp0 = 'None'
            temp1 = worksheet.range('F'+str(row)).value
            if temp1 == None:
                temp1 = 'None'
            if temp0 != temp1:
                content0.append(temp0)
                content1.append(temp1)

dic = {} 
for index, item in enumerate(content0):
    dic_value = []
    for value in dic.values():
        dic_value = dic_value + value
    if item not in dic.keys():
        if item not in dic_value:
            if content1[index] not in dic.keys():
                if content1[index] not in dic_value:
                    dic[item] = []
                    dic[item].append(content1[index])
                else:
                    for key, value in dic.items():
                        if content1[index] in value:
                            dic[key].append(item)
            else:
                dic[content1[index]].append(item)
        else:
            for key, value in dic.items():
                if item in value:
                    dic[key].append(content1[index])
    else:
        dic[item].append(content1[index])

for key, value in dic.items():
    dic[key] = list(set(value))

with open("corpus.txt", "w", encoding="utf-8") as corpus:
    corpus.truncate(0)  # 检查报告初始化
    for key, value in dic.items():
        str = key
        for item in value:
            str = str + '\t\t' + item
        str = str + '\n'
        corpus.write(str)
