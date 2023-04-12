import xlwings as xw  # 导入xlwings模块
import re
import os

"""
自定义函数
"""
# Name:     pretreatNodeName0
# Function: 处理node0名称含有“、\n”和“\n”的情况
# Input:    nodename0(str)
# Output:   temp(list)
def pretreatNodeName0(nodename0):
    temp = re.split(r'、\n|\n',nodename0)
    if isinstance(temp, str):
        temp = [nodename0]  # 将字符串以列表的形式储存
    return temp

# Name:     pretreatNodeName1
# Function: 处理node1名称含有“\n”的情况
# Input:    nodename1(list)
# Output:   temp(list)
def pretreatNodeName1(nodename1):
    temp = nodename1.replace("\n", "")
    return temp

# Name:     getNodeName0
# Function: 
# Input:    worksheet(worksheet):工作簿     flag(int):0-读取起点   1-读取终点
# Output:   dic(dict):  {'DY-X11(J30JHT100ZKSAB02)':'B3'}     {'MB-X1(J30JHT31TJSAB02)':'E3','MB-X2(J30JHT69TJSAB02)':'E3'}
def getNodeName0(worksheet,flag):
    dic = {}
    if not flag:
        nodename0 = worksheet.range('B3').value
        nodename0 = pretreatNodeName0(nodename0)
    else:
        nodename0 = worksheet.range('E3').value
        nodename0 = pretreatNodeName0(nodename0)
    for item in nodename0:
        dic.setdefault(item, []).append('B3')
    return dic

# Name:     getNodeName1
# Function: 
# Input:    worksheet(worksheet):工作簿     flag(int):0-读取起点   1-读取终点
# Output:   dic(dict):  {'DY-X11(J30JHT100ZKSAB02)':'A6'}     {'MB-X1(J30JHT31TJSAB02)':'D6','MB-X2(J30JHT69TJSAB02)':'D37'}
def getNodeName1(worksheet, flag):
    rng = worksheet.used_range
    Ncells = rng.count
    Nrows = rng.rows.count
    Ncols = rng.columns.count
    
    dic = {}
    if not flag:
        for row in range(6, Nrows+1):
            if worksheet.range('A'+str(row)).value is not None:
                nodename1 = worksheet.range('A'+str(row)).value
                nodename1 = pretreatNodeName1(nodename1)
                dic.setdefault(nodename1, []).append('A'+str(row))
    else:
        for row in range(6, Nrows+1):
            if worksheet.range('D'+str(row)).value is not None:
                nodename1 = worksheet.range('D'+str(row)).value
                nodename1 = pretreatNodeName1(nodename1)
                dic.setdefault(nodename1, []).append('D'+str(row))
    return dic

# Name:     assayName
# Function: 节点名称分析
# Input:    dic0(dict)  dic1(dict)
# Output:   Error_flag(int):0\1     Error_dic(dict)
def assayName(dic0, dic1):
    # 获取字典的键，并转换为集合
    keys0 = set(dic0.keys())
    keys1 = set(dic1.keys())
    # 计算两个字典的键的差集，即不同的键
    diff_keys = keys0.symmetric_difference(keys1)
    # 遍历一个字典，将不同的键值对存入Error_dic
    Error_flag = 0
    Error_dic = {}
    if diff_keys:
        Error_flag = 1
        for key in keys0:
            if key in diff_keys:
                Error_dic[key] = dic0[key]
        for key in keys1:
            if key in diff_keys:
                Error_dic[key] = dic1[key]
    
    return (Error_flag, Error_dic)

# Name:     CheckName
# Function: 节点名称检查主函数
# Input:    worksheet(worksheet):工作簿
# Output:   ErrorName_flag(int):0\1     ErrorName_index(list)       ErrorName_value(list)
def CheckName(worksheet):
    rng = worksheet.used_range
    Ncells = rng.count
    Nrows = rng.rows.count
    Ncols = rng.columns.count

    ErrorName_flag = 0
    ErrorName_index = []
    ErrorName_value = []

    StartNode0 = getNodeName0(worksheet, 0)
    # print(StartNode0)
    StartNode1 = getNodeName1(worksheet, 0)
    # print(StartNode1)

    # 比较内容是否一致，并输出结果
    ErrorStartName_flag = 0
    ErrorStartName_dic = {}
    ErrorStartName_flag, ErrorStartName_dic = assayName(StartNode0, StartNode1)

    EndNode0 = getNodeName0(worksheet, 1)
    # print(EndNode0)
    EndNode1 = getNodeName1(worksheet, 1)
    # print(EndNode1)


    # 比较内容是否一致，并输出结果
    ErrorEndName_flag = 0
    ErrorEndName_dic = {}
    ErrorEndName_flag, ErrorEndName_dic = assayName(EndNode0, EndNode1)

    ErrorName_flag = ErrorStartName_flag or ErrorEndName_flag
    for key, value in ErrorStartName_dic.items():
        for i in range(len(value)):
            ErrorName_value.append(key)
        ErrorName_index.extend(value)
    for key, value in ErrorEndName_dic.items():
        for i in range(len(value)):
            ErrorName_value.append(key)
        ErrorName_index.extend(value)
    
    return (ErrorName_flag, ErrorName_index, ErrorName_value)

# Name:     pretreatNumber
# Function: 处理序号中 "14, 15, 16~17", None 等情况
# Input:    list(list)
# Output:   处理后的序号内容 list类型   序号对应的位置 list类型
def pretreatNumber(list):
    new_number = []  # 记录处理过后的序号
    new_number_index = []  # 记录序号的原始位置
    for item in list:
        if isinstance(item[1],str):
            temp = re.split(r',|，|,\n|，\n',item[1])
            for element in temp:
                if '~' in element:
                    start, end = element.split('~')
                    for i in range(int(start), int(end)+1):
                        new_number.append(i)
                        new_number_index.append(item[0])
                else:
                    new_number.append(int(element))
                    new_number_index.append(item[0])
        else:
            new_number.append(item[1])
            new_number_index.append(item[0])
    # 去除列表中None的部分
    new_number, new_number_index = removeNone(new_number, new_number_index)
    
    return (new_number, new_number_index)

# Name:     removeNone
# Function: 处理序号中 None 等情况
# Input:    list(list)      list_index(list)
# Output:   new_list(list)  new_list_index(list)
def removeNone(list, list_index):
    new_list = []
    new_list_index = []
    for index, item in enumerate(list):
        if item is not None:
            new_list.append(int(item))
            new_list_index.append(list_index[index])

    return (new_list, new_list_index)

# Name:     assayNumber
# Function: 节点序号分析
# Input:    number(list)    number_index(list)     ErrorNumber_flag(list)      ErrorNumber_index(list)
# Output:   ErrorNumber_flag(int):0\1   ErrorNumber_index(list)
def assayNumber(number, number_index, ErrorNumber_flag, ErrorNumber_index):
    sorted_number = sorted(number)  # 序号重排 按从小到大的顺序
    for i in range(len(sorted_number)-1):
        if sorted_number[i+1] - sorted_number[i] != 1:
            ErrorNumber_flag = 1
            # 记录错误序号的位置
            for index, item in enumerate(number):
                if item == sorted_number[i+1]:
                    ErrorNumber_index.append(number_index[index])
    
    ErrorNumber_index = list(set(ErrorNumber_index))
    return (ErrorNumber_flag, ErrorNumber_index)

# Name:     getNumber
# Function: 读取序号内容
# Input:    worksheet(worksheet):工作簿    flag(int):0-读取起点   1-读取终点
# Output:   dic(dict):键为 Node Name   值为 list   list中的元素为tuple (row, value)
def getNumber(worksheet, flag):
    rng = worksheet.used_range
    Ncells = rng.count
    Nrows = rng.rows.count
    Ncols = rng.columns.count

    dic = {}
    if not flag:
        for row in range(6,Nrows+1):
            if worksheet.range('A'+str(row)).value is not None:
                key = worksheet.range('A'+str(row)).value
                key = key.replace('\n', '')
                tuple = (row, worksheet.range('B'+str(row)).value)
                dic.setdefault(key, []).append(tuple)
            else:
                tuple = (row, worksheet.range('B'+str(row)).value)
                dic.setdefault(key, []).append(tuple)
    else:
        for row in range(6,Nrows+1):
            if worksheet.range('D'+str(row)).value is not None:
                key = worksheet.range('D'+str(row)).value
                key = key.replace('\n', '')
                tuple = (row, worksheet.range('E'+str(row)).value)
                dic.setdefault(key, []).append(tuple)
            else:
                tuple = (row, worksheet.range('E'+str(row)).value)
                dic.setdefault(key, []).append(tuple)
    
    return dic
    
# Name:     CheckNumber
# Function: 节点序号检查主函数
# Input:    worksheet(worksheet):工作簿
# Output:   ErrorNumber_flag(int):0/1       ErrorNumber_index(list)     ErrorNumber_value(list)
def CheckNumber(worksheet):
    rng = worksheet.used_range
    Ncells = rng.count
    Nrows = rng.rows.count
    Ncols = rng.columns.count
    
    ErrorNumber_flag = 0
    ErrorNumber_index = []
    ErrorNumber_value = []

    # StartNumber检查
    ErrorStartNumber_flag = 0
    ErrorStartNumber_index = []

    dic0 = getNumber(worksheet,0)
    for key, value in dic0.items():
        new_number, new_number_index = pretreatNumber(value)
        ErrorStartNumber_flag, ErrorStartNumber_index = assayNumber(new_number, new_number_index, ErrorStartNumber_flag, ErrorStartNumber_index)

    ErrorStartNumber_index = ['B%d' % element for element in ErrorStartNumber_index]
    

    #EndNumber检查
    ErrorEndNumber_flag = 0
    ErrorEndNumber_index = []
    dic1 = getNumber(worksheet,1)
    for key, value in dic1.items():
        new_number, new_number_index = pretreatNumber(value)
        ErrorEndNumber_flag, ErrorEndNumber_index = assayNumber(new_number, new_number_index, ErrorEndNumber_flag, ErrorEndNumber_index)

    ErrorEndNumber_index = ['E%d' % element for element in ErrorEndNumber_index]
    

    ErrorNumber_flag = ErrorStartNumber_flag or ErrorEndNumber_flag
    ErrorNumber_index = ErrorStartNumber_index + ErrorEndNumber_index
    for element in ErrorNumber_index:
        ErrorNumber_value.append(worksheet.range(element).value)

    return (ErrorNumber_flag, ErrorNumber_index, ErrorNumber_value)

# Name:     
# Function: 
# Input:    
# Output:   
def getContent(worksheet, flag):
    rng = worksheet.used_range
    Ncells = rng.count
    Nrows = rng.rows.count
    Ncols = rng.columns.count

    dic = {}
    for row in range(6, Nrows+1):
        if not flag:
            cell = 'C' + str(row)
        else:
            cell = 'F' + str(row)
        dic[row] = worksheet.range(cell).value

    return dic

# Name:     
# Function: 
# Input:    
# Output:   
def loadCorpus():
    corpus = open("corpus.txt","r",encoding='utf-8')  # 读取txt文件
    synonyms = {}
    for line in corpus:
        word = line.strip().split("\t")
        num = len(word)
        for i in range(num):
            synonyms[word[i]] = word[0] # synonyms的每个键的值是列表的第一个内容
    # print(synonyms)
    return synonyms

# Name:     
# Function: 
# Input:    
# Output:  
def pretreatContent(dic):
    for key, value in dic.items():
        if value == None:
            dic[key] = 'None'
    return dic

# Name:     
# Function: 
# Input:    
# Output:   
def assayContent(dic0, dic1, synonyms):
    dic0 = pretreatContent(dic0)
    dic1 = pretreatContent(dic1)
    
    Error_flag = 0
    Error_cell = []

    for key,_ in dic0.items():
        if dic0[key] == dic1[key]:
            continue
        elif synonyms.get(dic0[key]) or synonyms.get(dic1[key]) != None:
            if synonyms.get(dic0[key]) == synonyms.get(dic1[key]):
                continue
            else:
                Error_flag = 1
                Error_cell.append('C'+str(key))
                Error_cell.append('F'+str(key))
        else:
            Error_flag = 1
            Error_cell.append('C'+str(key))
            Error_cell.append('F'+str(key))

    return (Error_flag, Error_cell)

# Name:     
# Function: 
# Input:    
# Output:   
def CheckContent(worksheet):
    rng = worksheet.used_range
    Ncells = rng.count
    Nrows = rng.rows.count
    Ncols = rng.columns.count

    ErrorContent_flag = 0
    ErrorContent_cell = []
    ErrorContent_value = []

    StartNodeContent = getContent(worksheet,0)
    EndNodeContent = getContent(worksheet, 1)
    sysnonyms = loadCorpus()

    ErrorContent_flag, ErrorContent_cell = assayContent(StartNodeContent, EndNodeContent, sysnonyms)

    if ErrorContent_cell:
        for item in ErrorContent_cell:
            ErrorContent_value.append(worksheet.range(item).value)
    
    return (ErrorContent_flag, ErrorContent_cell, ErrorContent_value)



