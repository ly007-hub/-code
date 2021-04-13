import xlrd
from itertools import groupby
from xlutils.copy import copy
from collections import Counter
import numpy as np

def get_ids():
    file_path = r'E:\ly\超声小组\label_ly\label模板.xlsx'
    data = xlrd.open_workbook(file_path)
    # 获取所有sheet名字
    sheet_names = data.sheet_names()

    table = data.sheet_by_name(sheet_names[0])

    nrows = table.nrows  # 包括标题

    # 获取总列数

    ncols = table.ncols
    id = list()
    col = []
    for i in range(1, nrows):
        col.append(table.cell_value(i, 0))
    for i in range(len(col)):
        s = col[i]
        ss = [''.join(list(g)) for k, g in groupby(s, key=lambda x: x.isdigit())]
        if ss == []:
            id.append(ss)
        else:
            ss = ss[0] + ss[1]

        id.append(ss)

    res = [ele for ele in id if ele != []]

    # 删除空格
    for i in res:
        if i == []:
            del i

    return res

def write_id_xlsx():
    file_path = r'E:\ly\超声小组\label_ly\label模板.xlsx'
    save_path = r'E:\ly\超声小组\label_ly\label_afterdoing.xlsx'
    id = get_ids()
    data = xlrd.open_workbook(file_path)
    workwheet = copy(data)
    table = workwheet.get_sheet(0)
    for i in range(1, len(id)+1):
        table.write(i, 0, id[i-1])

    workwheet.save(save_path)

    return workwheet, id

def information_pick(label_list):
    # label_list = label
    lists = []
    for i in range(len(label_list)):
        lists.append(label_list[i][0])



    # 找出重复元素
    list_dict = Counter(lists)

    for j in range(5):
        for i in range(len(label_list)):
            try:
                if list_dict[label_list[i][0]] == 1:
                    pass
                else:
                    del label_list[i]
            except:
                pass

    """
    for i in label_list:
        print(i)
    """
    return label_list

def label_end_write(workwheet=None, ids=None, information=None):
    # information = label_after_doing
    """
    for i in information:
        print(i)
    """
    save_path = r'E:\ly\超声小组\label_ly\label_afterdoing.xlsx'
    table = workwheet.get_sheet(0)
    # 匹配
    for i in information:
        for index_j, j in enumerate(ids):
            if i[0] == j:
                for index, k in enumerate(i):
                    # 写入数据
                    if k == 0:
                        continue
                    table.write(index_j+1, index, k)

    workwheet.save(save_path)



def write_label_xlsx(workwheet=None, label=None, ids=None):
    if workwheet==None or label==None:
        exit('something wrong')
    for i in range(len(label)):
        print(label[i])

    # 判断是否为只有一个病灶
    label_after_doing = information_pick(label)
    """
    for i in label_after_doing:
        print(i)
    """

    # 如果是则写入数据
    label_end_write(workwheet, ids, label_after_doing)






