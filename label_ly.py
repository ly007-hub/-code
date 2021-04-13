import xlrd
from label_writing import *

"""
def read_excel(file_path):
    # 获取数据
    data = xlrd.open_workbook(file_path)
    # 获取所有sheet名字
    sheet_names = data.sheet_names()
    for sheet in sheet_names:
        # 获取sheet
        table = data.sheet_by_name(sheet)
        # 获取总行数
        nrows = table.nrows  # 包括标题
        # 获取总列数
        ncols = table.ncols

        # 计算出合并的单元格有哪些
        colspan = {}
        if table.merged_cells:
            for item in table.merged_cells:
                for row in range(item[0], item[1]):
                    for col in range(item[2], item[3]):
                        # 合并单元格的首格是有值的，所以在这里进行了去重
                        if (row, col) != (item[0], item[2]):
                            colspan.update({(row, col): (item[0], item[2])})
        # 读取每行数据
        for i in range(1, nrows):
            row = []
            for j in range(ncols):
                # 假如碰见合并的单元格坐标，取合并的首格的值即可
                if colspan.get((i, j)):
                    row.append(table.cell_value(*colspan.get((i, j))))
                else:
                    row.append(table.cell_value(i, j))

        # # 读取每列数据
        # for j in range(ncols):
        #     col = []
        #     for i in range(1, nrows):
        #         # 假如碰见合并的单元格坐标，取合并的首格的值即可
        #         if colspan.get((i, j)):
        #             col.append(table.cell_value(*colspan.get((i, j))))
        #         else:
        #             col.append(table.cell_value(i, j))


def read_excel_1(file_path):
    workbook = xlrd.open_workbook(file_path)
    Data_sheet = workbook.sheets()[0]  # 通过索引获取

    rowNum = Data_sheet.nrows  # sheet行数
    colNum = Data_sheet.ncols  # sheet列数

    for i in range(25):
        print(str(Data_sheet.cell_value(7, i)))

    list = []
    for i in range(5, 4770):
        rowlist = []
        for j in range(25):
            rowlist.append(Data_sheet.cell_value(i, j))
        list.append(rowlist)
"""

def read_excel_2(file_path):
    workbook = xlrd.open_workbook(file_path)
    if file_path.split('.xl')[1] == 's':
        workbook = xlrd.open_workbook(file_path, formatting_info=True)
    # 获取sheet
    sheet = workbook.sheet_by_index(0)
    # 获取行数
    r_num = sheet.nrows
    # 获取列数
    c_num = sheet.ncols
    merge = sheet.merged_cells
    # print(merge)  # [(1, 5, 0, 1), (1, 5, 1, 2)], 对应上面两个合并的单元格

    read_data = []
    for r in range(r_num):
        li = []
        for c in range(c_num):
            # 读取每个单元格里的数据，合并单元格只有单元格内的第一行第一列有数据，其余空间都为空
            cell_value = sheet.row_values(r)[c]
            # 判断空数据是否在合并单元格的坐标中，如果在就把数据填充进去
            if cell_value is None or cell_value == '':
                for (rlow, rhigh, clow, chigh) in merge:
                    if rlow <= r < rhigh:
                        if clow <= c < chigh:
                            cell_value = sheet.cell_value(rlow, clow)
            li.append(cell_value)
        read_data.append(li)

    # 删除前面几行多余数据
    for i in range(4):
        del read_data[i]

    # 删除多余数据
    for j in range(20):
        try:
            for i in range(4774, len(read_data)):
                del read_data[i]
        except:
            pass


    for j in range(20):
        try:
            for i in range(len(read_data)):
                # print(read_data[i][0])
                if(read_data[i][2] == ''):
                    del read_data[i]
        except:
            pass

    for i in range(len(read_data)):
        # print(read_data[i][0])
        try:
            if (read_data[i][0] == ''):
                del read_data[i]
        except:
            pass
    return read_data


file_path = r'E:\ly\超声小组\label_ly\1 @ 20210406广一李主任500例数据\ai2019.xlsx'
label = read_excel_2(file_path)
"""
for i in label:
    print(i)
    
for i in ids:
    print(i)
"""
workwheet, ids = write_id_xlsx()
write_label_xlsx(workwheet, label, ids)
# read_excel_2(file_path)