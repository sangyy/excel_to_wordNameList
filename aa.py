import openpyxl
import codecs
from chinese_stroke_sorting import sort_by_stroke

# 解析xlsx并创建行列表
xlsx = "2017届初中毕结业生名册.xlsx"
workbook=openpyxl.load_workbook(xlsx)
worksheet=workbook.worksheets[0]
rows=worksheet.max_row
rowlist = []
for row in worksheet.rows:
    celllist = []
    for cell in row:
        celllist.append(str(cell.value))
    rowlist.append(celllist)
rowlist.pop(0)
# print(rowlist)

# 创建基础字典
clss_list = []
for box in rowlist:
    clss_list.append(int(box[2]))
clss_set = list(set(clss_list))
clss_dict = {}
for num in clss_set:
    clss_dict[num] = []

# 向字典内添加姓名
for student in rowlist:
    clss = int(student[2])
    name = student[3]
    clss_dict[clss].append(name)
# print(clss_dict)

# 姓名列表按照比划排序
for i in clss_set:
    clss_dict[i] = sort_by_stroke(clss_dict[i])
    utils_list = []
    for name in clss_dict[i]:
        if len(name) == 2:
            name = name[0] + "  " + name[1]
        utils_list.append(name)
    clss_dict[i] = utils_list
# print(clss_dict)

# 写入txt文件
for i in clss_set:
    with open("class_student_name.txt","a", encoding='utf-8') as txtfile:
        txtfile.write("2017年届初三" + str(i) + "班  " + str(len(clss_dict[i])) + " 人  班主任" + "\n")
        index = 0
        for name in clss_dict[i]:
            txtfile.write(str(name + "  "))
            index += 1
            if index == 12:
                txtfile.write("\n")
                index = 0
            if name == clss_dict[i][-1]:
                txtfile.write("\n")
                txtfile.write("\n")
        txtfile.close()
