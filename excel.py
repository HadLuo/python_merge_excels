# -*- coding: utf-8 -*-


# 导入需要使用的包
import xlrd  # 读取Excel文件的包
import xlsxwriter  # 将文件写入Excel的包
import sys
import os
from tkinter import messagebox
import yaml


def file_name(file_dir, fix):
    L = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if os.path.splitext(file)[1] == fix:
                L.append(os.path.join(root, file))
    return L


# 打开一个excel文件
def open_xls(file):
    f = xlrd.open_workbook(file)
    return f


# 加载 yml ,返回字典类型
def readYml(yamlPath):
    f = open(yamlPath, 'r', encoding='utf-8')
    return yaml.load(f.read(), Loader=yaml.FullLoader)


conf = {}
try:
    conf = readYml(os.path.abspath(os.curdir) + "/conf.txt")
except:
    pass
if not conf:
    conf = {
        "remove-title": 0
    }

xlsx_fils = file_name(os.path.abspath(os.curdir), '.xlsx')
xlsx_files = []
for i in xlsx_fils:
    if "合并后.xlsx" not in i:
        xlsx_files.append(i)

print(xlsx_files)
content = []
workBooks = []
for item in xlsx_files:
    workBooks.append(open_xls(item))

if not workBooks:
    messagebox.showinfo("提示琪", "当前文件夹没有要合并的excel")
    sys.exit(0)
if len(workBooks) == 1:
    messagebox.showinfo("提示琪", "当前文件夹就只有一个excel，无需合并")
    sys.exit(0)

names = []
contents = []
## 打開第一個
workBook = workBooks[0]

contains_wk = []
contains_wk.append(workBook)
## 遍历表单 名称
allSheetNames = workBook.sheet_names();
for name in allSheetNames:
    rows = []
    names.append(name)
    ## 根据sheet名称获取sheet
    sheet = workBook.sheet_by_name(name);
    ##获取sheet 行数
    nrows = sheet.nrows
    ## 获取sheet 列数
    ncols = sheet.ncols
    for r in range(0, nrows):
        ## 行的内容
        rows.append(sheet.row_values(r))

    ## 找出下一个 的内容
    for ws in workBooks:
        if ws in contains_wk:
            continue
        sheet = workBook.sheet_by_name(name);
        ##获取sheet 行数
        nrows = sheet.nrows
        for r in range(0, nrows):
            if "remove-title" in conf and conf["remove-title"] == 1 and r == 0:
                continue
            ## 行的内容
            print("添加行: " + str(sheet.row_values(r)))
            rows.append(sheet.row_values(r))
        break
    contents.append(rows)

print(names)
print(rows)

### 写excel
wb = xlsxwriter.Workbook(os.path.abspath(os.curdir) + "/合并后.xlsx")

for name in names:
    ws = wb.add_worksheet("name")
    for a in range(len(rows)):
        for b in range(len(rows[a])):
            c = rows[a][b]
            ws.write(a, b, c)
wb.close()
messagebox.showinfo("提示琪", "恭喜琪，合并成功！生成的文件为：" + os.path.abspath(os.curdir) + "/合并后.xlsx")
