import os
import time
import csv

try:
    import xlrd
    import xlwt
    from xlutils.copy import copy
except ImportError as e:
    print("缺少必要的库，请运行: pip install xlrd==1.2.0 xlwt xlutils")
    raise e

# 打开文件目录，并返回文件绝对路径列表
def openfiles(filelist, dir):
    for maindir, subdir, file_name_list in os.walk(dir):
        for filename in file_name_list:
            apath = os.path.join(maindir, filename)  # 获取文件绝对路径
            # 只添加.xls和.csv文件到列表中
            if filename.endswith(('.xls', '.csv')):
                filelist.append(apath)
    return filelist


# 初始化生成一个用于保存已加工数据的文件
def initsaveexcel(filepath):
    # 使用xlwt来写excel文件
    savebook = xlwt.Workbook()
    wsheet = savebook.add_sheet('结果表')
    savebook.save(filepath)


# 复制源数据并保存到指定文件
def dealexcel(filelist, aimfilpath):
    index = 0  # 定义并设定数据保存的行索引，索引从0开始
    borders = xlwt.Borders()  # Create borders
    borders.left = xlwt.Borders.THIN  # 添加边框-实线线边框
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style = xlwt.XFStyle()  # Create style
    style.borders = borders  # Add borders to style
    wbk2 = xlrd.open_workbook(aimfilpath)  # 打开已保存文件
    newbook = copy(wbk2)
    newsheet = newbook.get_sheet(0)
    first_excel_file = True  # 标记是否为第一个Excel文件
    for filepath in filelist:  # 遍历每一个源数据文件
        print('正在处理的文件路径：', filepath)
        # 判断文件类型并分别处理
        if filepath.endswith('.csv'):
            # 处理CSV文件
            with open(filepath, 'r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                header_skipped = False
                for row in reader:
                    # 跳过第一行（标题行），除了第一个文件
                    if not header_skipped and not first_excel_file:
                        header_skipped = True
                        continue
                    index = index + 1
                    for lens in range(len(row)):
                        newsheet.write(index, lens, row[lens], style)
            first_excel_file = False
        else:
            # 处理Excel文件
            wbk = xlrd.open_workbook(filepath)  # 打开源数据文件
            # 获取第1个sheet表
            sheet1 = wbk.sheet_by_index(0)
            nrows = sheet1.nrows  # 获取源数据表行数
            for k in range(0, nrows):  # 从第1行开始遍历
                # 跳过第一行（标题行），除了第一个文件
                if k == 0 and not first_excel_file:
                    continue
                index = index + 1
                nrow_value = sheet1.row_values(k)  # 获取第i行
                for lens in range(len(nrow_value)):
                    newsheet.write(index, lens, nrow_value[lens], style)  # 写入数据
            first_excel_file = False
    newbook.save(aimfilpath)
    print('------处理完成------')
    time.sleep(5)