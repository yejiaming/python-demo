# xlrd只能读取xls
import xlrd

book = xlrd.open_workbook('test.xlsl', on_demand=True)  # 当on_demand=True，只有被要求时才将worksheet载入内存（懒加载），一般在读取大文件时使用

# 表格的操作
print("excel中存在多少个sheets", book.nsheets)
print("sheeet列表", book.sheet_names()[0])

# 打开表的三种方法
sh = book.sheet_by_index(0)
sh = book.sheets()[0]
sh = book.sheet_by_name('Sheet1')

# 操作行列和单元格
print("sheet名称:" + sh.name, "sheet行数:" + str(sh.nrows), "sheet列数:" + str(sh.ncols))
print("第3行第一列数值：", sh.cell_value(rowx=10, colx=0))
print("第3行第一列数值：", sh.cell(2, 0).value)
print("第3行第一列数据类型：", sh.cell_type(2, 0))

# 操作列
print("第一列的所有数据单元格列表：", sh.col(0))
print("第3列的第1行到第6行数据列表：", sh.col_slice(2, start_rowx=0, end_rowx=5))
print("第3列的第1行到第6行数据单元格列表：", sh.col_slice(2, start_rowx=0, end_rowx=5))
print("第3列的第1行到第6行数据类型列表：", sh.col_types(2, start_rowx=0, end_rowx=5))
print("第3列的第1行到第6行数据值列表：", sh.col_values(2, start_rowx=0, end_rowx=5))

# 操作行（和列有同样的类似方法）
print("第一行的所有数据单元格列表：", sh.row(0))

# 循环行
for rx in range(sh.nrows):
    print("第" + str(rx + 1) + "行的数据单元格列表是：", sh.row(rx))

# 循环列
for cx in range(sh.ncols):
    print("第" + str(cx + 1) + "列的数据单元格列表是：", sh.col(cx))
