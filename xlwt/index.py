# xlwt只能新建xls（不可以在已有的xls上修改）
import xlwt

book = xlwt.Workbook('utf8')    # 也可以通过book.encoding = 'utf8'来设置编码
style = xlwt.easyxf(
    "font: name Arial;"
    "pattern: pattern solid, fore_colour red;"
    )

# 表格的操作
book.add_sheet('第一张表')
book.add_sheet('第二张表')
book.add_sheet('第三张表')
sh = book.get_sheet('第一张表')

# 操作行列和单元格
sh.write(0, 0, 'hello world',style)  # 编辑指定单元格
sh.row(1).write(5, 35)  # 编辑行单元格
# sh.col(1).write(3,45)   # 编辑列单元格
sh.flush_row_data()     # 将增加修改操作同步到表中，减少内存压力，flush之前行不可再修改,减少内存占用，建议每1000行刷新一次（若列很多当调整）

book.save('test.xls')
