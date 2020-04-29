import ExcelTool as excelTool

# 从桌面获取excel1
excel1 = excelTool.read_excel("new.xlsx")
# 从excel1获取名字为页面1的sheet
sheet1 = excel1.页面1
# 创建名为页面2的sheet
excel1.create_sheet("页面3")
# 将sheet1中bgm列的值改为和name列一样
def iter_func(row_data):
    row_data.bgm = row_data.name
excelTool.iter_sheet(sheet1, iter_func)
# 将excel1写入硬盘
excelTool.write_excel("new.xlsx", excel1)

# 任务1：读取excel1，并写入到excel2.xlsx中
# 任务2：读取excel1.sheetA，并修改bgm列的值，写入excel2.xlsx中
# 任务3：读取excel1.sheetA，在excel1中创建sheetB，写入excel2.xlsx中
# 任务3：读取excel1.sheetA， excel1.sheetB，sheetB的“名字”列值等于sheetA的“花名”，写入excel2.xlsx中。
# 任务3：读取excel1.sheetA， 读取excel2并在excel2中创建sheetC，excel2.sheetC的“名字2”列值等于excel1.sheetA的“花名”，excel2写入excel2.xlsx中。