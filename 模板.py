import ExcelTool as excelTool


# 创建一个新excel
excel_1 = excelTool.Excel()
sheet_1 = excel_1.create_sheet("sheet1")
# 写出这个新excel
excelTool.write_excel("结果.xlsx", excel_1)

# 获取桌面excel
excel_原1 = excelTool.read_excel("源表1名字.xlsx")
excel_原2 = excelTool.read_excel("源表2名字.xlsx")
# 得到表1原有sheet（固定租金）
sheet_原 =getattr(excel_原1, "固定租金")
# 创建表1新的sheet（求和）
sheet_新 = excel_原1.create_sheet("求和")

# 迭代两个sheet，从第二行开始迭代. 迭代更多sheet可以继续往后加 iter_data3
# 迭代一个sheet，，iter_data就是每次迭代时一行的数据
def iter_func(iter_data1, iter_data2):
    pass
excelTool.iter_sheets(iter_func, sheet_原, sheet_新)

# 迭代一个sheet，并进行处理
def iter_func1(iter_data):
    # 使用行数据
    name= iter_data["商户品牌"]    #或者iter_data.商户品牌
    # 修改这一行对应的某列数据，==用来判断是或否，如果是，则替换
    if iter_data["类别"] == "餐饮":
        iter_data["类别"] = "餐饮类"
    # 新加一个“统计”列
        iter_data["类别"] = 0
excelTool.iter_sheets(iter_func1, sheet_原)

# 将处理后的excel写入新的表，如果这个表已存在就覆盖掉，不存在就新建
excelTool.write_excel("结果sheet.xlsx", excel_原1)

# 迭代多个sheet，循环元组
title_tuple = ("2020.01", "2020.02","2020.03",  "2020.04", "2020.05","2020.06")
def iter_func2(iter_data1, iter_data2):
    for title in title_tuple:
        number1 = iter_data1[title] or 0
        number2 = iter_data2[title] or 0
        iter_data2[title+"求和"] = number1+number2
    pass
excelTool.iter_sheets(iter_func2, excel_原1, excel_原2)
# 迭代多个sheet，一个个加
def iter_func(iter_data1, iter_data2):
    iter_data2["品牌说明"]=iter_data1["品牌说明"]
    iter_data2["求和"]= iter_data1["应收金额"] + iter_data1["减免总金额"]
pass
excelTool.iter_sheets(iter_func, sheet_2, sheet_1)
excelTool.write_excel("结果.xlsx",excel_1)
