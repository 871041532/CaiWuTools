import ExcelTool as excelTool
# 创建一个新excel，用来存放结果
excel_1 = excelTool.Excel()
sheet_1 = excel_1.create_sheet("sheet1")
# 获取桌面excel
excel_2 = excelTool.read_excel("目标.xlsx")
# 得到源表中原有sheet（固定租金）
sheet_2 =getattr(excel_2, "固定租金")
sheet_3 =getattr(excel_2, "target_sheet")
# 迭代多个sheet，循环元组
title_tuple = ("2020.01", "2020.02","2020.03",  "2020.04", "2020.05","2020.06")
def iter_func(iter_data1, iter_data2, iter_data3):
    iter_data3["序号"]=iter_data1["序号"]
    iter_data3["商户品牌"]=iter_data1["商户品牌"]
    for title in title_tuple:
        number1 = iter_data1[title] or 0
        number2 = iter_data2[title] or 0
        iter_data3[title+"求和"] = number1+number2
    pass
#迭代iter行的表（公式名，iter_data2取自sheet2,  iter_data3取自sheeet3， iter_data1取自sheet1
excelTool.iter_sheets(iter_func, sheet_2, sheet_3, sheet_1)
# 将处理后的excel写入新的表，如果这个表已存在就覆盖掉，不存在就新建
excelTool.write_excel("结果.xlsx",excel_1)

