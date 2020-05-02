import ExcelTool as excelTool
# 创建一个新excel，用来存放结果
excel_1 = excelTool.Excel()
sheet_1 = excel_1.create_sheet("sheet1")
# 获取桌面excel
excel_2 = excelTool.read_excel("金地疫情期间减免租金数据收集及操作_2020.04.29_回稿2(1).xlsx")
# 得到源表中原有sheet（固定租金）
sheet_2 =getattr(excel_2, "减免数据导入模板")
# 迭代多个sheet，循环元组
def iter_func(iter_data1, iter_data2):
    iter_data2["品牌说明"]=iter_data1["品牌说明"]
    iter_data2["求和"]= iter_data1["应收金额"] + iter_data1["减免总金额"]
pass
#迭代iter行的表（公式名，iter_data3取自sheet2,  iter_data3取自sheeet3， iter_data1取自sheet1
excelTool.iter_sheets(iter_func, sheet_2, sheet_1)
# 将处理后的excel写入新的表，如果这个表已存在就覆盖掉，不存在就新建
excelTool.write_excel("结果.xlsx",excel_1)

