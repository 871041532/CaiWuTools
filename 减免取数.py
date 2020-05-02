import ExcelTool as excelTool
# 创建一个新excel，用来存放结果
excel_1 = excelTool.Excel()
sheet_1 = excel_1.create_sheet("减免")
sheet_2 = excel_1.create_sheet("减免后")
# 获取桌面excel
target_excel = excelTool.read_excel("3.合同应收-权责（按科目-含税）.xlsx")
source_excel = excelTool.read_excel("金地疫情期间减免租金数据收集及操作_2020.04.30.xlsx")

# 第一次迭代选出(名字年月)对应的金额映射
name_date_dict = {}
date_str_set = set()
def iter_func1(iter_data):
    # iter_data.品牌说明 = iter_data.品牌说明.lower()
    if iter_data.计费开始日期.year != iter_data.计费结束日期.year or iter_data.计费开始日期.month != iter_data.计费结束日期.month:
        raise Exception("开始、结束年或者月不一致！")

    if iter_data.品牌说明 not in name_date_dict:
        name_date_dict[iter_data.品牌说明] = {}

    dict = name_date_dict[iter_data.品牌说明]

    key = "%d.%02d"%(iter_data.计费开始日期.year, iter_data.计费开始日期.month)
    date_str_set.add(key)
    dict[key] = iter_data.减免总金额
excelTool.iter_sheets(iter_func1, source_excel.减免数据导入模板)

# 第二次迭代目标表写入减免金额数据
def iter_func2(iter_source, iter_target):
    if iter_target.row_index == 0:
        # 第一次的时候按顺序写入标题
        iter_target.序号 = None
        iter_target.商户品牌 = None
        date_str_list = list(date_str_set)
        date_str_list.sort()
        for date in date_str_list:
            iter_target[date] = None

    # iter_source.商户品牌 = iter_source.商户品牌.lower()
    iter_target.序号 = iter_source.序号
    iter_target.商户品牌 = iter_source.商户品牌
    if iter_target.商户品牌 in name_date_dict:
        for date, amount in name_date_dict[iter_target.商户品牌].items():
            iter_target[date] = amount
excelTool.iter_sheets(iter_func2, target_excel.固定租金, sheet_1)

# 第三次迭代写入固定租金+减免金额数据
title_tuple = ("2020.01", "2020.02","2020.03",  "2020.04", "2020.05","2020.06")
def iter_func3(iter_data1, iter_data2, iter_data3):
    iter_data3.序号=iter_data1.序号
    iter_data3.商户品牌=iter_data1.商户品牌
    for title in title_tuple:
        number1 = iter_data1[title] or 0
        number2 = iter_data2[title] or 0
        iter_data3[title+"求和"] = number1+number2
    pass
#迭代iter行的表（公式名，iter_data2取自sheet2,  iter_data3取自sheeet3， iter_data1取自sheet1
excelTool.iter_sheets(iter_func3, target_excel.固定租金, sheet_1, sheet_2)
# 最终将target写入磁盘
excelTool.write_excel("结果.xlsx", excel_1)