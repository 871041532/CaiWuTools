import ExcelTool as excelTool

target_excel = excelTool.read_excel("3.合同应收-权责（按科目-含税）.xlsx")
source_excel = excelTool.read_excel("金地疫情期间减免租金数据收集及操作_2020.04.29_回稿2(1).xlsx")
target_sheet = target_excel.create_sheet("target_sheet")

# 第一次迭代选出(名字年月)对应的金额映射
name_date_dict = {}
date_str_set = set()
def iter_func1(iter_data):
    if iter_data.计费开始日期.year != iter_data.计费结束日期.year or iter_data.计费开始日期.month != iter_data.计费结束日期.month:
        raise Exception("开始、结束年或者月不一致！")

    if iter_data.品牌说明 not in name_date_dict:
        name_date_dict[iter_data.品牌说明] = {}
    dict = name_date_dict[iter_data.品牌说明]

    key = "%d.%02d"%(iter_data.计费开始日期.year, iter_data.计费开始日期.month)
    date_str_set.add(key)
    dict[key] = iter_data.减免总金额
excelTool.iter_sheets(iter_func1, source_excel.减免数据导入模板)

# 第二次迭代目标表写入数据
def iter_func2(iter_source, iter_target):
    if iter_target.row_index == 0:
        # 第一次的时候按顺序写入标题
        iter_target.序号 = None
        iter_target.商户品牌 = None
        date_str_list = list(date_str_set)
        date_str_list.sort()
        for date in date_str_list:
            iter_target[date] = None
    iter_target.序号 = iter_source.序号
    iter_target.商户品牌 = iter_source.商户品牌
    if iter_target.商户品牌 in name_date_dict:
        for date, amount in name_date_dict[iter_target.商户品牌].items():
            iter_target[date] = amount
excelTool.iter_sheets(iter_func2, target_excel.固定租金, target_sheet)

# 最终将target写入磁盘
excelTool.write_excel("目标.xlsx", target_excel)