import ExcelTool as excelTool

# 获取桌面excel
excel_1 = excelTool.read_excel("2.合同押金及收入销账明细 - 副本.xlsx")
excel_2 = excelTool.read_excel("收入台账2020-4-30.xlsx")
sheet_1=excel_1.租金
sheet_2=excel_2.租金
#创建集合1，迭代 表1中的品牌说明添加到集合，统一转换为小写
jihe1=set()
def iter_func1(iter_data1):
    jihe1.add(iter_data1.商户品牌.lower())
    pass
excelTool.iter_sheets(iter_func1, sheet_1 ) #取自sheet_1
#创建集合2，迭代 表2中的商户品牌添加到集合，统一转换为小写
jihe2=set()
def iter_func2(iter_data2):
    jihe2.add(iter_data2.商户品牌.lower())
    pass
excelTool.iter_sheets(iter_func2, sheet_2 ) #取自sheet_2
#求差集，在集合1中却不在集合2里的东西
jihe3 =   jihe1 - jihe2
#如果差集有内容，打印差集内容，并报错
if len(jihe3) > 0:
    print(jihe3)
    raise  Exception("报错，有品牌写错")
else:
     print("品牌已全部包含，可以下一步")

