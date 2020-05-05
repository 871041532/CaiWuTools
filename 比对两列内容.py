import ExcelTool as excelTool

# 获取桌面excel
excel_1 = excelTool.read_excel("2.合同押金及收入销账明细 - 副本.xlsx")
excel_2 = excelTool.read_excel("收入台账2020-4-30标题.xlsx")
sheet_1=excel_1.租金
sheet_2=excel_2.Sheet1
#创建集合1，迭代 表1中的品牌说明添加到集合
jihe1=set()
def iter_func1(iter_data1):
    if iter_data1.商户品牌 != "":                 #判断不是空行再进行下一步
        if iter_data1.商户品牌 in jihe1:          #判断集合1里面有没有重复的KEY
            raise Exception("表2里面有重复key，" + iter_data1.商户品牌)
        else:
            jihe1.add(iter_data1.商户品牌)
excelTool.iter_sheets(iter_func1, sheet_1 )      #迭代取自sheet_1
#创建集合2，迭代 表2中的商户品牌添加到集合
jihe2=set()
def iter_func2(iter_data1):
    if iter_data1.商户品牌 != "":
        if iter_data1.商户品牌 in jihe2:
            raise Exception("表2里面有重复key，" + iter_data1.商户品牌  )
        else:
            jihe2.add(iter_data1.商户品牌)
excelTool.iter_sheets(iter_func2, sheet_2 )      #迭代取自sheet_2
#求差集，在集合1中却不在集合2里的东西，如果差集有内容，打印差集内容
jihe3 =   jihe1 - jihe2
if len(jihe3) > 0:
    print("差异为表1多了"+str(jihe3))
    # raise Exception("报错，表1中有，表2没有")
else:
     print("表1key已全部在表2中")
#求差集，在集合2中却不在集合1里的东西，如果差集有内容，打印差集内容
jihe4 =   jihe2 - jihe1
if len(jihe4) > 0:
    print("差异为表2多了"+str(jihe4))
    # raise Exception("报错，表2中有，表1没有")
else:
     print("表2key已全部在表1中")

