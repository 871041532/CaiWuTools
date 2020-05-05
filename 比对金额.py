import ExcelTool as excelTool

# 获取桌面excel
excel_1 = excelTool.read_excel("2.合同押金及收入销账明细 - 副本.xlsx")
excel_2 = excelTool.read_excel("收入台账2020-4-30标题.xlsx")
sheet_1=excel_1.租金
sheet_2=excel_2.Sheet1

#创建集合1，迭代 表1中的商户品牌添加到集合，创建字典，迭代的每一行放入字典
jihe1=set()
zidian1 = {}
def iter_func1(iter_data1):
    if iter_data1.商户品牌 != "":  # 判断不是空行再进行下一步
        if iter_data1.商户品牌 in jihe1:
            raise Exception("表1里面有重复key，" + iter_data1.商户品牌)
        else:
            jihe1.add(iter_data1.商户品牌)
            zidian1[iter_data1.商户品牌] = iter_data1
    pass
excelTool.iter_sheets(iter_func1, sheet_1 )   #取自sheet_1

#创建集合2，迭代 表2中的商户品牌添加到集合
jihe2=set()
def iter_func2(iter_data1):
    jihe2.add(iter_data1.商户品牌)
    pass
excelTool.iter_sheets(iter_func2, sheet_2)   #取自sheet_2

#先检查两表索引是否一致，求差集，如果有不一致的，打印差异内容，并停止
jihe3 =   jihe1 - jihe2
if len(jihe3) > 0:
    print("差异为表1多了" + str(jihe3))
    raise Exception("报错，表1中有，表2没有")
else:
    print("表1key已全部在表2中")
jihe4 =   jihe2 - jihe1
if len(jihe4) > 0:
    print("差异为表2多了"+str(jihe4))
    # raise Exception("报错，表2中有，表1没有")
else:
     print("表2key已全部在表1中")

# 迭代表2内容，表1根据key定位value，两者比对
def iter_func3(iter_data1):
    iter_data2 = zidian1[iter_data1.商户品牌]  #字典的KEY

    a1 = iter_data1.应收固租 or 0                 #迭代表2行，应收租金列
    a2 = iter_data2["202006应收固租"] or 0  # 字典的202006应收固租列
    if a1 != a2:   #不等于
        s = "202006应收固租差异：%s %s 和我 %s "%(iter_data1.商户品牌, a1, a2)
        print(s)
    # print("报错：" + iter_data1.商户品牌 + str(a1) +"和" + str(a2) + "不同")

    b1= iter_data1.疫情减免 or 0               #迭代表2行，疫情减免列
    b2 = iter_data2["202006固租减免"] or 0   # 字典的202006固租减免列
    if b1 != b2:  # 不等于
        s ="202006固租减免差异：%s %s 和我 %s "%(iter_data1.商户品牌, b1, b2)
        print(s)

    c1= iter_data1.应收物业 or 0
    c2 = iter_data2["202006应收物业"] or 0
    if c1 != c2:
        s ="202006应收物业差异：%s %s 和我 %s "%(iter_data1.商户品牌, c1, c2)
        print(s)

    d1= iter_data1.应收推广 or 0
    d2 = iter_data2["202006应收推广"] or 0
    if d1 != d2:
        s ="202006应收推广差异：%s %s 和我 %s "%(iter_data1.商户品牌, d1, d2)
        print(s)

    e1= iter_data1.已收固租 or 0
    e2 = iter_data2["已收固租"] or 0
    if e1 != e2:
        s ="已收固租 差异：%s %s 和我 %s "%(iter_data1.商户品牌, e1, e2)
        print(s)

    f1= iter_data1.已收物业 or 0
    f2 = iter_data2["已收物业"] or 0
    if f1 != f2:
        s ="已收物业 差异：%s %s 和我 %s "%(iter_data1.商户品牌, f1, f2)
        print(s)

    g1= iter_data1.已收推广 or 0
    g2 = iter_data2["已收推广"] or 0
    if g1 != g2:
        s ="已收推广 差异：%s %s 和我 %s "%(iter_data1.商户品牌, g1, g2)
        print(s)



excelTool.iter_sheets(iter_func3, sheet_2)  # 取自sheet_2
