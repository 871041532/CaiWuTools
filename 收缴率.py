# 收缴台账的表里先把校验的五个数去掉，再把最下面空白行去掉

import ExcelTool as excelTool
# 读取桌面excel
excel_1 = excelTool.read_excel("2.合同押金及收入销账明细.xlsx")
excel_2 = excelTool.Excel()
sheet_1 = excel_1.租金
sheet_1.SetLower(False)
sheet_2 = excel_2.create_sheet("201912")
sheet_2.SetLower(False)
sheet_3 = excel_2.create_sheet("202001")
sheet_3.SetLower(False)
sheet_4 = excel_2.create_sheet("202002")
sheet_4.SetLower(False)
sheet_5 = excel_2.create_sheet("202003")
sheet_5.SetLower(False)
sheet_6 = excel_2.create_sheet("202004")
sheet_6.SetLower(False)
sheet_7 = excel_2.create_sheet("202005")
sheet_7.SetLower(False)

#201912收缴：第一次迭代按顺序写入基本信息
def iter_func(iter_data1, iter_data2):
    iter_data2["序号"]=iter_data1["序号"]
    iter_data2["合同编号"] = iter_data1["合同编号"]
    iter_data2["铺位号"] = iter_data1["铺位号"]
    iter_data2["面积"] = iter_data1["面积"]
    iter_data2["商户品牌"] = iter_data1["商户品牌"]
    iter_data2["承租方"] = iter_data1["承租方"]
excelTool.iter_sheets(iter_func, sheet_1, sheet_2 )

def iter_func2(iter_data1,  iter_data2):
    number1 = iter_data1["202001应收固租"] or 0
    number2 = iter_data1["202001固租减免"] or 0
    number3 = iter_data1["202001已收固租"] or 0

    number4 = iter_data1["202001应收物业"] or 0
    number5 = iter_data1["202001已收物业"] or 0

    number6= iter_data1["202001应收推广"] or 0
    number7 = iter_data1["202001已收推广"] or 0

    number8= iter_data1["201911应收取高"] or 0
    number9 = iter_data1["201911已收取高"] or 0

    number10= iter_data1["201911应收提成"] or 0
    number11 = iter_data1["201911已收提成"] or 0

    iter_data2["201912现金流欠费合计"] = number1 + number2 - number3+number4 - number5+number6 - number7+number8 - number9+number10 - number11
    iter_data2["202001欠费固租"] = number1 + number2 - number3
    iter_data2["202001欠费物业"] = number4 - number5
    iter_data2["202001欠费推广"] = number6 - number7
    iter_data2["201911欠费取高"] = number8 - number9
    iter_data2["201911欠费提成"] = number10 - number11

    iter_data2["202001应收固租"] = iter_data1["202001应收固租"]
    iter_data2["202001固租减免"] = iter_data1["202001固租减免"]
    iter_data2["202001应收物业"] = iter_data1["202001应收物业"]
    iter_data2["202001应收推广"] = iter_data1["202001应收推广"]
    iter_data2["201911应收取高"] = iter_data1["201911应收取高"]
    iter_data2["201911应收提成"] = iter_data1["201911应收提成"]
    iter_data2["202001已收固租"] = iter_data1["202001已收固租"]
    iter_data2["202001已收物业"] = iter_data1["202001已收物业"]
    iter_data2["202001已收推广"] = iter_data1["202001已收推广"]
    iter_data2["201911已收取高"] = iter_data1["201911已收取高"]
    iter_data2["201911已收提成"] = iter_data1["201911已收提成"]
excelTool.iter_sheets(iter_func2, sheet_1, sheet_2)

#202001收缴：第一次迭代按顺序写入基本信息
def iter_func(iter_data1, iter_data2):
    iter_data2["序号"]=iter_data1["序号"]
    iter_data2["合同编号"] = iter_data1["合同编号"]
    iter_data2["铺位号"] = iter_data1["铺位号"]
    iter_data2["面积"] = iter_data1["面积"]
    iter_data2["商户品牌"] = iter_data1["商户品牌"]
    iter_data2["承租方"] = iter_data1["承租方"]
excelTool.iter_sheets(iter_func, sheet_1, sheet_3 )

def iter_func2(iter_data1,  iter_data2):
    number1 = iter_data1["202002应收固租"] or 0
    number2 = iter_data1["202002固租减免"] or 0
    number3 = iter_data1["202002已收固租"] or 0

    number4 = iter_data1["202002应收物业"] or 0
    number5 = iter_data1["202002已收物业"] or 0

    number6= iter_data1["202002应收推广"] or 0
    number7 = iter_data1["202002已收推广"] or 0

    number8= iter_data1["201912应收取高"] or 0
    number9 = iter_data1["201912已收取高"] or 0

    number10= iter_data1["201912应收提成"] or 0
    number11 = iter_data1["201912已收提成"] or 0

    iter_data2["202001现金流欠费合计"] = number1 + number2 - number3+number4 - number5+number6 - number7+number8 - number9+number10 - number11
    iter_data2["202002欠费固租"] = number1 + number2 - number3
    iter_data2["202002欠费物业"] = number4 - number5
    iter_data2["202002欠费推广"] = number6 - number7
    iter_data2["201912欠费取高"] = number8 - number9
    iter_data2["201912欠费提成"] = number10 - number11

    iter_data2["202002应收固租"] = iter_data1["202002应收固租"]
    iter_data2["202002固租减免"] = iter_data1["202002固租减免"]
    iter_data2["202002应收物业"] = iter_data1["202002应收物业"]
    iter_data2["202002应收推广"] = iter_data1["202002应收推广"]
    iter_data2["201912应收取高"] = iter_data1["201912应收取高"]
    iter_data2["201912应收提成"] = iter_data1["201912应收提成"]
    iter_data2["202002已收固租"] = iter_data1["202002已收固租"]
    iter_data2["202002已收物业"] = iter_data1["202002已收物业"]
    iter_data2["202002已收推广"] = iter_data1["202002已收推广"]
    iter_data2["201912已收取高"] = iter_data1["201912已收取高"]
    iter_data2["201912已收提成"] = iter_data1["201912已收提成"]
excelTool.iter_sheets(iter_func2, sheet_1, sheet_3)

#202002收缴：第一次迭代按顺序写入基本信息
def iter_func(iter_data1, iter_data2):
    iter_data2["序号"]=iter_data1["序号"]
    iter_data2["合同编号"] = iter_data1["合同编号"]
    iter_data2["铺位号"] = iter_data1["铺位号"]
    iter_data2["面积"] = iter_data1["面积"]
    iter_data2["商户品牌"] = iter_data1["商户品牌"]
    iter_data2["承租方"] = iter_data1["承租方"]
excelTool.iter_sheets(iter_func, sheet_1, sheet_4 )

def iter_func2(iter_data1,  iter_data2):
    number1 = iter_data1["202003应收固租"] or 0
    number2 = iter_data1["202003固租减免"] or 0
    number3 = iter_data1["202003已收固租"] or 0

    number4 = iter_data1["202003应收物业"] or 0
    number5 = iter_data1["202003已收物业"] or 0

    number6= iter_data1["202003应收推广"] or 0
    number7 = iter_data1["202003已收推广"] or 0

    number8= iter_data1["202001应收取高"] or 0
    number9 = iter_data1["202001已收取高"] or 0

    number10= iter_data1["202001应收提成"] or 0
    number11 = iter_data1["202001已收提成"] or 0

    iter_data2["202002现金流欠费合计"] = number1 + number2 - number3+number4 - number5+number6 - number7+number8 - number9+number10 - number11
    iter_data2["202003欠费固租"] = number1 + number2 - number3
    iter_data2["202003欠费物业"] = number4 - number5
    iter_data2["202003欠费推广"] = number6 - number7
    iter_data2["202001欠费取高"] = number8 - number9
    iter_data2["202001欠费提成"] = number10 - number11

    iter_data2["202003应收固租"] = iter_data1["202003应收固租"]
    iter_data2["202003固租减免"] = iter_data1["202003固租减免"]
    iter_data2["202003应收物业"] = iter_data1["202003应收物业"]
    iter_data2["202003应收推广"] = iter_data1["202003应收推广"]
    iter_data2["202001应收取高"] = iter_data1["202001应收取高"]
    iter_data2["202001应收提成"] = iter_data1["202001应收提成"]
    iter_data2["202003已收固租"] = iter_data1["202003已收固租"]
    iter_data2["202003已收物业"] = iter_data1["202003已收物业"]
    iter_data2["202003已收推广"] = iter_data1["202003已收推广"]
    iter_data2["202001已收取高"] = iter_data1["202001已收取高"]
    iter_data2["202001已收提成"] = iter_data1["202001已收提成"]
excelTool.iter_sheets(iter_func2, sheet_1, sheet_4)

#202003收缴：第一次迭代按顺序写入基本信息
def iter_func(iter_data1, iter_data2):
    iter_data2["序号"]=iter_data1["序号"]
    iter_data2["合同编号"] = iter_data1["合同编号"]
    iter_data2["铺位号"] = iter_data1["铺位号"]
    iter_data2["面积"] = iter_data1["面积"]
    iter_data2["商户品牌"] = iter_data1["商户品牌"]
    iter_data2["承租方"] = iter_data1["承租方"]
excelTool.iter_sheets(iter_func, sheet_1, sheet_5 )

def iter_func2(iter_data1,  iter_data2):
    number1 = iter_data1["202004应收固租"] or 0
    number2 = iter_data1["202004固租减免"] or 0
    number3 = iter_data1["202004已收固租"] or 0

    number4 = iter_data1["202004应收物业"] or 0
    number5 = iter_data1["202004已收物业"] or 0

    number6= iter_data1["202004应收推广"] or 0
    number7 = iter_data1["202004已收推广"] or 0

    number8= iter_data1["202002应收取高"] or 0
    number9 = iter_data1["202002已收取高"] or 0

    number10= iter_data1["202002应收提成"] or 0
    number11 = iter_data1["202002已收提成"] or 0

    iter_data2["202003现金流欠费合计"] = number1 + number2 - number3+number4 - number5+number6 - number7+number8 - number9+number10 - number11
    iter_data2["202004欠费固租"] = number1 + number2 - number3
    iter_data2["202004欠费物业"] = number4 - number5
    iter_data2["202004欠费推广"] = number6 - number7
    iter_data2["202002欠费取高"] = number8 - number9
    iter_data2["202002欠费提成"] = number10 - number11

    iter_data2["202004应收固租"] = iter_data1["202004应收固租"]
    iter_data2["202004固租减免"] = iter_data1["202004固租减免"]
    iter_data2["202004应收物业"] = iter_data1["202004应收物业"]
    iter_data2["202004应收推广"] = iter_data1["202004应收推广"]
    iter_data2["202002应收取高"] = iter_data1["202002应收取高"]
    iter_data2["202002应收提成"] = iter_data1["202002应收提成"]
    iter_data2["202004已收固租"] = iter_data1["202004已收固租"]
    iter_data2["202004已收物业"] = iter_data1["202004已收物业"]
    iter_data2["202004已收推广"] = iter_data1["202004已收推广"]
    iter_data2["202002已收取高"] = iter_data1["202002已收取高"]
    iter_data2["202002已收提成"] = iter_data1["202002已收提成"]
excelTool.iter_sheets(iter_func2, sheet_1, sheet_5)

#202004收缴：第一次迭代按顺序写入基本信息
def iter_func(iter_data1, iter_data2):
    iter_data2["序号"]=iter_data1["序号"]
    iter_data2["合同编号"] = iter_data1["合同编号"]
    iter_data2["铺位号"] = iter_data1["铺位号"]
    iter_data2["面积"] = iter_data1["面积"]
    iter_data2["商户品牌"] = iter_data1["商户品牌"]
    iter_data2["承租方"] = iter_data1["承租方"]
excelTool.iter_sheets(iter_func, sheet_1, sheet_6 )

def iter_func2(iter_data1,  iter_data2):
    number1 = iter_data1["202005应收固租"] or 0
    number2 = iter_data1["202005固租减免"] or 0
    number3 = iter_data1["202005已收固租"] or 0

    number4 = iter_data1["202005应收物业"] or 0
    number5 = iter_data1["202005已收物业"] or 0

    number6= iter_data1["202005应收推广"] or 0
    number7 = iter_data1["202005已收推广"] or 0

    number8= iter_data1["202003应收取高"] or 0
    number9 = iter_data1["202003已收取高"] or 0

    number10= iter_data1["202003应收提成"] or 0
    number11 = iter_data1["202003已收提成"] or 0

    iter_data2["202004现金流欠费合计"] = number1 + number2 - number3+number4 - number5+number6 - number7+number8 - number9+number10 - number11
    iter_data2["202005欠费固租"] = number1 + number2 - number3
    iter_data2["202005欠费物业"] = number4 - number5
    iter_data2["202005欠费推广"] = number6 - number7
    iter_data2["202003欠费取高"] = number8 - number9
    iter_data2["202003欠费提成"] = number10 - number11

    iter_data2["202005应收固租"] = iter_data1["202005应收固租"]
    iter_data2["202005固租减免"] = iter_data1["202005固租减免"]
    iter_data2["202005应收物业"] = iter_data1["202005应收物业"]
    iter_data2["202005应收推广"] = iter_data1["202005应收推广"]
    iter_data2["202003应收取高"] = iter_data1["202003应收取高"]
    iter_data2["202003应收提成"] = iter_data1["202003应收提成"]
    iter_data2["202005已收固租"] = iter_data1["202005已收固租"]
    iter_data2["202005已收物业"] = iter_data1["202005已收物业"]
    iter_data2["202005已收推广"] = iter_data1["202005已收推广"]
    iter_data2["202003已收取高"] = iter_data1["202003已收取高"]
    iter_data2["202003已收提成"] = iter_data1["202003已收提成"]
excelTool.iter_sheets(iter_func2, sheet_1, sheet_6)

#202005收缴：第一次迭代按顺序写入基本信息
def iter_func(iter_data1, iter_data2):
    iter_data2["序号"]=iter_data1["序号"]
    iter_data2["合同编号"] = iter_data1["合同编号"]
    iter_data2["铺位号"] = iter_data1["铺位号"]
    iter_data2["面积"] = iter_data1["面积"]
    iter_data2["商户品牌"] = iter_data1["商户品牌"]
    iter_data2["承租方"] = iter_data1["承租方"]
excelTool.iter_sheets(iter_func, sheet_1, sheet_7 )

def iter_func2(iter_data1,  iter_data2):
    number1 = iter_data1["202006应收固租"] or 0
    number2 = iter_data1["202006固租减免"] or 0
    number3 = iter_data1["202006已收固租"] or 0

    number4 = iter_data1["202006应收物业"] or 0
    number5 = iter_data1["202006已收物业"] or 0

    number6= iter_data1["202006应收推广"] or 0
    number7 = iter_data1["202006已收推广"] or 0

    number8= iter_data1["202004应收取高"] or 0
    number9 = iter_data1["202004已收取高"] or 0

    number10= iter_data1["202004应收提成"] or 0
    number11 = iter_data1["202004已收提成"] or 0

    iter_data2["202005现金流欠费合计"] = number1 + number2 - number3+number4 - number5+number6 - number7+number8 - number9+number10 - number11
    iter_data2["202006欠费固租"] = number1 + number2 - number3
    iter_data2["202006欠费物业"] = number4 - number5
    iter_data2["202006欠费推广"] = number6 - number7
    iter_data2["202004欠费取高"] = number8 - number9
    iter_data2["202004欠费提成"] = number10 - number11

    iter_data2["202006应收固租"] = iter_data1["202006应收固租"]
    iter_data2["202006固租减免"] = iter_data1["202006固租减免"]
    iter_data2["202006应收物业"] = iter_data1["202006应收物业"]
    iter_data2["202006应收推广"] = iter_data1["202006应收推广"]
    iter_data2["202004应收取高"] = iter_data1["202004应收取高"]
    iter_data2["202004应收提成"] = iter_data1["202004应收提成"]
    iter_data2["202006已收固租"] = iter_data1["202006已收固租"]
    iter_data2["202006已收物业"] = iter_data1["202006已收物业"]
    iter_data2["202006已收推广"] = iter_data1["202006已收推广"]
    iter_data2["202004已收取高"] = iter_data1["202004已收取高"]
    iter_data2["202004已收提成"] = iter_data1["202004已收提成"]
excelTool.iter_sheets(iter_func2, sheet_1, sheet_7)

# sheet中加入排序，每次写入之前排序
sheet_2.Sort("铺位号")
sheet_3.Sort("铺位号")
sheet_4.Sort("铺位号")
sheet_5.Sort("铺位号")
sheet_6.Sort("铺位号")
sheet_7.Sort("铺位号")

# 排序之后加入序号，铺位号为空，直接忽略
class Idx(object):
    def __init__(self):
        self.idx = 1
    def increase(self):
        self.idx = self.idx + 1
idx_object = Idx()

def iter_funcx(i2, i3, i4, i5, i6, i7):
    if i2.铺位号 != "":
        i2.序号 = idx_object.idx
        i3.序号 = idx_object.idx
        i4.序号 = idx_object.idx
        i5.序号 = idx_object.idx
        i6.序号 = idx_object.idx
        i7.序号 = idx_object.idx
        idx_object.increase()

excelTool.iter_sheets(iter_funcx, sheet_2, sheet_3, sheet_4, sheet_5, sheet_6, sheet_7)
# 将处理后的excel2写入名字为收缴率的表中
excelTool.write_excel("收缴率.xlsx", excel_2)