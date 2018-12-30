# coding:utf8
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
import getpass
from collections import OrderedDict
from pyexcel_xlsx import get_data, save_data
from pyexcel_xls import get_data
import csv
import re
import copy
import datetime
from copy import deepcopy
from Globals import Globals

class ShowWindow(QWidget):
    LOG_WORLD = 1
    LOG_ERROR = 2

    def closeEvent(self, QCloseEvent):
        if hasattr(Globals, "MainWin"):
            Globals.current_win = Globals.MainWin()

    """docstring for ClassName"""

    def __init__(self):
        super(ShowWindow, self).__init__()
        self.initUI()

    def initUI(self):
        mainLayout = QVBoxLayout()

        self.button1 = QPushButton("选择古北当月银行流水")
        self.button1.clicked.connect(self.click_select_my)
        self.button1.setMinimumHeight(70)

        self.button2 = QPushButton("选择EAS编码基础信息表")
        self.button2.clicked.connect(self.click_select_refer)
        self.button2.setMinimumHeight(70)

        cur = datetime.datetime.now()
        month = cur.month - 1
        if month == 0:
            month = 12
        time_str = str(cur.year) + "/" + str(month) + "/28"
        self.label = QLineEdit(time_str)
        self.label.setPlaceholderText("输入凭证日期 如2018/11/6...")
        self.label.setMinimumHeight(50)

        self.button3 = QPushButton("开始处理")
        self.button3.clicked.connect(self.click_judge_btn)
        self.button3.setMinimumHeight(70)

        self.text_browser = QTextBrowser()

        mainLayout.addWidget(self.button1)
        # mainLayout.addWidget(self.button2)
        mainLayout.addWidget(self.label)
        mainLayout.addWidget(self.button3)

        mainLayout.addWidget(self.text_browser)

        # data
        self.my_filename = None  # "C:\\Users\\pangjm\\Desktop\\源.csv" #""
        # self.out_put_filename2 = "C:\\Users\\pangjm\\Desktop\\收到云pos汇总表.xlsx"
        self.out_put_filename = Globals.desktop_path + "停车场微信返款.xlsx"
        self.refer_filename = ""
        self.refer_data = None
        self.my_data = None

        self.setLayout(mainLayout)
        self.setWindowTitle("停车场微信返款")
        self.setGeometry(550, 150, 900, 800)
        self.show()

        if not Globals.user:
            self.close()

    def click_select_refer(self):
        info = QFileDialog.getOpenFileName(self, '选择EAS编码基础信息表')
        if info and info[0]:
            self.refer_filename = info[0]
            self.log("选择了" + self.refer_filename)

    def click_select_my(self):
        info = QFileDialog.getOpenFileName(self, '选择古北当月银行流水')
        if info and info[0]:
            self.my_filename = info[0]
            self.log("选择了古北当月银行流水：" + self.my_filename)

    def clear_log(self):
        self.text_browser.clear()

    def click_judge_btn(self):
        self.clear_log()
        if not (self.my_filename):
            self.log("请先选择古北当月银行流水~(^v^)~", ShowWindow.LOG_ERROR)
        #elif not self.refer_filename:
        #    self.log("请先选择EAS编码基础信息表~(^v^)~", ShowWindow.LOG_ERROR)
        elif not self.label.text():
            self.log("请先输入凭证日期！", ShowWindow.LOG_ERROR)
        else:
            self.load_data()
            self.deal_data()
            self.log("处理完毕。")

    def load_data(self):
        self.log(self.my_filename.split('\\')[-1] + " 导入成功。")
        csv_file = open(self.my_filename, 'r')
        reader = csv.reader(csv_file)
        data = []
        for line in reader:
            data.append(line)
        csv_file.close()
        del data[0]

        duifanganwei_idx = data[0].index("对方单位名称")
        target_data = data[0:1]
        for i in range(len(data) - 1):
            idx = i + 1
            row = data[idx]
            if row[duifanganwei_idx] == "财付通支付科技有限公司" or row[duifanganwei_idx] == "财付通支付科技有限公司客户备付金":
                target_data.append(row)
        self.my_data = target_data

    # target_data
    def deal_data(self):
        target_excel = {"凭证": []}
        target_data = target_excel["凭证"]
        row1 = [
            "公司",
            "记账日期",
            "业务日期",
            "会计期间",
            "凭证类型",
            "凭证号",
            "分录号",
            "摘要",
            "科目",
            "币种",
            "汇率",
            "方向",
            "原币金额",
            "数量",
            "单价",
            "借方金额",
            "贷方金额",
            "制单人",
            "过账人",
            "审核人",
            "附件数量",
            "过账标记",
            "机制凭证模块",
            "删除标记",
            "凭证序号",
            "单位",
            "参考信息",
            "是否有现金流量",
            "现金流量标记",
            "业务编号",
            "结算方式",
            "结算号",
            "辅助账摘要",
            "核算项目1",
            "编码1",
            "名称1",
            "核算项目2",
            "编码2",
            "名称2",
            "核算项目3",
            "编码3",
            "名称3",
            "核算项目4",
            "编码4",
            "名称4",
            "核算项目5",
            "编码5",
            "名称5",
            "核算项目6",
            "编码6",
            "名称6",
            "核算项目7",
            "编码7",
            "名称7",
            "核算项目8",
            "编码8",
            "名称8",
            "发票号",
            "换票证号",
            "客户",
            "费用类别",
            "收款人",
            "物料",
            "财务组织",
            "供应商",
            "辅助账业务日期",
            "到期日",
        ]
        target_data.append(row1)
        row2 = ["公司", "记账日期", "会计期间", "凭证类型", "凭证号", "币种", "分录号", "对方分录号", "主表信息", "附表信息", "原币", "本位币", "报告币", "主表金额系数",
                "附表金额系数", "性质", "核算项目1", "编码1", "名称1", "核算项目2", "编码2", "名称2", "核算项目3", "编码3", "名称3", "核算项目4", "编码4",
                "名称4", "核算项目5", "编码5", "名称5", "核算项目6", "编码6", "名称6", "核算项目7", "编码7", "名称7", "核算项目8", "编码8", "名称8"]
        target_excel["现金流量"] = [row2]
        refer_data = self.refer_data
        bank_data = self.my_data

        target_rows = []
        row_num = len(bank_data) - 1

        date_str = self.label.text()
        month = date_str.split('/')[1]
        last_month = int(month) - 2
        if last_month == 0:
            last_month = 12
        elif last_month == -1:
            last_month = 11
        if last_month < 10:
            last_month = "0" + str(last_month)
        else:
            last_month = str(last_month)

        year = date_str.split('/')[0]
        day = date_str.split('/')[2]
        # 替换信息
        # replace_dict = self.get_replace_dict()
        # 银行凭证号
        #pingzhenghao_bank_idx = bank_data[0].index("凭证号")
        # 原币金额索引
        #yuanbijine_bank_idx = bank_data[0].index("原币金额")
        # 核算项目索引
        #hesuanxiangmu_bank_idx = bank_data[0].index("核算项目")
        # 贷方金额
        daifangjine_bank_idx = bank_data[0].index("贷方发生额")
        # 摘要
        zhaiyao_bank_idx = bank_data[0].index("摘要")
        # 交易时间
        jiaoyitime_bank_idx = bank_data[0].index("交易时间")

        for i in range(row_num):
            idx = i + 1
            row = []
            row2 = []
            target_rows.append(row)
            # target_rows.append(row2)
            # 公司
            row.append("E018")
            row2.append("E018")
            # 记账日期
            row.append(date_str)
            row2.append(date_str)
            # 业务日期
            row.append(date_str)
            row2.append(date_str)
            # 会计期间
            row.append(int(month))
            row2.append(int(month))
            # 凭证类型
            row.append("收")
            row2.append("收")
            # 凭证号
            row.append("20181100963")
            row2.append("20181100963")
            # 分录号
            row.append(idx)
            row2.append(idx * 2)
            # 摘要
            jiaoyitime_str = bank_data[idx][jiaoyitime_bank_idx].split()[0]
            middle_str = "收到 微信一点停微信返款-"
            zhiayao_str = bank_data[idx][zhaiyao_bank_idx]
            zhaiyao_value = jiaoyitime_str + middle_str + zhiayao_str
            row.append(zhaiyao_value)
            row2.append("")
            # 科目
            row.append("1002.01")
            row2.append("1002.01")
            # 币种
            row.append("BB01")
            row2.append("BB01")
            # 汇率
            row.append(1)
            row2.append(1)
            # 方向
            row.append(1)
            row2.append(0)
            # 原币金额
            money_str = bank_data[idx][daifangjine_bank_idx].replace("," , "")
            money_str = money_str.replace(" ", "")
            curr_money = float(money_str) if money_str else 0
            row.append(curr_money)
            row2.append(curr_money)
            # sum = sum + int(curr_money * 100)
            # 数量
            row.append(0)
            row2.append(0)
            # 单价
            row.append(0)
            row2.append(0)
            # 借方金额
            row.append(curr_money)
            row2.append("")
            # 贷方金额
            row.append("")
            row2.append("")
            # 制单人
            row.append(Globals.eas_name)
            row2.append(Globals.eas_name)
            # 过账人
            row.append("")
            row2.append("")
            # 审核人
            row.append("")
            row2.append("")
            # 附件数量
            row.append(1)
            row2.append(1)
            # 过账标记
            row.append("TRUE")
            row2.append("TRUE")
            # 机制凭证
            row.append("")
            row2.append("")
            # 删除标记
            row.append("FALSE")
            row2.append("FALSE")
            # 凭证序号
            row.append("1544876875657--0")
            row2.append("1544876875657--0")
            # 单位
            row.append("")
            row2.append("")
            # 参考信息
            row.append("")
            row2.append("")
            # 是否有现金流量
            row.append("")
            row2.append("")
            # 现金流量标记
            row.append(6)
            row2.append(6)
            # 业务编号
            row.append("")
            row2.append("")
            # 结算方式
            row.append("")
            row2.append("")
            # 结算号
            row.append("")
            row2.append("")
            # 辅助账摘要
            row.append(zhaiyao_value)
            row2.append("")
            # 核算项目1
            row.append("银行账户")
            row2.append("")
            # 编码1
            row.append("001.E018")
            row2.append("")
            # 名称1
            row.append("工商银行古北新区支行1001282509300071274")
            row2.append("")
            # 核算项目2
            row.append("现金流量项目")
            row2.append("")
            # 编码2
            row.append("111")
            row2.append("")
            # 名称2
            row.append("出租业务所收到的现金")
            row2.append("")

            # 项目3
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            # 项目4
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            # 项目5
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            # 项目6
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            # 项目7
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            # 项目8
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            row.append("")
            row2.append("")
            # 发票号
            row.append("")
            row2.append("")
            # 换票证号
            row.append("")
            row2.append("")
            # 客户
            row.append("")
            row2.append("")
            # 费用类别
            row.append("")
            row2.append("")
            # 收款人
            row.append("")
            row2.append("")
            # 物料
            row.append("")
            row2.append("")
            # 财务组织
            row.append("")
            row2.append("")
            # 供应商
            row.append("")
            row2.append("")
            # 辅助帐业务日期
            row.append(date_str)
            row2.append("")
            # 到期日
            row.append("")
            row2.append("")

        for row in target_rows:
            target_data.append(row)

        # 处理结算的
        last_row = deepcopy(target_data[-1])
        title_row = target_data[0]
        jine_idx = title_row.index("原币金额")
        sum = 0
        for i in range(row_num):
            idx = i + 1
            row = target_data[idx]
            sum = sum + float(row[jine_idx])

        last_row[title_row.index("摘要")] = "收到 上海七宝万科停车场停车费-微信返款"
        last_row[title_row.index("科目")] = "1122.99"
        last_row[title_row.index("方向")] = "0"
        last_row[title_row.index("原币金额")] = float("%.2f" % (sum))
        last_row[title_row.index("借方金额")] = ""
        last_row[title_row.index("贷方金额")] = float("%.2f" % (sum))
        last_row[title_row.index("辅助账摘要")] = "收到 上海七宝万科停车场停车费-微信返款"
        last_row[title_row.index("核算项目1")] = "客户"
        last_row[title_row.index("编码1")] = "0503.E018"
        last_row[title_row.index("分录号")] = len(target_data)
        idx = title_row.index("名称1")
        last_row[idx] = "微信一点停"
        last_row[idx + 1] = ""
        last_row[idx + 2] = ""
        last_row[idx + 3] = ""
        target_data.append(last_row)

        save_data(self.out_put_filename, target_excel)
        # 处理日期
        Globals.eval_date_format(self.out_put_filename, ["B", "C", "BN"])

    #
    def get_replace_dict(self):
        return_data = {}
        sheet = get_data(self.refer_filename)["EAS基础信息"]
        mingcheng_idx = sheet[0].index("名称")
        bianma_idx = sheet[0].index("编码")
        sheet = sheet[1:]

        #名称查找编码
        name_to_number = {}
        return_data["name_number"] = name_to_number
        for row in sheet:
            name_to_number[row[mingcheng_idx]] = row[bianma_idx]
        return return_data

    def log(self, strs, enum=1):
        if enum == self.LOG_ERROR:
            pre_str = """<font face="微软雅黑" size="6" color="red">%s</font>"""
            self.text_browser.append(pre_str % str(strs))
        else:
            pre_str = """<font face="微软雅黑" size="4" color="black">%s</font>"""
            self.text_browser.append(pre_str % str(strs))


def into():
    return ShowWindow()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ShowWindow()
    sys.exit(app.exec_())
