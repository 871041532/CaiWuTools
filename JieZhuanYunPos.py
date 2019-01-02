
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

        self.button1 = QPushButton("选择【凭证查询-高级】导出表")
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
        mainLayout.addWidget(self.button2)
        mainLayout.addWidget(self.label)
        mainLayout.addWidget(self.button3)

        mainLayout.addWidget(self.text_browser)

        # data
        self.my_filename = None  # "C:\\Users\\pangjm\\Desktop\\原始表.xlsx" #""
        self.out_put_filename2 = Globals.desktop_path + "收到云pos汇总表.xlsx"
        self.out_put_filename = Globals.desktop_path + "结转云Pos.xlsx"
        self.refer_filename = ""
        self.refer_data = None
        self.my_data = None

        self.setLayout(mainLayout)
        self.setWindowTitle("结转云Pos")
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
        info = QFileDialog.getOpenFileName(self, '选择【凭证查询-高级】导出表')
        if info and info[0]:
            self.my_filename = info[0]
            self.log("选择了【凭证查询-收字-高级】导出表：" + self.my_filename)

    def clear_log(self):
        self.text_browser.clear()

    def click_judge_btn(self):
        self.clear_log()
        if not (self.my_filename):
            self.log("请先选择【凭证查询-收字-高级】导出表~(^v^)~", ShowWindow.LOG_ERROR)
        elif not self.refer_filename:
            self.log("请先选择EAS编码基础信息表~(^v^)~", ShowWindow.LOG_ERROR)
        elif not self.label.text():
            self.log("请先输入凭证日期！", ShowWindow.LOG_ERROR)
        else:
            self.load_data()
            self.deal_data()
            self.log("处理完毕。")

    def load_data(self):
        self.log(self.my_filename.split('\\')[-1] + " 导入成功。")
        self.my_data = get_data(self.my_filename)["Sheet1"]
        title_row = self.my_data[0]
        kemu_idx = title_row.index("科目编码")
        kemu_value = "1122.01.01"
        zhaiyao_idx = title_row.index("摘要")
        del self.my_data[0]
        temp_data = [title_row]

        for row in self.my_data:
            if kemu_value == row[kemu_idx]:
                if "pos" in row[zhaiyao_idx] or "POS" in row[zhaiyao_idx] or "Pos" in row[zhaiyao_idx]:
                    temp_data.append(row)
        self.my_data =temp_data
        save_data(self.out_put_filename2, {"筛选收款":temp_data})


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
        replace_dict = self.get_replace_dict()
        # 银行凭证号
        pingzhenghao_bank_idx = bank_data[0].index("凭证号")
        # 原币金额索引
        yuanbijine_bank_idx = bank_data[0].index("原币金额")
        # 核算项目索引
        hesuanxiangmu_bank_idx = bank_data[0].index("核算项目")
        # 贷方金额
        daifangjine_bank_idx = bank_data[0].index("贷方")
        # 摘要
        zhaiyao_bank_idx = bank_data[0].index("摘要")

        for i in range(row_num):
            idx = i + 1
            row = []
            row2 = []
            target_rows.append(row)
            target_rows.append(row2)
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
            row.append("转")
            row2.append("转")
            # 凭证号
            row.append("20181100961")
            row2.append("20181100961")
            # 分录号
            row.append(idx * 2 - 1)
            row2.append(idx * 2)
            # 摘要
            replace_zhaiyao_str = "结转" + year + last_month
            zhaiyao_value = replace_zhaiyao_str + bank_data[idx][zhaiyao_bank_idx][2:]
            row.append(zhaiyao_value)
            row2.append(zhaiyao_value)
            # 科目
            row.append("1122.01.01")
            row2.append("6001.13")
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
            row.append(bank_data[idx][yuanbijine_bank_idx])
            row2.append(bank_data[idx][yuanbijine_bank_idx])
            # 数量
            row.append(0)
            row2.append(0)
            # 单价
            row.append(0)
            row2.append(0)
            # 借方金额
            row.append(bank_data[idx][daifangjine_bank_idx])
            row2.append("")
            # 贷方金额
            row.append("")
            row2.append(bank_data[idx][daifangjine_bank_idx])
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
            row.append("1544857876410--0")
            row2.append("1544857876410--0")
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
            row.append(2)
            row2.append(2)
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
            row.append("长益租户")
            row2.append("")
            # 编码1
            row.append("")
            row2.append("")
            # 名称1
            name1 = bank_data[idx][hesuanxiangmu_bank_idx].replace("长益租户:", "", 1)
            name1 = name1[:-1]
            row.append(name1)
            row2.append("")
            # 根据名称1修改编码1
            name_number_dict = replace_dict["name_number"]
            if name1 in name_number_dict:
                row[-2] = name_number_dict[name1]

            # 核算项目2
            row.append("")
            row2.append("")
            # 编码2
            row.append("")
            row2.append("")
            # 名称2
            row.append("")
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

        idx = 0
        for row in target_rows:
            if len(row) == len(target_data[0]):
                target_data.append(row)
            else:
                print(idx,len(row) - len(target_data[0]))
                idx = idx + 1

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
