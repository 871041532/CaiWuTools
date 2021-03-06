from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
import os
import getpass
from collections import OrderedDict
from pyexcel_xlsx import get_data, save_data
from pyexcel_xls import get_data, save_data
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

        self.button1 = QPushButton("选择2019年-06成本费用计提明细表－上海莘宝")
        self.button1.clicked.connect(self.click_select_my)
        self.button1.setMinimumHeight(70)

        self.button2 = QPushButton("选择基础信息表")
        self.button2.clicked.connect(self.click_select_refer)
        self.button2.setMinimumHeight(70)

        self.button3 = QPushButton("开始处理")
        self.button3.clicked.connect(self.click_judge_btn)
        self.button3.setMinimumHeight(70)

        import datetime
        cur = datetime.datetime.now()
        year = cur.year
        month = cur.month
        time_str = str(year) + "/" + str(month)

        self.text_browser = QTextBrowser()
        mainLayout.addWidget(self.button1)
        mainLayout.addWidget(self.button2)
        mainLayout.addWidget(self.button3)
        mainLayout.addWidget(self.text_browser)

        # data
        self.my_filename = None #"C:\\Users\\pangjm\\Desktop\\1.xls" #""
        self.refer_filename = None #"C:\\Users\\pangjm\\Desktop\\部门.xlsx" #""
        self.out_put_filename = Globals.desktop_path + "成本费用既提引入表.xlsx"

        self.refer_data = None
        self.my_data = None

        self.setLayout(mainLayout)
        self.setWindowTitle("工资社保公积金")
        self.setGeometry(550, 150, 900, 800)
        self.show()

        if not Globals.user:
            self.close()

    def click_select_refer(self):
        info = QFileDialog.getOpenFileName(self, '选择基础信息表')
        if info and info[0]:
            self.refer_filename = info[0]
            self.log("选择了基础信息表：" + self.refer_filename)

    def click_select_my(self):
        info = QFileDialog.getOpenFileName(self, '选择源表')
        if info and info[0]:
            self.my_filename = info[0]
            self.log("选择了源表：" + self.my_filename)

    def clear_log(self):
        self.text_browser.clear()

    def click_judge_btn(self):
        self.clear_log()
        if not (self.my_filename):
            self.log("请先选择源表~(^v^)~", ShowWindow.LOG_ERROR)
        elif not (self.refer_filename):
            self.log("请先选择基础信息表~(^v^)~", ShowWindow.LOG_ERROR)
        else:
            self.load_my_data()
            if self.check_my_data():
                self.load_refer_data()
                if self.check_refer_data():
                    self.deal_data()
                    self.log("处理完毕。")

    def load_my_data(self):
        self.my_data = get_data(self.my_filename)

    def check_my_data(self):
        return True

    def load_refer_data(self):
        self.refer_data = get_data(self.refer_filename)

    def check_refer_data(self):
        if 0:
            self.log("一定要注意删掉禁用的自定义核算项目！",self.LOG_ERROR)
            return False
        else:
            self.log("基础表检测ok!")
            return True

    def deal_data(self):
        cur = datetime.datetime.now()
        year = cur.year
        month = cur.month
        # 获取demo
        excel_data = deepcopy(Globals.get_origin_excel_data())
        target_file_name = self.out_put_filename
        origin_data = []
        target_rows = []
        # 处理数据
        for i in range(len(origin_data)):
            one_data = origin_data[i]
            row = deepcopy(Globals.pingzheng_demo)
            row2 = deepcopy(Globals.pingzheng_demo)
            target_rows.append(row)
            target_rows.append(row2)
            row[Globals.get_pingzheng_idx("记账日期")] = data_str
            row2[Globals.get_pingzheng_idx("记账日期")] = data_str
            row[Globals.get_pingzheng_idx("业务日期")] = data_str
            row2[Globals.get_pingzheng_idx("业务日期")] = data_str
            row[Globals.get_pingzheng_idx("辅助账业务日期")] = data_str
            row2[Globals.get_pingzheng_idx("辅助账业务日期")] = data_str
            row[Globals.get_pingzheng_idx("会计期间")] = str(cur.month)
            row2[Globals.get_pingzheng_idx("会计期间")] = str(cur.month)
            row[Globals.get_pingzheng_idx("凭证类型")] = "转"
            row2[Globals.get_pingzheng_idx("凭证类型")] = "转"
            row[Globals.get_pingzheng_idx("凭证号")] = "20180600269"
            row2[Globals.get_pingzheng_idx("凭证号")] = "20180600269"
            row[Globals.get_pingzheng_idx("分录号")] = i * 2 + 1
            row2[Globals.get_pingzheng_idx("分录号")] = i * 2 + 2
            zhaiyao = ""
            row[Globals.get_pingzheng_idx("摘要")] = zhaiyao
            row2[Globals.get_pingzheng_idx("摘要")] = zhaiyao
            row[Globals.get_pingzheng_idx("科目")] = ""
            row2[Globals.get_pingzheng_idx("科目")] = ""
            row[Globals.get_pingzheng_idx("方向")] = 1
            row[Globals.get_pingzheng_idx("原币金额")] = 1
            row2[Globals.get_pingzheng_idx("原币金额")] = 1
            row2[Globals.get_pingzheng_idx("原币金额")] = -1
            row[Globals.get_pingzheng_idx("借方金额")] = 1
            row[Globals.get_pingzheng_idx("现金流量标记")] = 2
            row2[Globals.get_pingzheng_idx("现金流量标记")] = 2
            row[Globals.get_pingzheng_idx("辅助账摘要")] = zhaiyao
            row2[Globals.get_pingzheng_idx("辅助账摘要")] = zhaiyao
            row[Globals.get_pingzheng_idx("核算项目1")] = "长益租户"
            row[Globals.get_pingzheng_idx("名称1")] = ""
            row[Globals.get_pingzheng_idx("编码1")] = ""

        for row in target_rows:
            excel_data["凭证"].append(row)
        save_data(target_file_name, excel_data)
        Globals.eval_date_format(target_file_name, ["B", "C", "BN"])

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
