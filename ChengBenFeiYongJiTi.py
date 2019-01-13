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
import datetime

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
        self.my_filename = Globals.desktop_path + "2019年-06成本费用计提明细表－上海莘宝(2).xlsx"
        self.refer_filename = Globals.desktop_path + "莘宝结转基础信息(2).xlsx"
        self.out_put_filename = Globals.desktop_path + "成本费用既提引入表.xlsx"

        self.refer_data = None
        self.my_data = None

        self.setLayout(mainLayout)
        self.setWindowTitle("成本费用计提")
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
        self.my_data = get_data(self.my_filename).popitem(last=False)[1][3:]

    def check_my_data(self):
        return True

    def load_refer_data(self):
        refer_data = get_data(self.refer_filename)["费用计提"]
        all_dict = {"部门":{}, "科目名称":{}, "供应商全称":{}}
        # 供应商全称
        for i in range(1, len(refer_data)):
            if refer_data[i] and len(refer_data[i]) >= 2 and refer_data[i][1]:
                if refer_data[i][1] not in all_dict["供应商全称"]:
                    all_dict["供应商全称"][refer_data[i][1]] = {"编码":refer_data[i][0]}
                else:
                    self.log("有重复的：" + str(refer_data[i][1]), self.LOG_ERROR)
        # 部门
        for i in range(1, len(refer_data)):
            if refer_data[i] and len(refer_data[i]) >= 7 and refer_data[i][3]:
                if refer_data[i][3] not in all_dict["部门"]:
                    all_dict["部门"][refer_data[i][3]] = {"编码1":refer_data[i][4], "名称1":refer_data[i][5], "科目":refer_data[i][6], }
                else:
                    self.log("有重复的：" + str(refer_data[i][3]), self.LOG_ERROR)
        # 科目名称
        for i in range(1, len(refer_data)):
            if refer_data[i] and len(refer_data[i]) >= 10 and refer_data[i][9]:
                if refer_data[i][9] not in all_dict["科目名称"]:
                    all_dict["科目名称"][refer_data[i][9]] = {"科目":refer_data[i][8]}
                else:
                    self.log("有重复的：" + str(refer_data[i][9]), self.LOG_ERROR)
        self.refer_data = all_dict

    def check_refer_data(self):
        bumen_idx = self.my_data[0].index("部门")
        kemumingcheng_idx = self.my_data[0].index("科目名称")
        gongyingshang_idx = self.my_data[0].index("供应商全称")

        is_check_ok = True
        for i in range(1, len(self.my_data)):
            if self.my_data[i][bumen_idx] and self.my_data[i][bumen_idx] not in self.refer_data["部门"]:
                self.log("缺少部门:" + self.my_data[i][bumen_idx], self.LOG_ERROR)
                is_check_ok = False
            if self.my_data[i][kemumingcheng_idx] and self.my_data[i][kemumingcheng_idx] not in self.refer_data["科目名称"]:
                self.log("缺少科目名称:" + self.my_data[i][kemumingcheng_idx], self.LOG_ERROR)
                is_check_ok = False
            if self.my_data[i][gongyingshang_idx] and self.my_data[i][gongyingshang_idx] not in self.refer_data["供应商全称"]:
                self.log("缺少供应商全称:" + self.my_data[i][gongyingshang_idx], self.LOG_ERROR)
                is_check_ok = False
        if not is_check_ok:
            self.log("缺少基础信息，停止后续处理！",self.LOG_ERROR)
            return False
        else:
            self.log("基础表检测ok!")
            return True

    def deal_data(self):
        cur = datetime.datetime.now()
        year = cur.year
        month = cur.month
        data_str = str(year) + "/" +str(month) + "/" + str(cur.day)
        # 获取demo
        excel_data = deepcopy(Globals.get_origin_excel_data())
        target_file_name = self.out_put_filename
        origin_data = []
        cur_year_month_str = Globals.get_time_text_year_curmonth()
        try:
            dangyue_jine_idx = self.my_data[0].index(cur_year_month_str)
        except:
            dangyue_jine_idx = self.my_data[0].index(int(cur_year_month_str))
        for i in range(1, len(self.my_data)):
            row = self.my_data[i]
            if row[0] and row[1] and row[dangyue_jine_idx]:
                origin_data.append(row)
        bumen_idx = self.my_data[0].index("部门")
        gongyingshang_idx = self.my_data[0].index("供应商全称")
        kemu_idx = self.my_data[0].index("科目名称")
        zhaiyao_idx = self.my_data[0].index("摘要")
        hetonghao_idx = self.my_data[0].index("合同订单号")
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
            kemu_name = one_data[kemu_idx]
            kemu_bianma1 = self.refer_data["科目名称"][kemu_name]["科目"]
            row[Globals.get_pingzheng_idx("科目")] = kemu_bianma1
            kemu_bianma2 = self.refer_data["部门"][one_data[bumen_idx]]["科目"]
            row2[Globals.get_pingzheng_idx("科目")] = kemu_bianma2

            row[Globals.get_pingzheng_idx("方向")] = 1
            row2[Globals.get_pingzheng_idx("方向")] = 0
            jine = float("%.2f" % (one_data[dangyue_jine_idx]))
            row[Globals.get_pingzheng_idx("原币金额")] = jine
            row2[Globals.get_pingzheng_idx("原币金额")] = jine
            row[Globals.get_pingzheng_idx("借方金额")] = jine
            row2[Globals.get_pingzheng_idx("贷方金额")] = jine
            row[Globals.get_pingzheng_idx("现金流量标记")] = 2
            row2[Globals.get_pingzheng_idx("现金流量标记")] = 2
            row[Globals.get_pingzheng_idx("核算项目1")] = "部门"
            row2[Globals.get_pingzheng_idx("核算项目1")] = "供应商"
            mingcheng1 = one_data[bumen_idx]
            row[Globals.get_pingzheng_idx("名称1")] = self.refer_data["部门"][mingcheng1]["名称1"]
            gongyingshang_name = one_data[gongyingshang_idx]
            row2[Globals.get_pingzheng_idx("名称1")] = gongyingshang_name
            row[Globals.get_pingzheng_idx("编码1")] = self.refer_data["部门"][mingcheng1]["编码1"]
            row2[Globals.get_pingzheng_idx("编码1")] = self.refer_data["供应商全称"][gongyingshang_name]["编码"]
            hesuanxiangmu2 = ""
            if str(kemu_bianma1) == "6401.38":
                hesuanxiangmu2 = "关联公司"
                row[Globals.get_pingzheng_idx("核算项目2")] = hesuanxiangmu2
                row[Globals.get_pingzheng_idx("编码2")] = "Y023"
                row[Globals.get_pingzheng_idx("名称2")] = "上海万科物业服务有限公司"

            zhaiyao = cur_year_month_str + "预提（暂估）" + str(one_data[bumen_idx]) + "-" + gongyingshang_name + "-" + str(one_data[zhaiyao_idx]) + "-" + str(one_data[hetonghao_idx])
            row[Globals.get_pingzheng_idx("摘要")] = zhaiyao
            row2[Globals.get_pingzheng_idx("摘要")] = zhaiyao
            row[Globals.get_pingzheng_idx("辅助账摘要")] = zhaiyao
            row2[Globals.get_pingzheng_idx("辅助账摘要")] = zhaiyao

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
