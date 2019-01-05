
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
import datetime
import openpyxl

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

        self.button1 = QPushButton("选择海鼎导出结转表")
        self.button1.clicked.connect(self.click_select_my)
        self.button1.setMinimumHeight(70)

        self.button2 = QPushButton("选择需要填充的7个结转模板")
        self.button2.clicked.connect(self.click_select_output)
        self.button2.setMinimumHeight(70)

        cur = datetime.datetime.now()
        self.label = QLineEdit(str(cur.month) + "月")
        self.label.setPlaceholderText("输入结转模板sheet名字 如: 1月")
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
        self.my_filename = None #Globals.desktop_path + "原始表.xls"
        self.out_put_filenames = None #(
        #         Globals.desktop_path + "2-2.2019年收入结转与应计 (1)\\1.1莘宝结转仓库租赁收入凭证.xlsx",
        #         Globals.desktop_path + "2-2.2019年收入结转与应计 (1)\\1.2莘宝结转浮动提成租金凭证-上月.xlsx",
        #         Globals.desktop_path + "2-2.2019年收入结转与应计 (1)\\1.3莘宝结转固定租金凭证.xlsx",
        #         Globals.desktop_path + "2-2.2019年收入结转与应计 (1)\\1.4莘宝结转广告位租赁收入凭证.xlsx",
        #         Globals.desktop_path + "2-2.2019年收入结转与应计 (1)\\1.5莘宝结转推广费（提成）凭证-上月.xlsx",
        #         Globals.desktop_path + "2-2.2019年收入结转与应计 (1)\\1.6莘宝结转推广费（固定）凭证.xlsx",
        #         Globals.desktop_path + "2-2.2019年收入结转与应计 (1)\\1.7莘宝结转物业管理费凭证.xlsx",
        # )
        self.my_data = None
        self.out_put_datas = None

        self.setLayout(mainLayout)
        self.setWindowTitle("月中结转主合同")
        self.setGeometry(550, 150, 900, 800)
        self.show()

        if not Globals.user:
            self.close()

    # 选择目标表
    def click_select_output(self):
        info = QFileDialog.getOpenFileNames(self, '选择需要填充的7个结转模板')
        if info and info[0]:
            self.out_put_filenames = info[0]
            self.log("选择了结转模板：" + str(self.out_put_filenames))

    # 选择源
    def click_select_my(self):
        info = QFileDialog.getOpenFileName(self, '选择海鼎导出结转表')
        if info and info[0]:
            self.my_filename = info[0]
            self.log("选择海鼎导出结转表：" + self.my_filename)

    def clear_log(self):
        self.text_browser.clear()

    def click_judge_btn(self):
        self.clear_log()
        if not (self.my_filename):
            self.log("请先选择海鼎导出结转表~(^v^)~", ShowWindow.LOG_ERROR)
        elif not self.label.text():
            self.log("请先输入月份！", ShowWindow.LOG_ERROR)
        elif not self.out_put_filenames:
            self.log("请先选择一点几输出文件~(^v^)~", ShowWindow.LOG_ERROR)
        else:
            self.load_data()
            self.deal_data()
            self.log("处理完毕。")

    def load_data(self):
        self.log(self.my_filename.split('/')[-1] + " 导入成功。")
        self.my_data = get_data(self.my_filename)
        self.my_data = list(self.my_data.values())[0]
        self.my_data = self.my_data[1:]
        row_0 = self.my_data[0]
        row_1 = self.my_data[1]
        for i in range(len(row_1)):
            if row_0[i] == "":
                row_0[i] = row_1[i]
        self.my_data = self.my_data[1:]
        self.my_data[0] = row_0

        self.out_put_datas = {}
        for full_path in self.out_put_filenames:
            file_name = full_path.split('/')[-1]
            self.out_put_datas[file_name] = (full_path, openpyxl.load_workbook(full_path))
            self.log(file_name + " 导入成功。")

    # 处理数据
    def deal_data(self):
        # 处理 1.1
        for k, v in self.out_put_datas.items():
            title_key = ""
            if "1.1" == k[0:3]:
                # 获取原始数据
                title_key = "仓库租赁费"
            elif "1.2" == k[0:3]:
                title_key = "浮动提成租金"
            elif "1.3" == k[0:3]:
                title_key = "固定租金"
            elif "1.4" == k[0:3]:
                title_key = "广告位租赁费"
            elif "1.5" == k[0:3]:
                title_key = "推广费(销售提成）"
            elif "1.6" == k[0:3]:
                title_key = "推广费（固定）"
            elif "1.7" == k[0:3]:
                title_key = "物业管理费"
            if title_key:
                title_idx = self.my_data[0].index(title_key)
                origin_data = []
                for i in range(1, len(self.my_data)):
                    if len(self.my_data[i]) > title_idx and self.my_data[i][title_idx] and self.my_data[i][0]:
                        origin_data.append(self.my_data[i])
                gongsimingcheng_idx = self.my_data[0].index("公司名称")
                dianpuzhaopai_idx = self.my_data[0].index("店铺招牌")
                puweihao_idx = self.my_data[0].index("铺位号")
                def key_func(elem):
                    return elem[puweihao_idx]
                origin_data.sort(key = key_func)
                # 获取目标数据
                excel = v[1]
                sheet = excel[self.label.text()]
                # 写入
                for i in range(len(origin_data)):
                    save_row_idx = i + 3
                    # 租户名称
                    sheet.cell(row = save_row_idx, column = 1).value = origin_data[i][gongsimingcheng_idx]
                    # 店铺招牌
                    sheet.cell(row = save_row_idx, column = 2).value = origin_data[i][dianpuzhaopai_idx]
                    # 铺位号
                    sheet.cell(row = save_row_idx, column = 3).value = origin_data[i][puweihao_idx]
                    # 金额
                    sheet.cell(row=save_row_idx, column = 4).value = origin_data[i][title_idx]
                v[1].save(v[0])

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
