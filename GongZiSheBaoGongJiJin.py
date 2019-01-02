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

        self.button1 = QPushButton("选择人事拆账表-印力")
        self.button1.clicked.connect(self.click_select_my1)
        self.button1.setMinimumHeight(70)

        self.button2 = QPushButton("选择人事拆账表-印铭")
        self.button2.clicked.connect(self.click_select_my2)
        self.button2.setMinimumHeight(70)

        self.button3 = QPushButton("选择部门编码表")
        self.button3.clicked.connect(self.click_select_refer)
        self.button3.setMinimumHeight(70)

        time_str = Globals.get_time_text_str()
        self.label = QLineEdit(time_str)
        self.label.setPlaceholderText("输入凭证日期 如2018/11/6...")
        self.label.setMinimumHeight(50)

        self.button4 = QPushButton("开始处理")
        self.button4.clicked.connect(self.click_judge_btn)
        self.button4.setMinimumHeight(70)

        self.text_browser = QTextBrowser()

        mainLayout.addWidget(self.label)
        mainLayout.addWidget(self.button1)
        mainLayout.addWidget(self.button2)
        mainLayout.addWidget(self.button3)
        mainLayout.addWidget(self.button4)

        mainLayout.addWidget(self.text_browser)

        # data
        self.my_filename1 = "C:\\Users\\pangjm\\Desktop\\1.xls" #""
        self.my_filename2 = "C:\\Users\\pangjm\\Desktop\\2.xlsx" #""
        self.refer_filename = "C:\\Users\\pangjm\\Desktop\\部门.xlsx" #""

        self.output_dir = Globals.desktop_path + "工资社保公积金\\"
        self.out_put_filename1 = Globals.desktop_path + "工资社保公积金\\工资社保公积金汇总.xlsx"

        self.out_put_filename2 = Globals.desktop_path + "工资社保公积金\\结转工资.xlsx"
        self.out_put_filename3 = Globals.desktop_path + "工资社保公积金\\结转公积金.xlsx"
        self.out_put_filename4 = Globals.desktop_path + "工资社保公积金\\结转社保.xlsx"
        self.out_put_filename5 = Globals.desktop_path + "工资社保公积金\\预提工资.xlsx"
        self.out_put_filename6 = Globals.desktop_path + "工资社保公积金\\预提社保.xlsx"


        self.refer_data = None
        self.my_data = None

        self.setLayout(mainLayout)
        self.setWindowTitle("工资社保公积金")
        self.setGeometry(550, 150, 900, 800)
        self.show()

        if not Globals.user:
            self.close()

    def click_select_refer(self):
        info = QFileDialog.getOpenFileName(self, '选择部门编码表')
        if info and info[0]:
            self.refer_filename = info[0]
            self.log("选择了部门编码表：" + self.refer_filename)

    def click_select_my1(self):
        info = QFileDialog.getOpenFileName(self, '选择表1')
        if info and info[0]:
            self.my_filename1 = info[0]
            self.log("选择了表1：" + self.my_filename1)

    def click_select_my2(self):
        info = QFileDialog.getOpenFileName(self, '选择表2')
        if info and info[0]:
            self.my_filename2 = info[0]
            self.log("选择了表2：" + self.my_filename2)

    def clear_log(self):
        self.text_browser.clear()

    def click_judge_btn(self):
        self.clear_log()
        if not (self.my_filename1):
            self.log("请先选择表1~(^v^)~", ShowWindow.LOG_ERROR)
        elif not (self.my_filename2):
            self.log("请先选择表2~(^v^)~", ShowWindow.LOG_ERROR)
        elif not (self.refer_filename):
            self.log("请先选择部门编码表~(^v^)~", ShowWindow.LOG_ERROR)
        elif not self.label.text():
            self.log("请先输入凭证日期！", ShowWindow.LOG_ERROR)
        else:
            if not os.path.exists(self.output_dir):
                os.mkdir(self.output_dir)
            self.load_data()
            self.deal_jiezhuangongzi_data(self.out_put_filename2)
            #self.deal_data(self.out_put_filename3)
            #self.deal_data(self.out_put_filename4)
            #self.deal_data(self.out_put_filename5)
            #self.deal_data(self.out_put_filename6)
            self.log("处理完毕。")

    def load_data(self):
        excel1 = get_data(self.my_filename1)
        the_keys1 = list(excel1.keys())[0:3]
        excel2 = get_data(self.my_filename2)
        the_keys2 = list(excel2.keys())[0:3]

        gongzi_list = self.eval_gongzi(excel1[the_keys1[0]], excel2[the_keys2[0]])
        shebao_list = self.eval_shebao(excel1[the_keys1[1]], excel2[the_keys2[1]])
        gongjijin_list = self.eval_gongjijin(excel1[the_keys1[2]], excel2[the_keys2[2]])

        self.my_data = {
            "工资":gongzi_list,
            "社保":shebao_list,
            "公积金":gongjijin_list,
        }
        save_data(self.out_put_filename1, self.my_data)
        self.refer_data = get_data(self.refer_filename)

    def eval_gongzi(self, sheet1, sheet2):
        return_data = []
        title_list1 = sheet1[0]
        title_dict1 = {}
        title_list2 = sheet2[0]
        title_dict2 = {}
        return_data.append(title_list2)
        for i in range(len(title_list1)):
            title_dict1[title_list1[i]] = i
        for i in range(len(title_list2)):
            title_dict2[title_list2[i]] = i
        for row in sheet2:
            chengben_idx = title_dict2["成本部门"]
            if len(row) > chengben_idx and row[chengben_idx] and "七宝" in row[chengben_idx]:
                # row[title_dict2["核算-成本公司"]] = "上海新宝置业有限公司"
                return_data.append(row)
        for row in sheet1:
            chengben_idx1 = title_dict1["成本部门"]
            if len(row) > chengben_idx1 and "七宝" in row[chengben_idx1]:
                new_row = []
                for key in title_list2:
                    if key in title_dict1:
                        new_row.append(row[title_dict1[key]])
                    else:
                        new_row.append("")
                return_data.append(new_row)
        return return_data


    def eval_shebao(self, sheet1, sheet2):
        return self.eval_gongzi(sheet1, sheet2)

    def eval_gongjijin(self, sheet1, sheet2):
        for i in range(len(sheet1)):
            if not sheet1[0]:
                del sheet1[0]
            else:
                break
        for i in range(len(sheet2)):
            if not sheet2[0]:
                del sheet2[0]
            else:
                break
        return self.eval_gongzi(sheet1, sheet2)


    def save_data(self, target_file_name, data):
        save_data(target_file_name, data)
        # 处理日期
        Globals.eval_date_format(target_file_name, ["B", "C", "BN"])

    # 处理结转工资
    def deal_jiezhuangongzi_data(self, target_file_name):
        excel_data = Globals.get_origin_excel_data()

        origin_data = self.my_data["工资"]
        refer_data = self.refer_data

        for i in range(1, len(origin_data)):
            origin_row = origin_data[i]

            target_rows = []
            #if origin_row[self.get_gongzi_idx("求和项:1、应付工资")]:
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            #if origin_row[self.get_gongzi_idx("求和项:1、应付工资")]:
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))

            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))

            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))

            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            target_rows.append(deepcopy(Globals.pingzheng_demo))
            for row in target_rows:
                row[Globals.get_pingzheng_idx("摘要")] = origin_row[self.get_gongzi_idx("成本部门")]
                row[Globals.get_pingzheng_idx("记账日期")] = self.label.text()
                row[Globals.get_pingzheng_idx("业务日期")] = self.label.text()
                excel_data["凭证"].append(row)



        self.save_data(target_file_name,excel_data)

    def get_gongzi_idx(self, title):
        return self.my_data["工资"][0].index(title)

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
