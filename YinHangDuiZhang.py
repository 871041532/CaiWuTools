# coding:utf8
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
import getpass
from collections import OrderedDict
from pyexcel_xlsx import get_data
import csv
import re
import copy
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

		self.button1 = QPushButton("选择财务表")
		self.button1.clicked.connect(self.click_select_my)
		self.button2 = QPushButton("选择银行表")
		self.button2.clicked.connect(self.click_select_bank)
		self.button3 = QPushButton("开始比较")
		self.button3.clicked.connect(self.click_judge_btn)
		self.text_browser = QTextBrowser()

		mainLayout.addWidget(self.button1)
		mainLayout.addWidget(self.button2)
		mainLayout.addWidget(self.button3)
		mainLayout.addWidget(self.text_browser)

		# data
		self.my_filename = ""
		self.bank_filename = ""
		self.my_data= None
		self.bank_data = None

		self.setLayout(mainLayout)
		self.setWindowTitle("银行对账")
		self.resize(1200, 850)
		self.show()

		if not Globals.user:
			self.close()

	def load_data(self):
		self.log("财务表：" + self.my_filename.split('\\')[-1] + " 导入成功。")
		self.my_data = get_data(self.my_filename)["Sheet1"]

		self.log("银行表：" + self.bank_filename.split('\\')[-1] + " 导入成功。")
		csv_file = open(self.bank_filename, 'r')
		reader = csv.reader(csv_file)
		data = []
		for line in reader:
			data.append(line)
		csv_file.close()
		del data[0]
		self.bank_data = data


	def click_select_my(self):
		info = QFileDialog.getOpenFileName(self, '选择财务表')
		if info and info[0]:
			self.my_filename = info[0]
			self.log("选择了财务表:" + self.my_filename)

	def click_select_bank(self):
		info = QFileDialog.getOpenFileName(self, '选择银行表')
		if info and info[0]:
			self.bank_filename  = info[0]
			self.log("选择了银行表:" + self.bank_filename)

	def clear_log(self):
		self.text_browser.clear()

	def click_judge_btn(self):
		self.clear_log()
		if self.my_filename and self.bank_filename:
			self.load_data()
			result = self.judge_data(1, self.my_data, self.bank_data, "借方", "贷方发生额")
			if result:
				self.judge_data(2, self.my_data, self.bank_data, "贷方", "借方发生额")
		else:
			self.log("请先选择需要对比的财务表与银行表！", ShowWindow.LOG_ERROR)


	# 自己数据，银行数据， 自己title，银行title
	def judge_data(self, idx, old_my_data, old_bank_data, my_title, bank_title):
		self.log("\n\n")
		idx_str = self.get_idx_str(idx)
		self.log(idx_str + "比较财务【" + my_title +"】与银行【" + bank_title +"】...")
		my_idx = old_my_data[0].index(my_title)
		bank_idx = old_bank_data[0].index(bank_title)
		my_data = []
		bank_data = []
		for row in old_my_data:
			if row[my_idx]:
				my_data.append(row)
		del my_data[0]
		del my_data[-1]
		del my_data[-1]
		del my_data[-1]
		my_data.sort(key = lambda x: x[my_idx])

		for row in old_bank_data:
			if row[bank_idx] and str(row[bank_idx]) != " ":
				bank_data.append(row)
				try:
					row[bank_idx] = re.sub('[,]', '', row[bank_idx])
					row[bank_idx] = float(row[bank_idx])			
				except:
					pass
		del bank_data[0]
		bank_data.sort(key = lambda x: x[bank_idx])


		# 检测财务表数据有无负数
		for row in my_data:
			if row[my_idx] < 0:
				self.log("财务表[" + my_title + "]有负数！", ShowWindow.LOG_ERROR)
				return False

		
		lack_bank_rows = []
		surplus_my_rows = []


		my_data_dict = self.get_my_data_dict(my_data, my_idx)
		bank_data_dict = self.get_bank_data_dict(bank_data, bank_idx)
		origin_my_data_dict = copy.deepcopy(my_data_dict)

		keys = bank_data_dict.keys()
		for key in keys:
			rows = bank_data_dict[key]
			for row in rows:
				if key in my_data_dict:
					# 删除一个
					del my_data_dict[key][-1]
					if not my_data_dict[key]:
						del my_data_dict[key]
				else:
					lack_bank_rows.append(row)

		for rows in my_data_dict.values():
			for row in rows:
				surplus_my_rows.append(row)

		if lack_bank_rows or surplus_my_rows:
			self.log("发现不一致:", ShowWindow.LOG_ERROR)
			if lack_bank_rows:
				self.log("银行有，财务没有:")
				for row in lack_bank_rows:
					self.log(self.get_bank_row_error(row, bank_idx, bank_data_dict))

			if surplus_my_rows:
				self.log("")
				self.log("银行没有，财务有:")
				for row in surplus_my_rows:
					self.log(self.get_my_row_error(row, my_idx, origin_my_data_dict))
		else:
			self.log("数据一致，你是最棒的！")

		return True

	# 金额:[row1, row2]
	def get_my_data_dict(self, my_data, my_idx):
		return_data = {}
		for row in my_data:
			value = row[my_idx]
			if value in return_data:
				return_data[value].append(row)
			else:
				return_data[value] = [row]
		return return_data

	# 金额:[row1, row2]
	def get_bank_data_dict(self, bank_data, bank_idx):
		return_data = {}
		for row in bank_data:
			value = row[bank_idx]
			if value in return_data:
				return_data[value].append(row)
			else:
				return_data[value] = [row]
		return return_data

	def get_idx_str(self, idx):
		return {
			1:"一. ",
			2:"二. ",
		}[idx]

	# 获取我的多余行描述
	def get_my_row_error(self, row1, idx1, full_dict):
		str1 = "【未拆分】: "
		str2 = "金额：" + str(row1[idx1]) + " "
		temp_idx = self.my_data[0].index("摘要")
		str2 = str2 + row1[temp_idx] + " "
		temp_idx = self.my_data[0].index("凭证编号")
		str2 = str2 + str(row1[temp_idx])
		if len(full_dict[row1[idx1]]) > 1:
			str2 = str2 + "（金额有重复，不一定这家）"
		return str1 + str2

	# 获取我的却少行描述（银行多余行）
	def get_bank_row_error(self, row2, idx2, full_dict):
		str3 = "【未入账】: "
		str4 = "金额：" + str(row2[idx2]) + " "
		temp_idx = self.bank_data[0].index("对方单位名称")
		str4 = str4 + " " + row2[temp_idx] + " "	
		temp_idx = self.bank_data[0].index("交易时间")
		str4 = str4 + " " + row2[temp_idx] + " "
		if len(full_dict[row2[idx2]]) > 1:
			str4 = str4 + "（金额有重复，不一定这家）"
		return str3 + str4

	def log(self, strs, enum = 1):
		if enum == self.LOG_ERROR:
			pre_str= """<font face="微软雅黑" size="6" color="red">%s</font>"""
			self.text_browser.append(pre_str%str(strs))
		else:
			pre_str= """<font face="微软雅黑" size="4" color="black">%s</font>"""
			self.text_browser.append(pre_str%str(strs))


def into():
	# app = QApplication(sys.argv)
	return ShowWindow()
	# sys.exit(app.exec_())

if __name__ == '__main__':
	app = QApplication(sys.argv)
	ex = ShowWindow()
	sys.exit(app.exec_())
