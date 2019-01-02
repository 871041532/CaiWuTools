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

		self.button1 = QPushButton("选择地产表：标准凭证引出-GX凭证类型")
		self.button1.clicked.connect(self.click_select_my)
		self.button1.setMinimumHeight(70)

		self.button2 = QPushButton("选择基础信息表")
		self.button2.clicked.connect(self.click_select_bank)
		self.button2.setMinimumHeight(70)

		time_str = Globals.get_time_text_str()
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
		self.my_filename = None  # "C:\\Users\\pangjm\\Desktop\\共享导入\\10月导出.xls" #""
		self.bank_filename = None  # "C:\\Users\\pangjm\\Desktop\\共享导入\\共享凭证导入模板.xlsx" #""
		self.out_put_filename = Globals.desktop_path + "共享生成凭证.xlsx"
		self.my_data= None
		self.bank_data = None

		self.setLayout(mainLayout)
		self.setWindowTitle("共享导入")
		self.setGeometry(550, 150, 900, 800)
		self.show()

		if not Globals.user:
			self.close()

	def click_select_my(self):
		info = QFileDialog.getOpenFileName(self, '选择地产表：标准凭证引出-GX凭证类型')
		if info and info[0]:
			self.my_filename = info[0]
			self.log("选择了地产表：" + self.my_filename)

	def click_select_bank(self):
		info = QFileDialog.getOpenFileName(self, '选择基础信息表')
		if info and info[0]:
			self.bank_filename  = info[0]
			self.log("选择了基础信息表：" + self.bank_filename)

	def clear_log(self):
		self.text_browser.clear()

	def click_judge_btn(self):
		self.clear_log()
		if not (self.my_filename and self.bank_filename):
			self.log("请先选择需要对比的表！", ShowWindow.LOG_ERROR)
		elif not self.label.text():
			self.log("请先输入凭证日期！", ShowWindow.LOG_ERROR)
		else:
			self.load_data()
			self.deal_data()
			self.log("处理完毕。")


	def load_data(self):
		self.log("地产导出表：" + self.my_filename.split('\\')[-1] + " 导入成功。")
		self.my_data = get_data(self.my_filename)["凭证"]

		self.log("基础信息表：" + self.bank_filename.split('\\')[-1] + " 导入成功。")
		self.bank_data = get_data(self.bank_filename)["基础信息"]

	# target_data
	def deal_data(self):
		target_excel = {"凭证":[]}
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
		row2 = ["公司",	"记账日期",	"会计期间",	"凭证类型",	"凭证号",	"币种",	"分录号",	"对方分录号",	"主表信息",	"附表信息",	"原币",	"本位币",	"报告币", "主表金额系数",	"附表金额系数",	"性质",	"核算项目1",	"编码1",	 "名称1",	"核算项目2",	"编码2",	"名称2",	"核算项目3",	"编码3",	 "名称3", 	"核算项目4",	"编码4",	"名称4",	"核算项目5",	"编码5",	"名称5",	"核算项目6",	"编码6",	"名称6",	"核算项目7",	"编码7",	"名称7",	"核算项目8",	"编码8", "名称8"	]
		target_excel["现金流量"] = [row2]
		refer_data = self.bank_data
		bank_data = self.my_data

		target_rows = []
		row_num = len(bank_data) - 1
		for i in range(row_num):
			target_rows.append([])



		date_str = self.label.text()
		month = date_str.split('/')[1]
		# 替换信息
		replace_dict = self.get_replace_dict(refer_data)

		# 银行凭证号
		pingzhenghao_bank_idx = bank_data[0].index("凭证号")
		# 摘要索引
		zhaoyao_bank_idx = bank_data[0].index("摘要")
		# 方向索引
		fangxiang_bank_idx = bank_data[0].index("方向")
		# 原币金额索引
		yuanbijine_bank_idx = bank_data[0].index("原币金额")
		# 借方金额索引
		jiefangjine_bank_idx = bank_data[0].index("借方金额")
		# 贷方金额
		daifangjine_bank_idx = bank_data[0].index("贷方金额")
		# 科目索引
		kemu_bank_idx = bank_data[0].index("科目")
		# 名称1索引
		mingcheng1_idx = bank_data[0].index("名称1")

		idx = 0
		for row in target_rows:
			idx = idx + 1
			# 公司
			row.append("E018")
			# 记账日期
			row.append(date_str)
			# 业务日期
			row.append(date_str)
			# 会计期间
			row.append(int(month))
			# 凭证类型
			row.append("付")
			# 凭证号
			row.append("20180600828")
			# 分录号
			row.append(idx)
			# 摘要
			zhaiyao_value = bank_data[idx][pingzhenghao_bank_idx] + " "+ bank_data[idx][zhaoyao_bank_idx]
			row.append(zhaiyao_value)
			# 科目
			kemu_bank_data = str(bank_data[idx][kemu_bank_idx])
			if kemu_bank_data in replace_dict["科目"]:
				kemu_bank_data = replace_dict["科目"][kemu_bank_data]
				# 如果科目是6401.06并且摘要里面有报销
				if kemu_bank_data == "6401.06" and "报销：" in zhaiyao_value:
					kemu_bank_data = "6401.07.01"
				row.append(kemu_bank_data)
			else:
				self.log("基础信息需要补充地产科目：" + kemu_bank_data, ShowWindow.LOG_ERROR)
				row.append("")
			# 币种
			row.append("BB01")
			# 汇率
			row.append(1)
			# 方向
			row.append(int(bank_data[idx][fangxiang_bank_idx]))
			# 原币金额
			row.append(bank_data[idx][yuanbijine_bank_idx])
			# 数量
			row.append(0)
			# 单价
			row.append(0)
			# 借方金额
			row.append(bank_data[idx][jiefangjine_bank_idx])
			# 贷方金额
			row.append(bank_data[idx][daifangjine_bank_idx])
			# 制单人
			row.append(Globals.eas_name)
			# 过账人
			row.append("")
			# 审核人
			row.append("")
			# 附件数量
			row.append(1)
			# 过账标记
			row.append("TRUE")
			# 机制凭证
			row.append("")
			# 删除标记
			row.append("FALSE")
			# 凭证序号
			row.append("1533086430813--0")
			# 单位
			row.append("")
			# 参考信息
			row.append("")
			# 是否有现金流量
			row.append("")
			# 现金流量标记
			row.append(6)
			# 业务编号
			row.append("")
			# 结算方式
			row.append("")
			# 结算号
			row.append("")
			# 辅助账摘要
			row.append(zhaiyao_value)

			# 核算项目1
			hesuanxiangmu1_value = ""
			if kemu_bank_data in replace_dict["核算项目1"]:
				hesuanxiangmu1_value = replace_dict["核算项目1"][kemu_bank_data]
				row.append(hesuanxiangmu1_value)
			else:
				row.append("")
			# 编码1
			if kemu_bank_data in replace_dict["编码1"]:
				row.append(replace_dict["编码1"][kemu_bank_data])
			else:
				row.append("")
			# 名称1
			mingcheng1_value = ""
			if kemu_bank_data in replace_dict["名称1"]:
				mingcheng1_value = replace_dict["名称1"][kemu_bank_data]
				row.append(mingcheng1_value)
			else:
				if hesuanxiangmu1_value == "部门" or hesuanxiangmu1_value == "职员":
					mingcheng1_value = bank_data[idx][mingcheng1_idx]
					row.append(mingcheng1_value)
				else:
					row.append("")

			# 根据部门名称再次计算编码1
			if hesuanxiangmu1_value == "部门" or hesuanxiangmu1_value == "职员":
				if mingcheng1_value in replace_dict["编码1(用于部门职员)"]:
					row[-2] = replace_dict["编码1(用于部门职员)"][mingcheng1_value]
				else:
					self.log("部门：" + mingcheng1_value + "没有对应编码", self.LOG_ERROR)

			# 核算项目2
			if kemu_bank_data in replace_dict["核算项目2"]:
				row.append(replace_dict["核算项目2"][kemu_bank_data])
			else:
				row.append("")
			# 编码2
			if kemu_bank_data in replace_dict["编码2"]:
				row.append(replace_dict["编码2"][kemu_bank_data])
			else:
				row.append("")
			# 名称2
			if kemu_bank_data in replace_dict["名称2"]:
				row.append(replace_dict["名称2"][kemu_bank_data])
			else:
				row.append("")

			# 项目3
			row.append("")
			row.append("")
			row.append("")
			# 项目4
			row.append("")
			row.append("")
			row.append("")
			# 项目5
			row.append("")
			row.append("")
			row.append("")
			# 项目6
			row.append("")
			row.append("")
			row.append("")
			# 项目7
			row.append("")
			row.append("")
			row.append("")
			# 项目8
			row.append("")
			row.append("")
			row.append("")
			# 发票号
			row.append("")
			# 换票证号
			row.append("")
			# 客户
			row.append("")
			# 费用类别
			row.append("")
			# 收款人
			row.append("")
			# 物料
			row.append("")
			# 财务组织
			row.append("")
			# 供应商
			row.append("")
			# 辅助帐业务日期
			row.append(date_str)
			# 到期日
			row.append("")


		for row in target_rows:
			target_data.append(row)
		save_data(self.out_put_filename, target_excel)
		# 处理日期
		Globals.eval_date_format(self.out_put_filename, ["B", "C", "BN"])
		
	# 
	def get_replace_dict(self, refer_data):
		return_data = {}
		# 科目替换
		kemu_dict = {}
		dichankemu_idx = refer_data[0].index("地产科目")
		yinlikemu_idx = refer_data[0].index("印力科目")
		for i in range(len(refer_data) - 1):
			row = refer_data[i + 1]
			v1 = str(row[dichankemu_idx])
			v2 = str(row[yinlikemu_idx])
			if v1 and v2:
				kemu_dict[v1] = v2
		return_data["科目"] = kemu_dict
		# 核算项目1替换
		hesuanxiangmu_dict = {}
		keumu_idx = refer_data[0].index("科目")
		hesuanxiangmu1_idx = refer_data[0].index("核算项目1")
		for i in range(len(refer_data) - 1):
			row = refer_data[i + 1]
			v1 = str(row[keumu_idx]) if len(row) > keumu_idx else ""
			v2 = str(row[hesuanxiangmu1_idx]) if len(row) > hesuanxiangmu1_idx else ""
			if v1:
				hesuanxiangmu_dict[v1] = v2
		return_data["核算项目1"] = hesuanxiangmu_dict
		# 名称1替换
		mingcheng1_dict = {}
		mingcheng1_idx = refer_data[0].index("名称1")
		for i in range(len(refer_data) - 1):
			row = refer_data[i + 1]
			v1 = str(row[keumu_idx]) if len(row) > keumu_idx else ""
			v2 = str(row[mingcheng1_idx]) if len(row) > mingcheng1_idx else ""
			judge_v3 = hesuanxiangmu_dict.get(v1)
			if v1 and judge_v3 != "部门" and judge_v3 != "职员":
				mingcheng1_dict[v1] = v2
		return_data["名称1"] = mingcheng1_dict
		# 编码1
		bianma1_dict = {}
		bianma1_idx = refer_data[0].index("编码1")
		for i in range(len(refer_data) - 1):
			row = refer_data[i + 1]
			v1 = str(row[keumu_idx]) if len(row) > keumu_idx else ""
			v2 = str(row[bianma1_idx]) if len(row) > bianma1_idx else ""
			judge_v3 = hesuanxiangmu_dict.get(v1)
			if v1 and judge_v3 != "部门" and judge_v3 != "职员":
				bianma1_dict[v1] = v2
		return_data["编码1"] = bianma1_dict
		# 核算项目2
		hesuanxiangmu_dict2 = {}
		hesuanxiangmu2_idx = refer_data[0].index("核算项目2")
		for i in range(len(refer_data) - 1):
			row = refer_data[i + 1]
			v1 = str(row[keumu_idx]) if len(row) > keumu_idx else ""
			v2 = str(row[hesuanxiangmu2_idx]) if len(row) > hesuanxiangmu2_idx else ""
			if v1:
				hesuanxiangmu_dict2[v1] = v2
		return_data["核算项目2"] = hesuanxiangmu_dict2
		# 编码2
		bianma2_dict = {}
		bianma2_idx = refer_data[0].index("编码2")
		for i in range(len(refer_data) - 1):
			row = refer_data[i + 1]
			v1 = str(row[keumu_idx]) if len(row) > keumu_idx else ""
			v2 = str(row[bianma2_idx]) if len(row) > bianma2_idx else ""
			if v1:
				bianma2_dict[v1] = v2
		return_data["编码2"] = bianma2_dict
		# 名称2
		mingcheng2_dict = {}
		mingcheng2_idx = refer_data[0].index("名称2")
		for i in range(len(refer_data) - 1):
			row = refer_data[i + 1]
			v1 = str(row[keumu_idx]) if len(row) > keumu_idx else ""
			v2 = str(row[mingcheng2_idx]) if len(row) > mingcheng2_idx else ""
			if v1:
				mingcheng2_dict[v1] = v2
		return_data["名称2"] = mingcheng2_dict
		# 编码1(用于部门职员)
		bianma1_for_bumen_dict = {}
		bianma1_for_bumen_idx = refer_data[0].index("编码1(用于部门职员)")
		mingcheng1_for_bumen_idx = refer_data[0].index("名称1(用于部门职员)")
		for i in range(len(refer_data) - 1):
			row = refer_data[i + 1]
			v1 = str(row[mingcheng1_for_bumen_idx]) if len(row) > mingcheng1_for_bumen_idx else ""
			v2 = str(row[bianma1_for_bumen_idx]) if len(row) > bianma1_for_bumen_idx else ""
			if v1:
				bianma1_for_bumen_dict[v1] = v2
		return_data["编码1(用于部门职员)"] = bianma1_for_bumen_dict
		return return_data



	def log(self, strs, enum = 1):
		if enum == self.LOG_ERROR:
			pre_str= """<font face="微软雅黑" size="6" color="red">%s</font>"""
			self.text_browser.append(pre_str%str(strs))
		else:
			pre_str= """<font face="微软雅黑" size="4" color="black">%s</font>"""
			self.text_browser.append(pre_str%str(strs))


def into():
	return ShowWindow()

if __name__ == '__main__':
	app = QApplication(sys.argv)
	ex = ShowWindow()
	sys.exit(app.exec_())
