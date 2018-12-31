import getpass

class Globals_Class(object):
	def __init__(self):
		# 依赖模块
		self.module_names = [
			"getpass",
			"openpyxl",
			"PyQt5",
			"pyexcel_xlsx",
			"pyexcel_xls",
		]
		# 用户名
		self.user = getpass.getuser()
		# 窗口名
		self.user_title = ""
		# 桌面路径
		self.desktop_path = ""
		# 制单人
		self.eas_name = ""
		title_dict = {
				"pangjm": "金梅君御用",
				"jiangys02": "珊珊专用",
				"zhoukaibing": "金梅君御用",
				"Administrator": "金梅君御用",
		}
		eas_name_dict = {
			"pangjm":"pangjinmei",
			"jiangys02":"jiangys02",
			"zhoukaibing":"zhoukaibing",
			"Administrator":"Administrator",
		}
		if self.user not in title_dict:
			self.user = None
		else:
			self.user_title = title_dict[self.user]
			self.desktop_path = "C:\\Users\\"+ self.user +"\\Desktop\\"
			self.eas_name = eas_name_dict[self.user]

		# 凭证表头
		self.pingzheng_title = [
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
		self.liuliang_title = ["公司", "记账日期", "会计期间", "凭证类型", "凭证号", "币种", "分录号", "对方分录号", "主表信息", "附表信息", "原币", "本位币", "报告币", "主表金额系数",
                "附表金额系数", "性质", "核算项目1", "编码1", "名称1", "核算项目2", "编码2", "名称2", "核算项目3", "编码3", "名称3", "核算项目4", "编码4",
                "名称4", "核算项目5", "编码5", "名称5", "核算项目6", "编码6", "名称6", "核算项目7", "编码7", "名称7", "核算项目8", "编码8", "名称8"]
		self.pingzheng_demo = [
			"E018",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"BB01",
			1,
			"",
			"",
			0,
			0,
			"",
			"",
			self.eas_name,
			"",
			"",
			1,
			"FALSE",
			"",
			"FALSE",
			"1545383744616--0",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
			"",
		]

	# 处理一下日期格式
	def eval_date_format(self, file_path, col_names):
		import openpyxl
		excel = openpyxl.load_workbook(file_path)
		sheet = excel["凭证"]
		for col_name in col_names:
			idx = 0
			col = sheet[col_name]
			for cell in col:
				if idx > 0:
					cell.number_format = "mm-dd-yy"
				else:
					cell.number_format = "General"
				idx = idx + 1
		excel.save(file_path)

	# 根据时间text获取本月
	def get_year_month_day_lastmonth(self, date_str):
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
		return year, month, day, last_month

	# 获取凭证头索引
	def get_pingzheng_idx(self, title):
		return self.pingzheng_title.index(title)

	# 获取空的模板
	def get_origin_excel_data(self):
		return {"凭证":[self.pingzheng_title],"现金流量":[self.liuliang_title]}

Globals = Globals_Class()