import getpass

class Globals_Class(object):
	def __init__(self):
		# 依赖模块
		self.module_names = [
			("getpass", "getpass"),
			("openpyxl", "openpyxl"),
			("PyQt5", "PyQt5"),
			("pyexcel_xlsx", "pyexcel_xlsx"),
			("pyexcel_xls", "pyexcel_xls"),
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
				"Administrator": "111",
				"PC":"222",
				"Admin":"333",
				"pangjinmei":"333",
				"qqqqq":"1111",
		}
		if self.user not in title_dict:
			self.user = None
		else:
			self.user_title = title_dict[self.user]
			self.desktop_path = "C:/Users/"+ self.user +"/Desktop/"

	# 处理一下日期格式
	def eval_date_format(self, file_path, col_names):
		import openpyxl
		import datetime
		excel = openpyxl.load_workbook(file_path)
		sheet = excel["凭证"]
		cur_date = None
		for col_name in col_names:
			idx = 0
			col = sheet[col_name]
			for cell in col:
				if idx > 0:
					cell.number_format = "mm-dd-yy"
					if not cur_date:
						year, month, day = [int(x) for x in cell.value.split('/')]
						cur_date = datetime.datetime(year, month, day)
					cell.value = cur_date
				else:
					cell.number_format = "General"
				idx = idx + 1
		excel.save(file_path)

	# 获取下个月13号
	def get_next_year_month_day(self):
		import datetime
		cur = datetime.datetime.now()
		year = cur.year
		month = cur.month
		day = 28
		month = month + 1
		if month == 13:
			month = 1
			year = year + 1
		return year, month, day

	# 获取上个月28号
	def get_year_month_day(self):
		import datetime
		cur = datetime.datetime.now()
		year = cur.year
		month = cur.month
		day = 28
		month = month - 1
		if month == 0:
			month = 12
			year = year - 1
		return year, month, day

	# 获取上月28号时间str
	def get_time_text_str(self):
		year, month, day = self.get_year_month_day()
		return str(year) + "/" + str(month) + "/" + str(day)

	# 获取上月的年xx月
	def get_time_text_year_lastmonth(self):
		year, month, day = self.get_year_month_day()
		return str(year) + "%02d"%int(month)

	# 获取下月的年xx月
	def get_time_text_year_nextmonth(self):
		year, month, day = self.get_next_year_month_day()
		return str(year) + "%02d" % int(month)

	#获取本月年xx月
	def get_time_text_year_curmonth(self):
		import datetime
		cur = datetime.datetime.now()
		year = cur.year
		month = cur.month
		return str(year) + "%02d" % int(month)

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

Globals = Globals_Class()
