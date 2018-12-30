from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
import getpass
import YinHangDuiZhang
import GongXiangDaoRu
import JieZhuanYunPos
import JieZhuanShuJuCaiJiSheBei
import ZhiFuBaoPark
import WeiXinPark
from Globals import Globals


class ShowWindow(QWidget):

	"""docstring for ClassName"""
	def __init__(self):
		super(ShowWindow, self).__init__()
		self.initUI()

	def initUI(self):
		mainLayout = QVBoxLayout()
		self.button1 = QPushButton("银行对账")
		self.button1.clicked.connect(self.click_select_my)
		self.button2 = QPushButton("共享导入")
		self.button2.clicked.connect(self.click_select_bank)
		jiezhuan_layout = QHBoxLayout()
		self.button3 = QPushButton("结转云Pos")
		self.button3.clicked.connect(self.click_jiezhuanyun_pos)
		self.button4 = QPushButton("结转数据采集设备")
		self.button4.clicked.connect(self.JieZhuanShuJuCaiJiSheBei)
		jiezhuan_layout.addWidget(self.button3)
		jiezhuan_layout.addWidget(self.button4)

		zhifubao_layout = QHBoxLayout()
		self.button5 = QPushButton("停车场支付宝返款")
		self.button5.clicked.connect(self.ZhiFuBaoPark)
		self.button6 = QPushButton("停车场微信返款")
		self.button6.clicked.connect(self.WeiXinPark)
		zhifubao_layout.addWidget(self.button5)
		zhifubao_layout.addWidget(self.button6)


		self.button1.setMinimumHeight(100)
		self.button2.setMinimumHeight(100)
		self.button3.setMinimumHeight(100)
		self.button4.setMinimumHeight(100)
		self.button5.setMinimumHeight(100)
		self.button6.setMinimumHeight(100)

		mainLayout.addLayout(jiezhuan_layout)
		mainLayout.addLayout(zhifubao_layout)
		mainLayout.addWidget(self.button2)
		mainLayout.addWidget(self.button1)

		self.setLayout(mainLayout)
		self.setGeometry(700, 300, 400, 250)
		self.show()

		if not Globals.user:
			self.close()
		self.setWindowTitle(Globals.user_title)

	# 选择了银行对账
	def click_select_my(self):
		Globals.current_win = YinHangDuiZhang.into()

	# 选择了共享导入
	def click_select_bank(self):
		Globals.current_win = GongXiangDaoRu.into()

	# 选择了结转云pos
	def click_jiezhuanyun_pos(self):
		Globals.current_win = JieZhuanYunPos.into()

	# 选择了结转数据采集设备
	def JieZhuanShuJuCaiJiSheBei(self):
		Globals.current_win = JieZhuanShuJuCaiJiSheBei.into()

	# 支付宝停车场
	def ZhiFuBaoPark(self):
		Globals.current_win = ZhiFuBaoPark.into()

	# 微信停车场
	def WeiXinPark(self):
		Globals.current_win = WeiXinPark.into()


if __name__ == '__main__':
	app = QApplication(sys.argv)
	Globals.current_win = ShowWindow()
	Globals.MainWin = ShowWindow
	sys.exit(app.exec_())
