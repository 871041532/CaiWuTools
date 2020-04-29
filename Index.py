from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
import getpass
from Globals import Globals


class ShowWindow(QWidget):

	"""docstring for ClassName"""
	def __init__(self):
		super(ShowWindow, self).__init__()
		self.initUI()

	def initUI(self):
		mainLayout = QVBoxLayout()
		self.button1 = QPushButton("1")
		self.button1.clicked.connect(self.select_one)
		# self.button2 = QPushButton("2")
		# self.button2.clicked.connect(self.select_one)
		# jiezhuan_layout = QHBoxLayout()
		# self.button3 = QPushButton("3")
		# self.button3.clicked.connect(self.select_one)
		# self.button4 = QPushButton("4")
		# self.button4.clicked.connect(self.select_one)
		# jiezhuan_layout.addWidget(self.button3)
		# jiezhuan_layout.addWidget(self.button4)
		#
		# yuezhong_jiezhuan_layout = QHBoxLayout()
		# self.button7 = QPushButton("5")
		# self.button7.clicked.connect(self.select_one)
		# self.button8 = QPushButton("6")
		# self.button8.clicked.connect(self.select_one)
		# yuezhong_jiezhuan_layout.addWidget(self.button7)
		# yuezhong_jiezhuan_layout.addWidget(self.button8)
		#
		# self.button9 = QPushButton("7")
		# self.button9.clicked.connect(self.select_one)
		#
		# zhifubao_layout = QHBoxLayout()
		# self.button5 = QPushButton("8")
		# self.button5.clicked.connect(self.select_one)
		# self.button6 = QPushButton("9")
		# self.button6.clicked.connect(self.select_one)
		# zhifubao_layout.addWidget(self.button5)
		# zhifubao_layout.addWidget(self.button6)
		#
		#
		# self.button1.setMinimumHeight(90)
		# self.button2.setMinimumHeight(90)
		# self.button3.setMinimumHeight(90)
		# self.button4.setMinimumHeight(90)
		# self.button5.setMinimumHeight(90)
		# self.button6.setMinimumHeight(90)
		# self.button7.setMinimumHeight(90)
		# self.button8.setMinimumHeight(90)
		# self.button9.setMinimumHeight(90)
		#
		# mainLayout.addLayout(yuezhong_jiezhuan_layout)
		# mainLayout.addWidget(self.button9)
		# mainLayout.addLayout(jiezhuan_layout)
		# mainLayout.addLayout(zhifubao_layout)
		# mainLayout.addWidget(self.button2)
		mainLayout.addWidget(self.button1)

		self.setLayout(mainLayout)
		self.setGeometry(700, 200, 400, 50)
		self.show()

		if not Globals.user:
			self.close()
		self.setWindowTitle(Globals.user_title)

	# 选择了银行对账
	def select_one(self):
		pass
		# Globals.current_win = YinHangDuiZhang.into()

if __name__ == '__main__':
	app = QApplication(sys.argv)
	Globals.current_win = ShowWindow()
	Globals.MainWin = ShowWindow
	sys.exit(app.exec_())
