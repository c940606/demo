from compare_sheet.compare_sheet import get_sheets, compare_sheet

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import time

TIME_LIMIT = 100
from PyQt5.QtCore import QThread, pyqtSignal


class External(QThread):
    """
    Runs a counter thread.
    """
    countChanged = pyqtSignal(int)
    text = pyqtSignal(str)

    def __init__(self, base_dir, compare_dir, sheet_name):
        super().__init__()
        self.base_dir = base_dir
        self.compare_dir = compare_dir
        self.sheet_name = sheet_name

    def run(self):
        compare_sheet(self.base_dir, self.compare_dir, self.sheet_name, self.countChanged, self.text)


class Ui_MWin(object):
    def setupUi(self, MWin):
        MWin.setObjectName("MWin")
        MWin.resize(781, 508)
        MWin.setMinimumSize(QtCore.QSize(781, 508))
        MWin.setMaximumSize(QtCore.QSize(781, 514))
        # 设置字体
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        MWin.setFont(font)

        # 设置图标
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/images/icon"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MWin.setWindowIcon(icon)

        self.centralWidget = QtWidgets.QWidget(MWin)
        self.centralWidget.setObjectName("centralWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralWidget)
        self.verticalLayout.setContentsMargins(11, 11, 11, 11)
        self.verticalLayout.setSpacing(6)
        self.verticalLayout.setObjectName("verticalLayout")

        self.tabWidget = QtWidgets.QTabWidget(self.centralWidget)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(12)
        self.tabWidget.setFont(font)
        self.tabWidget.setTabPosition(QtWidgets.QTabWidget.North)
        self.tabWidget.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.tabWidget.setIconSize(QtCore.QSize(30, 30))
        self.tabWidget.setElideMode(QtCore.Qt.ElideRight)
        self.tabWidget.setObjectName("tabWidget")

        self.tab1 = QtWidgets.QWidget()
        self.tab1.setObjectName("tab2")

        self.verticalLayout_1 = QtWidgets.QVBoxLayout(self.tab1)
        self.verticalLayout_1.setContentsMargins(11, 11, 11, 11)
        self.verticalLayout_1.setSpacing(6)
        self.verticalLayout_1.setObjectName("verticalLayout_2")

        # 输入文件
        self.Layout_input = QtWidgets.QHBoxLayout()
        self.Layout_input.setSpacing(6)
        self.Layout_input.setObjectName("Layout_input")

        self.lineEdit_input1 = QtWidgets.QLineEdit(self.tab1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.lineEdit_input1.setFont(font)
        self.lineEdit_input1.setInputMask("")
        self.lineEdit_input1.setObjectName("lineEdit_input")
        self.Layout_input.addWidget(self.lineEdit_input1)
        self.base_dir = QtWidgets.QPushButton(self.tab1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.base_dir.setFont(font)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/images/b4"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.base_dir.setIcon(icon1)
        self.base_dir.setObjectName("search")
        self.Layout_input.addWidget(self.base_dir)
        self.verticalLayout_1.addLayout(self.Layout_input)

        self.lineEdit_input2 = QtWidgets.QLineEdit(self.tab1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.lineEdit_input2.setFont(font)
        self.lineEdit_input2.setInputMask("")
        self.lineEdit_input2.setObjectName("lineEdit_input")
        self.Layout_input.addWidget(self.lineEdit_input2)
        self.compare_dir = QtWidgets.QPushButton(self.tab1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.compare_dir.setFont(font)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/images/b4"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.compare_dir.setIcon(icon1)
        self.compare_dir.setObjectName("search")
        self.Layout_input.addWidget(self.compare_dir)
        self.verticalLayout_1.addLayout(self.Layout_input)

        # 对比
        self.Layout_compare = QtWidgets.QHBoxLayout()
        self.Layout_compare.setSizeConstraint(QtWidgets.QLayout.SetMinAndMaxSize)
        self.Layout_compare.setSpacing(6)
        self.Layout_compare.setObjectName("Layout_download")
        self.compareTypes = QtWidgets.QComboBox(self.tab1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.compareTypes.setFont(font)
        self.compareTypes.setObjectName("downloadTypes")
        self.compareTypes.addItem("")
        self.compareTypes.addItem("")
        self.compareTypes.addItem("")
        self.Layout_compare.addWidget(self.compareTypes)
        self.compareProgress = QtWidgets.QProgressBar(self.tab1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.compareProgress.setFont(font)
        self.compareProgress.setObjectName("downloadLinks")
        self.Layout_compare.addWidget(self.compareProgress)
        self.compare = QtWidgets.QPushButton(self.tab1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.compare.setFont(font)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(":/images/b5"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.compare.setIcon(icon2)
        self.compare.setObjectName("download")
        self.Layout_compare.addWidget(self.compare)
        self.Layout_compare.setStretch(0, 2)
        self.Layout_compare.setStretch(1, 10)
        self.Layout_compare.setStretch(2, 2)
        self.verticalLayout_1.addLayout(self.Layout_compare)

        # 显示框
        self.Layout_info = QtWidgets.QVBoxLayout()
        self.Layout_info.setSpacing(6)
        self.Layout_info.setObjectName("Layout_info")
        self.info = QtWidgets.QPlainTextEdit(self.tab1)
        self.info.setReadOnly(True)
        self.info.setObjectName("info")
        self.Layout_info.addWidget(self.info)
        self.verticalLayout_1.addLayout(self.Layout_info)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap(":/images/b2"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.tabWidget.addTab(self.tab1, icon4, "")

        self.verticalLayout.addWidget(self.tabWidget)
        MWin.setCentralWidget(self.centralWidget)

        self.retranslateUi(MWin)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MWin)

        self.connectBtn()

    def retranslateUi(self, MWin):
        _translate = QtCore.QCoreApplication.translate
        MWin.setWindowTitle(_translate("MWin", "VBA开发助手 - ©cw"))


        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab1), _translate("MWin", "对比"))
        self.lineEdit_input1.setPlaceholderText(_translate("MWin", "基准文件"))
        self.compare_dir.setText(_translate("MWin", "浏览"))
        self.lineEdit_input2.setPlaceholderText(_translate("MWin", "比较文件"))
        self.base_dir.setText(_translate("MWin", "浏览"))
        for idx, sheet in enumerate(get_sheets()):
            self.compareTypes.setItemText(idx, _translate("MWin", sheet))

        self.compare.setText(_translate("MWin", "对比"))
        self.compare.setShortcut(_translate("MWin", "Ctrl+Return"))

        self.compareProgress.setValue(0)

    def connectBtn(self):
        self.base_dir.clicked.connect(self.open_base_dire)
        self.compare_dir.clicked.connect(self.open_compare_dire)
        self.compare.clicked.connect(self.compare_file)
        # self.compareTypes.currentIndexChanged.connect(self.downloadTypesChanges) # 类型改变事件

    def open_base_dire(self):
        _translate = QtCore.QCoreApplication.translate
        directory = QFileDialog.getOpenFileName(None, "选择文件")
        self.lineEdit_input1.setText(_translate("MWin", directory[0]))

    def open_compare_dire(self):
        _translate = QtCore.QCoreApplication.translate
        directory = QFileDialog.getOpenFileName(None, "选择文件")

        self.lineEdit_input2.setText(_translate("MWin", directory[0]))

    def compare_file(self):
        _translate = QtCore.QCoreApplication.translate
        self.calc = External(self.lineEdit_input1.text(), self.lineEdit_input2.text(), self.compareTypes.currentText())
        self.calc.countChanged.connect(self.onCountChanged)
        self.calc.text.connect(self.appendText)
        self.calc.start()

    def onCountChanged(self, value):
        self.compareProgress.setValue(value)

    def appendText(self, text):
        _translate = QtCore.QCoreApplication.translate
        self.info.appendPlainText(_translate("MWin", text))

    def downloadTypesChanges(self):
        _translate = QtCore.QCoreApplication.translate
        self.info.setPlainText(_translate("MWin", "hello world"))

