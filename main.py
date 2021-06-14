from ui.vba_development_assistant_ui import Ui_MWin
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow
import sys

if __name__ == '__main__':
    app = QApplication(sys.argv)
    base = QMainWindow()
    # try:
    #     with open(r'D:\projects\demo\ui\style.qss') as f:
    #         style = f.read()  # 读取样式表
    #         base.setStyleSheet(style)
    # except:
    #     print("open stylesheet error")
    w = Ui_MWin()
    w.setupUi(base)
    base.show()
    sys.exit(app.exec_())
