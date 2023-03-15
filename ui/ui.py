import sys
from PyQt5.QtWidgets import *
from kiwoom.kiwoom import *


class UI:
    def __init__(self):
        print("UI Class")

        self.app = QApplication(sys.argv)

        self.kiwoom = Kiwoom()

        self.app.exec_()
