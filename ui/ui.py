import sys
from PyQt5.QtWidgets import *
from kiwoom.kiwoom import *


class UI:
    def __init__(self):
        print("UI Class")

        self.app = QApplication(sys.argv)

        Kiwoom()

        self.app.exec_()
