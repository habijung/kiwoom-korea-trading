from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.error_code import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Class Kiwoom")

        # Event Loop
        self.login_event_loop = None
        ############################

        # Variables
        self.account_num = None

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect()
        self.get_account_info()

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)

    def login_slot(self, errCode):
        print(errors(errCode))

        self.login_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        account_num = account_list.split(";")[0]

        print(f"Account: {account_num}")

        return account_num
