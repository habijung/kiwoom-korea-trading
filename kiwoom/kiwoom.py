from PyQt5.QAxContainer import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Class Kiwoom")

        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)

    def login_slot(self, errCode):
        print(errCode)

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")
