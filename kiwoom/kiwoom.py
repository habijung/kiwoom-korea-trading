from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.error_code import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Class Kiwoom")

        # Event Loop
        self.login_event_loop = None
        self.detail_account_info_event_loop = None
        self.detail_account_info_event_loop_2 = None
        ############################

        # Variables
        self.account_num = None
        self.use_money = 0
        self.use_money_percent = 0.5

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()  # 예수금 요청
        self.detail_account_mystock()  # 계좌평가잔고내역요청

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def login_slot(self, errCode):
        print(errors(errCode))

        self.login_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(";")[0]

        print(f"Account: {self.account_num}")

        return self.account_num

    def detail_account_info(self):
        print("예수금 요청")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")

        # RQName은 마음대로 작성해도 됨
        self.dynamicCall(
            "CommRqData(String, String, int, String)",
            "예수금상세현황요청",
            "opw00001",
            "0",
            "2000",
        )

        # 요청 후 event loop를 실행해서 TR 요청 데이터가 올 때까지 대기
        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        print("계좌평가잔고내역요청")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall(
            "CommRqData(String, String, int, String)",
            "계좌평가잔고내역요청",
            "opw00018",
            sPrevNext,
            "2000",
        )

        self.detail_account_info_event_loop_2 = QEventLoop()
        self.detail_account_info_event_loop_2.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        """
        TR 요청을 받는 구역
        :param sScrNo:
        :param sRQName:
        :param sTrCode:
        :param sRecordName:
        :param sPrevNext: 다음 페이지가 있는지
        :return:
        """

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall(
                "GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금"
            )
            print(f"예수금: {int(deposit)}")

            # 예수금을 일정 부분만 사용하도록 설정
            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 4

            ok_deposit = self.dynamicCall(
                "GetCommData(String, String, int, String)",
                sTrCode,
                sRQName,
                0,
                "출금가능금액",
            )
            print(f"출금가능금액: {int(ok_deposit)}")

            # TR 요청 처리를 했기 때문에 event loop를 종료
            self.detail_account_info_event_loop.exit()

        elif sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall(
                "GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액"
            )
            total_buy_money_result = int(total_buy_money)
            print(f"총매입금액: {total_buy_money_result}")

            total_profit_loss_rate = self.dynamicCall(
                "GetCommData(String, String, int, String",
                sTrCode,
                sRQName,
                0,
                "총수익률(%)",
            )
            total_profit_loss_rate_result = float(total_profit_loss_rate)
            print(f"총수익률(%): {total_profit_loss_rate_result}")

            self.detail_account_info_event_loop_2.exit()
