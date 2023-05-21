from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.error_code import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Class Kiwoom")

        # Event Loop
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        ############################

        # Variables
        self.account_num = None
        self.use_money = 0
        self.use_money_percent = 0.5
        self.account_stock_dict = {}
        self.not_account_stock_dict = {}

        # Screen Number
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()  # 예수금 요청
        self.detail_account_mystock()  # 계좌평가잔고내역요청
        self.not_concluded_account()  # 실시간미체결요청
        self.calculator_fnc()  # 종목 분석용 임시로 실행

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
            self.screen_my_info,
        )

        # 요청 후 event loop를 실행한 상태에서 TR 요청 데이터가 올 때까지 대기
        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        print(f"계좌평가 잔고내역 요청하기 연속 조회 {sPrevNext}")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall(
            "CommRqData(String, String, int, String)",
            "계좌평가잔고내역요청",
            "opw00018",
            sPrevNext,
            self.screen_my_info,
        )

        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall(
            "CommRqData(QString, QString, int, QString)",
            "실시간미체결요청",
            "opt10075",
            sPrevNext,
            self.screen_my_info,
        )

        self.detail_account_info_event_loop.exec_()

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

            # 계좌 평가 잔고 내역에서 종목 가져오기
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            cnt = 0

            for i in range(rows):
                code = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "종목번호",
                )
                code = code.strip()[1:]

                code_name = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "종목명",
                )
                stock_quantity = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "보유수량",
                )
                buy_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "매입가",
                )
                learn_rate = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "수익률(%)",
                )
                current_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "현재가",
                )
                total_chegual_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "매입금액",
                )
                possible_quantity = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "매매가능수량",
                )

                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict.update({code: {}})

                # string으로 나온 값을 int로 변환
                code_name = code_name.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명": code_name})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})

                cnt += 1

            print(f"계좌에 가지고 있는 종목: {self.account_stock_dict}")
            print(f"계좌에 가지고 있는 종목 개수: {cnt}")

            # 계좌 평가 잔고 내역은 한 페이지에 20개만 조회 가능해서 다음 페이지가 있는지 확인
            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            for i in range(rows):
                code = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "종목번호",
                )
                code_name = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "종목명",
                )
                order_no = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "주문번호",
                )
                order_status = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "주문상태",
                )
                order_quantity = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "주문수량",
                )
                order_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "주문가격",
                )
                order_gubun = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "주문구분",
                )
                not_quantity = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "미체결수량",
                )
                ok_quantity = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)",
                    sTrCode,
                    sRQName,
                    i,
                    "체결량",
                )

                code = code.strip()
                code_name = code_name.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip("+").lstrip("-")
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}

                self.not_account_stock_dict[order_no].update({"종목코드": code})
                self.not_account_stock_dict[order_no].update({"종목명": code_name})
                self.not_account_stock_dict[order_no].update({"주문번호": order_no})
                self.not_account_stock_dict[order_no].update({"주문상태": order_status})
                self.not_account_stock_dict[order_no].update({"주문수량": order_quantity})
                self.not_account_stock_dict[order_no].update({"주문가격": order_price})
                self.not_account_stock_dict[order_no].update({"주문구분": order_gubun})
                self.not_account_stock_dict[order_no].update({"미체결수량": not_quantity})
                self.not_account_stock_dict[order_no].update({"체결량": ok_quantity})

                print(f"미체결 종목: {self.not_account_stock_dict[order_no]}")

            self.detail_account_info_event_loop.exit()

    def get_code_list_by_market(self, market_code):
        """
        종목 코드 반환
        :param market_code:
        :return:
        """
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]

        return code_list

    def calculator_fnc(self):
        """
        종목 분석 실행용 함수
        :return:
        """
        code_list = self.get_code_list_by_market("10")
        print(f"코스닥 갯수: {len(code_list)}")

    def day_kiwoom_db(self, code=None, date=None, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")

        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall(
            "CommRqData(QString, QString, int, QString)",
            "주식일봉차트조회",
            "opt10081",
            sPrevNext,
            self.screen_calculation_stock,
        )
