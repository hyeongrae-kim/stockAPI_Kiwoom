from PyQt5.QAxContainer import *  # PyQt5에서 제공하는 api 제어하기위한  내용 파이썬으로 가져오기 위한 라이브러리
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import *

# OCX방식의 컴포넌트 객체를 설치한것. 응용프로그램에서 키움Open Api를 실행할 수 있게 한 것. 제어가 가능.

class Kiwoom(QAxWidget):  # .QAxContainer내의 QAxWidget 상속
    def __init__(self):
        super().__init__()

        print("Kiwoom 클래스 입니다.")

        ######eventloop 모음
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()
        ########################

        ########스크린 번호 모음
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"

        ###########변수모음
        self.account_num = None
        self.account_stock_dict = {}
        self.not_account_stock_dict = {}


        ###종목 분석 용
        self.calcul_data = []

        ######################계좌 관련 변수
        self.use_money = 0
        self.use_money_percent = 0.5

        self.get_ocx_instance()  # 키움 api 제어하기 위한 초기화
        self.event_slots()  # 이벤트 모아 둔 것들 실행

        self.signal_login_commConnect()  # 자동 로그인 시도
        self.get_account_info()  # 계좌번호 가져오기
        self.detail_account_info()  # 예수금 가져오는 곳
        self.detail_account_mystock()  # 계좌평가 잔고내역 요청
        self.not_concluded_account()  # 미체결 내역 불러오기

        self.calculator_fnc()  # 종목분석용 임시실행


    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # QAxWidget 내 함수. 응용프로그램 내 키움 api 제어
                                                        # 레지스트리에 등록되어있는 키움 Api = KHOPENAPI.KHOpenAPICtrl.1

    def event_slots(self):  # 이벤트 모아 둘 구역
        self.OnEventConnect.connect(self.login_slot)  # 로그인 처리 이벤트, errCode 반환 하고 login_slot에 결과값 저장
        self.OnReceiveTrData.connect(self.trdata_slot)  # tr data 받는 이벤트

    def login_slot(self, errCode):
        print(errors(errCode))
        self.login_event_loop.exit()

    def signal_login_commConnect(self):  # commConnect = 수동 로그인 설정된경우 로그인 창 출력해서 로그인 시도 or 자동로그인일 경우 로그인 시도
        self.dynamicCall("CommConnect()")  # dynamicCall = PyQt에서 제공하는 함수, 네트워크 or 다른서버에 데이터 전송하도록 함
        self.login_event_loop = QEventLoop()  # QtCore내 QEventLoop, 로그인 핸들값 오류 해결
        self.login_event_loop.exec_()  # 로그인이 완료될때까지 다음 코드 실행안됨

    def get_account_info(self):  # 계좌번호 가져오기
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(';')[0]
        print("나의 보유 계좌번호 ", self.account_num)

    def detail_account_info(self):
        print("예수금 요청하는 부분")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)  # OpenAPI 조회함수 입력값 설정
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, String, int, String)", "예수금상세현황요청", "opw00001", "0", "self.screen_my_info")  # 조회함수를 호출해서 전문을 서버로 전송
        # 스크린번호? 그룹화를 해줌. 한개당 100개씩 요청내용 그룹 가능, 스크린번호는 200개 가능
        #  위에 줄에서는 3번째 인자가 스크린번호
        # 스크린안에 한개의 종목에 대한 내용을 요청받지 않겠다 = SetRealRemove(QString, QString, ), "스크린번호", "스크린번호 내 내용"
        # 스크린 날림 = DisconnectRealData(QString), 스크린번호

        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        print("계좌평가 잔고내역 요청")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)  # OpenAPI 조회함수 입력값 설정
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, "self.screen_my_info")

        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext="0"):
        print("미체결 요청")
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)  # OpenAPI 조회함수 입력값 설정
        self.dynamicCall("SetInputValue(String, String)", "체결구분", "1")
        self.dynamicCall("SetInputValue(String, String)", "매매구분", "0")
        self.dynamicCall("CommRqData(String, String, String, int, String)", "실시간미체결요청", "opt10075", sPrevNext,"self.screen_my_info")

        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr요청을 받는 구역이다! 슬롯
        :param sScrNo: 스크린 번호
        :param sRQName: 요청했을 때 지은 이름
        :param sTrCode: 요청 id, tr코드
        :param sRecordName: 사용 안함
        :param sPrevNext: 다음 페이지가 있는지
        :return:
        '''

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            deposit = int(deposit)
            print("에수금: ", deposit)

            self.use_money = deposit * self.use_money_percent
            self.use_money = self.use_money / 4

            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print(ok_deposit)

            self.detail_account_info_event_loop.exit()

        elif sRQName =="계좌평가잔고내역요청":  # 총 20개만 보여줌

            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money = int(total_buy_money)

            print("총 매입금액: ", total_buy_money)

            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate = float(total_profit_loss_rate)

            print("총 수익률: ", total_profit_loss_rate)

            # GetRepeatCnt = 멀티데이터 조회 용도, GetCommData = 싱글데이터 조회 용도도
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)  # 보유종목 불러오기
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]

                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")

                if code in self.account_stock_dict:
                    pass
                else:
                   self.account_stock_dict.update({code: {}})

                code_nm = code_nm.strip()
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_buy_price = int(total_buy_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_buy_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})

                cnt += 1

            print("계좌에 가지고 있는 종목", cnt)

            if sPrevNext == "2":  # 다음페이지가 있을 때
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        if sRQName =="실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)  # 보유종목 불러오기

            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태")  # 접수, 확인, 체결
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분")  #매도, 매수
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")

                code = code.strip()
                code_nm = code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}

                self.not_account_stock_dict[order_no].update({"종목코드": code})
                self.not_account_stock_dict[order_no].update({"종목명": code_nm})
                self.not_account_stock_dict[order_no].update({"주문번호": order_no})
                self.not_account_stock_dict[order_no].update({"주문상태": order_status})
                self.not_account_stock_dict[order_no].update({"주문수량": order_quantity})
                self.not_account_stock_dict[order_no].update({"주문가격": order_price})
                self.not_account_stock_dict[order_no].update({"주문구분": order_gubun})
                self.not_account_stock_dict[order_no].update({"미체결수량": not_quantity})
                self.not_account_stock_dict[order_no].update({"체결량": ok_quantity})

                print("미체결 종목: ", self.not_account_stock_dict[order_no])

            self.detail_account_info_event_loop.exit()

        if sRQName == "주식일봉차트조회":
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()
            print(code, "일봉데이터 요청")

            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            print("데이터일수 ", cnt)

            # data = self.dynamicCall("GetCommDataEx(QString, QString)", sTrCode, sRQName) 이 방식쓰면 아래와같이 나옴
            # [[ '', '현재가', '거래량', '거래대금', '날짜', '시가', '고가', '저가', ''][]]

            # 한번 조회하면 600일치까지 일봉데이터를 받을 수 있다.
            for i in range(cnt):
                data = []
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"거래량")
                trading_value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래대금")
                date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "일자")
                start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")

                data.append("")  # 형식맞춰주기 위함. 필수 x
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append("")

                self.calcul_data.append(data.copy())

            print(self.calcul_data)


            if sPrevNext == "2":
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                print("총 일수: ", len(self.calcul_data))

                pass_success = False
                # 120일 이평선을 그릴만큼의 데이터가 있는지 체크
                if self.calcul_data == None or len(self.calcul_data) <120:
                    pass_success = False
                else:  # 120일 이상이 되면은
                    total_price = 0
                    for value in self.calcul_data[:120]:
                        total_price += int(value[1])  # value[1] == 현재가, total_price에 현재가(종가) 다 더하기

                    moving_average_price = total_price/120

                    if int(self.calcul_data[0][7]) <= moving_average_price and moving_average_price <= int(self.calcul_data[0][6]):  # [0][7] == 오늘 일자의 최저가 [0][6] == 고가
                        print("오늘 주가 120이평선에 걸쳐있는 것 확인")

                self.calculator_event_loop.exit()


    def get_code_list_by_market(self, market_code):
        '''
        종목 코드를 반환하는 코드
        :param market_code:
        :return:
        '''
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(";")[:-1]

        return code_list

    def calculator_fnc(self):
        '''
        종목분석용 실행 함수
        :return:
        '''

        code_list = self.get_code_list_by_market("10")  # 10은 코스닥 코드
        print("코스닥 갯수 ", len(code_list))

        for idx,  code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock)  # screen number 끊어줌, 필수아님 오류방지
            print(idx+1, " / ", len(code_list), " : KOSDAQ Stock Code:", code, " is updating...")
            self.day_kiwoom_db(code=code)

    def day_kiwoom_db(self, code=None, date=None, sPrevNext="0"):  # signal 요청
        QTest.qWait(3600)  # 오류방지 딜레이, event process를 멈추지 않고 딜레이를 줌

        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", 1)

        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QString, QString, int ,QString)", "주식일봉차트조회", "opt10081", sPrevNext, self.screen_calculation_stock)

        if not sPrevNext == "2":
            self.calculator_event_loop.exec_()
