from kiwoom.kiwoom import *

import sys  # 파이썬 시스템 라이브러리
from PyQt5.QtWidgets import *  # PyQt5 = ui꾸밀때 사용

class UI_class():
    def __init__(self):
        print("UI 클래스 입니다.")

        self.app = QApplication(sys.argv)  # PyQt5내 QApplication 클래스, ui 실행하기위한 함수, 변수 초기화
                                            # sys.argv 파일들을 사용하여 app용도로 사용
                                            # sys.argv엔 파이썬 파일 경로 담겨져있음

        self.kiwoom = Kiwoom()
        # Kiwoom()이대로 실행해도 됨

        self.app.exec_()  # 프로그램 종료 막음