import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time
import pandas as pd
import sqlite3

TR_REQ_TIME_INTERVAL = 0.2

# 기본적인 함수 구성, 함수 이름 등은 "https://wikidocs.net/2814" - "파이썬으로 배우는 알고리즘 트레이딩" 을 참고하여 제작되었습니다. 좋은 강의 감사합니다 #

class kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        # Log in
        self._create_kiwoom_instance()
        # Log in event
        self._set_signal_slots()
        self.ohlcv = {'date': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}
        self.opw00018_output = {'single': [], 'multi': []}
        self.auto_order_output = []

    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._event_connect)
        self.OnReceiveTrData.connect(self._receive_tr_data)
        self.OnReceiveChejanData.connect(self._receive_chejan_data)

    def comm_connect(self):
        self.dynamicCall("CommConnect()")
        # 직접 eventloop 생성
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()

    def get_code_list_by_market(self, market):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market)
        code_list = code_list.split(';')
        return code_list[:-1]

    def get_master_code_name(self, code):
        code_name = self.dynamicCall("GetMasterCodeName(QString)", code)
        return code_name

    def get_connect_state(self):
        ret = self.dynamicCall("GetConnectState()")
        return ret

    def set_input_value(self, id, value):
        self.dynamicCall("SetInputValue(QString, QString)", id, value)

    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen_no)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    # 멀티데이터 개수 얻기
    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def get_comm_data(self, trcode, rqname, index, item):
        ret = self.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, index, item)
        return ret.strip()

    def _receive_tr_data(self, screen_no, rqname, trcode, recordname, next, unused1, unused2, unused3, unused4):
        if next == '2':
            self.remained_data = True
        else:
            self.remained_data = False

        if rqname == "opt10081_req":
            self._opt10081(rqname, trcode)

        elif rqname == "opw00001_req":
            self._opw00001(rqname, trcode)

        elif rqname == "opw00018_req":
            self._opw00018(rqname, trcode)

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    def _opt10081(self, rqname, trcode):
        data_cnt = self._get_repeat_cnt(trcode, rqname)

        for i in range(data_cnt):
            date = self.get_comm_data(trcode, rqname, i, "일자")
            open = self.get_comm_data(trcode, rqname, i, "시가")
            high = self.get_comm_data(trcode, rqname, i, "고가")
            low = self.get_comm_data(trcode, rqname, i, "저가")
            close = self.get_comm_data(trcode, rqname, i, "현재가")
            volume = self.get_comm_data(trcode, rqname, i, "거래량")

            self.ohlcv['date'].append(date)
            self.ohlcv['open'].append(int(open))
            self.ohlcv['high'].append(int(high))
            self.ohlcv['low'].append(int(low))
            self.ohlcv['close'].append(int(close))
            self.ohlcv['volume'].append(int(volume))

    def _opw00001(self, rqname, trcode):
        d2_deposit = self.get_comm_data(trcode, rqname, 0, "d+2추정예수금")
        self.d2_deposit = kiwoom.change_form(d2_deposit)

    def reset_opw00018_output(self):
        self.opw00018_output = {'single': [], 'multi': []}

    def reset_auto_order_output(self):
        self.auto_order_output = []

    def _opw00018(self, rqname, trcode):
        #single data
        total_purchase_price = self.get_comm_data(trcode, rqname, 0, "총매입금액")
        total_eval_price = self.get_comm_data(trcode, rqname, 0, "총평가금액")
        total_eval_profit_loss_price = self.get_comm_data(trcode, rqname, 0, "총평가손익금액")
        total_earning_rate = self.get_comm_data(trcode, rqname, 0, "총수익률(%)")
        estimated_deposit = self.get_comm_data(trcode, rqname, 0, "추정예탁자산")

        self.opw00018_output['single'].append(kiwoom.change_form(total_purchase_price))
        self.opw00018_output['single'].append(kiwoom.change_form(total_eval_price))
        self.opw00018_output['single'].append(kiwoom.change_form(total_eval_profit_loss_price))
        # total_earning_rate = kiwoom.change_form(total_earning_rate)
        # if self.get_server_gubun():
        #     total_earning_rate = float(total_earning_rate) / 100
        #     total_earning_rate = str(total_earning_rate)
        self.opw00018_output['single'].append(kiwoom.change_form(total_earning_rate))
        self.opw00018_output['single'].append(kiwoom.change_form(estimated_deposit))

        rows = self._get_repeat_cnt(trcode, rqname)
        for i in range(rows):
            name = self.get_comm_data(trcode, rqname, i, "종목명")
            quantity = self.get_comm_data(trcode, rqname, i, "보유수량")
            purchase_price = self.get_comm_data(trcode, rqname, i, "매입가")
            current_price = self.get_comm_data(trcode, rqname, i, "현재가")
            eval_profit_loss_price = self.get_comm_data(trcode, rqname, i, "평가손익")
            earning_rate = self.get_comm_data(trcode, rqname, i, "수익률(%)")

            quantity = kiwoom.change_form(quantity)
            purchase_price = kiwoom.change_form(purchase_price)
            current_price = kiwoom.change_form(current_price)
            eval_profit_loss_price = kiwoom.change_form(eval_profit_loss_price)
            earning_rate = kiwoom.change_form2(earning_rate)

            self.opw00018_output['multi'].append([name, quantity, purchase_price, current_price, eval_profit_loss_price, earning_rate])
            self.auto_order_output.append([name, quantity, earning_rate])

    def send_order(self, rqname, screen_no, account_no, ordertype, code, qty, price, hoga, orgorder_no):
        self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", [rqname, screen_no, account_no, ordertype, code, qty, price, hoga, orgorder_no])
        self.order_event_loop = QEventLoop()
        self.order_event_loop.exec_()

    def get_chejan_data(self, fid):
        ret = self.dynamicCall("GetChejanData(int)", fid)
        return ret

    def _receive_chejan_data(self, gubun, item_cnt, fidlist):
        self.com_gubun = gubun
        print(self.com_gubun)
        print(self.get_chejan_data(9203))
        print(self.get_chejan_data(302))
        print(self.get_chejan_data(900))
        print(self.get_chejan_data(901))

    # 계좌 정보 출력 함수
    def get_login_info(self, tag):
        ret = self.dynamicCall("GetLoginInfo(QString)", tag)
        return ret

    # 실 서버와 모의투자 서버에 제공되는 데이터의 형식이 다르다. 그것을 구분하기 위한 함수
    def get_server_gubun(self):
        ret = self.dynamicCall("KOA_Functions(QString, QString)", "GetServerGubun", "")
        return ret

    @staticmethod
    def change_form(data):
        strip_data = data.lstrip('-0')
        if strip_data == '':
            strip_data = 0

        # try : 보유량, 매입가, 현재가 등을 제대로 반환하기 위한 코드.
        try:
            format_data = format(int(strip_data), ',d')
        # except : 수익률 등의 float 값을 제대로 반환하기 위한 코드?
        except:
            format_data = format(float(strip_data))

        if data.startswith('-'):
            format_data = '-' + format_data

        return format_data

    # 수익률 변환 함수
    @staticmethod
    def change_form2(data):
        strip_data = data.lstrip('-0')

        if strip_data == '':
            strip_data = 0

        if strip_data.startswith('.'):
            strip_data = '0' + strip_data

        if data.startswith('-'):
            strip_data = '-' + strip_data

        return strip_data
