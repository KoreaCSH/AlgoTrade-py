import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from kiwoom import *
import time

form_class = uic.loadUiType("pytrader.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.kiwoom = kiwoom()
        self.kiwoom.comm_connect()
        self.set_name_list()
        self.stop_loss_rate = {}

        # 현재 시간 statusBar에 출력하기
        self.timer = QTimer(self)
        # start 메서드에 1000을 주면 1초마다 timeout 시그널이 발생한다.
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        # 계좌를 QComboBox 위젯에 출력
        accouns_num = int(self.kiwoom.get_login_info("ACCOUNT_CNT"))
        accounts = self.kiwoom.get_login_info("ACCNO")
        accounts_list = accounts.split(';')[0:accouns_num]
        self.comboBox.addItems(accounts_list)

        # 종목명을 종목코드로 변환
        self.lineEdit.textChanged.connect(self.name_changed)

        # 계좌정보조회
        self.check_balance()
        self.pushButton_3.clicked.connect(self.check_balance)

        # 실시간 계좌 조회를 위한 timer - 10초에 한 번
        self.timer2 = QTimer(self)
        self.timer2.start(1000*5)
        self.timer2.timeout.connect(self.timeout2)

        # 매매리스트 읽기
        self.getBuyTableItem()
        self.pushButton_5.clicked.connect(self.getBuyTableItem)

        # 매매 리스트에 종목 추가 및 종목 매수
        self.pushButton.clicked.connect(self.set_buylist)

        # 자동 매도 알고리즘 / 2/17에 테스트 해보기.
        self.stop_loss()

    def timeout(self):
        current_time = QTime.currentTime()
        # 시간:분:초 형태로 변환
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        state = self.kiwoom.get_connect_state()
        if state == 1:
            state_msg = "서버 연결 성공"
        else:
            state_msg = "서버 연결 실패"

        # statusBar 따로 선언하지 않아도 사용할 수 있는 것인가?
        self.statusbar.showMessage(state_msg + " | " + time_msg)

    def set_name_list(self):
        # name_list 만들기
        self.kospi_code_list = self.kiwoom.get_code_list_by_market("0")
        self.kosdaq_code_list = self.kiwoom.get_code_list_by_market("10")
        self.name_list = {}
        for code in self.kospi_code_list:
            self.name_list[code] = self.kiwoom.get_master_code_name(code)

        for code in self.kosdaq_code_list:
            self.name_list[code] = self.kiwoom.get_master_code_name(code)

        self.real_name_list = {v: k for k, v in self.name_list.items()}

    def name_changed(self):
        name = self.lineEdit.text()
        if name in self.real_name_list:
            code = self.real_name_list[name]
            self.lineEdit_2.setText(code)

    def send_order(self):
        order_type_lookup = {'신규매수': 1, '신규매도': 2, '매수취소': 3, '매도취소': 4}
        hoga_lookup = {'지정가': "00", '시장가': "03"}

        account_no = self.comboBox.currentText()
        ordertype = self.comboBox_2.currentText()
        code = self.lineEdit_2.text()
        qty = self.spinBox.value()
        price = self.spinBox_2.value()
        hoga = self.comboBox_3.currentText()

        self.kiwoom.send_order("send_order_req", "0101", account_no, order_type_lookup[ordertype], code, qty, price, hoga_lookup[hoga], "")

    def auto_send_order(self, name, qty):
        account_no = self.comboBox.currentText()
        code = self.real_name_list[name]
        quantity = qty

        self.kiwoom.send_order("send_order_req", "0101", account_no, 2, code, quantity, 0, "03", "")

    def check_balance(self):
        # pushButton_3 을 누를 때마다 opw00018_output을 reset 한다.
        self.kiwoom.reset_opw00018_output()
        account_number = self.kiwoom.get_login_info("ACCNO")
        account_number = account_number.split(';')[0]

        self.kiwoom.set_input_value("계좌번호", account_number)
        # 세 번째 인자 2는 연속데이터 조회 요청!
        self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")

        # remained_data 의 첫 상태는 TRUE 인가? 무조건? 지금은 remained data 가 없어서 그냥 실행된 듯 함.
        while self.kiwoom.remained_data:
            time.sleep(0.2)
            self.kiwoom.set_input_value("계좌번호", account_number)
            self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 2, "2000")

        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00001_req", "opw00001", 0, "2000")

        # balance
        item = QTableWidgetItem(self.kiwoom.d2_deposit)
        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.tableWidget.setItem(0, 0, item)

        for i in range(1, 6):
            item = QTableWidgetItem(self.kiwoom.opw00018_output['single'][i-1])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget.setItem(0, i, item)

        # item 크기에 맞춰 행 높이 조절하기
        self.tableWidget.resizeRowsToContents()

        item_count = len(self.kiwoom.opw00018_output['multi'])
        self.tableWidget_2.setRowCount(item_count)

        for j in range(item_count):
            row = self.kiwoom.opw00018_output['multi'][j]
            for i in range(len(row)):
                item = QTableWidgetItem(row[i])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_2.setItem(j, i, item)

        self.tableWidget_2.resizeRowsToContents()

    def timeout2(self):
        if self.checkBox.isChecked():
            self.check_balance()
            self.stop_loss()

    def set_buylist(self):
        f = open("buy_list.txt", 'rt')
        buy_list = f.readlines()
        f.close()

        if len(buy_list) == 0:
            code = self.lineEdit.text()
            term = self.comboBox_4.currentText()
            ex_earning_rate = self.spinBox_3.value()
            stop_loss_rate = self.spinBox_4.value()
            earning_rate = 10
            gubun = self.comboBox_2.currentText()

            if code == '' or term == '' or ex_earning_rate == '' or stop_loss_rate == '':
                # 추후 경고창으로 update
                print("code, term, ex_earning_rate 를 입력해주세요")

            else:
                f = open("buy_list.txt", 'wt')
                f.writelines("%s;%s;%s;%s;%s;%s" % (code, term, ex_earning_rate, stop_loss_rate, earning_rate, gubun))
                f.close()
                self.send_order()

        else:
            code = self.lineEdit.text()
            term = self.comboBox_4.currentText()
            ex_earning_rate = self.spinBox_3.value()
            stop_loss_rate = self.spinBox_4.value()
            earning_rate = 10
            gubun = self.comboBox_2.currentText()

            flag = 0

            for row_data in buy_list:
                split_row_data = row_data.split(';')
                if code in split_row_data:
                    flag = 1
                    break

            if code == '' or term == '' or ex_earning_rate == '' or stop_loss_rate == '':
                print("code, term, ex_earning_rate 를 입력해주세요")

            elif flag == 1:
                self.send_order()

            else:
                f = open("buy_list.txt", 'at')
                # 'a' 모드는 파일의 가장 끝으로 이동하므로 새로운 종목을 추가할 때 줄바꿈을 해 줘야 한다.
                f.write("\n")
                f.writelines("%s;%s;%s;%s;%s;%s" % (code, term, ex_earning_rate, stop_loss_rate, earning_rate, gubun))
                f.close()
                self.send_order()

        # if ~ else 문의 self.send_order() 이 끝이 나야 실행된다.

    def getBuyTableItem(self):
        f = open("buy_list.txt", 'rt')
        buy_list = f.readlines()
        f.close()

        rowCount = len(buy_list)
        self.tableWidget_3.setRowCount(rowCount)

        for j in range(len(buy_list)):
            row_data = buy_list[j]
            split_row_data = row_data.split(';')

            for i in range(len(split_row_data)):
                item = QTableWidgetItem(split_row_data[i].rstrip())
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignCenter)
                self.tableWidget_3.setItem(j, i, item)

        # 초기화 어느 위치에서 진행하는게 좋을까?
        self.stop_loss_rate = {}
        for j in range(len(buy_list)):
            row_data = buy_list[j]
            split_row_data = row_data.split(';')
            self.stop_loss_rate[split_row_data[0]] = [split_row_data[2], split_row_data[3]]

    def stop_loss(self):
        for rate_data in self.kiwoom.auto_order_output:
            ex_earning_rate = self.stop_loss_rate[rate_data[0]][0]
            stop_loss_rate = self.stop_loss_rate[rate_data[0]][1]
            earning_rate = rate_data[2]

            if float(earning_rate) < float(stop_loss_rate) or float(earning_rate) > float(ex_earning_rate):
                self.auto_send_order(rate_data[0], rate_data[1])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
