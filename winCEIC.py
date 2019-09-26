import os
import re
import sys
import datetime
import requests
import xlsxwriter
import fix_qt_import_error
import res_rc
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow
from bs4 import BeautifulSoup

from mainWindow import *


class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setWindowIcon(QtGui.QIcon(':res/xxx.ico'))
        self.setupUi(self)
        self.isGettingData = False
        self.Data = []
        self.name = 'default'
        self.filename = 'default'
        self.startD = ''
        self.endD = ''
        self.type = ''

    def getData(self, URL, startDate, endDate, type):
        self.Output.setPlainText('')
        self.isGettingData = True
        Data = []
        iURL = URL
        s = startDate.split('-')
        e = endDate.split('-')
        n = int(0)
        sy = int(s[0])
        sm = int(s[1])
        sd = int(s[2])
        ey = int(e[0])
        em = int(e[1])
        ed = int(e[2])
        if type == 0:
            n = ey - sy
        elif type == 1:
            n = em - sm
        elif type == 2:
            n = ed - sd
        times = n // 10 + 1
        for i in range(times):
            isDate = str(sy).zfill(4) + '-' + str(sm).zfill(2) + '-' + str(sd).zfill(2)
            iURL = re.sub(r'from=....-..-..', 'from=' + isDate, iURL, 1)
            if i == times - 1:
                iURL = re.sub(r'to=....-..-..', 'to=' + endDate, iURL, 1)
            else:
                if type == 0:
                    sy += 9
                if type == 1:
                    sm += 9
                if type == 2:
                    sd += 9
                ieDate = str(sy).zfill(4) + '-' + str(sm).zfill(2) + '-' + str(sd).zfill(2)
                iURL = re.sub(r'to=....-..-..', 'to=' + ieDate, iURL, 1)
            response = requests.get(url=iURL)
            ieDate = str(sy).zfill(4) + '-' + str(sm).zfill(2) + '-' + str(sd).zfill(2)
            if i == times - 1:
                ieDate = endDate
            self.Output.appendPlainText('已获取' + isDate + '到' + ieDate + '的数据')
            soup = BeautifulSoup(response.text, 'html.parser')
            iData = soup.find_all(class_='highcharts-text-outline')
            j = 0
            for j in range(len(iData)):
                iData[j] = float(iData[j].string.replace(' ', ''))
                Data.append(iData[j])
            if len(iData) < 10 and i != times - 1:
                self.Output.appendPlainText('警告！该段数据有缺失')
            if len(iData) < n % 10 and i == times - 1:
                self.Output.appendPlainText('警告！该段数据有缺失')
            if type == 0:
                sy += 1
            if type == 1:
                sm += 1
            if type == 2:
                sd += 1
        return Data

    def gdBtn(self):
        try:
            url = self.URL.toPlainText()
            if url == '':
                self.Output.setPlainText('请输入URL！')
                return
            self.startD = self.startDate.date().toString(Qt.ISODate)
            self.endD = self.endDate.date().toString(Qt.ISODate)
            self.type = 0
            if self.yearly.isChecked():
                self.type = 0
            elif self.monthly.isChecked():
                self.type = 1
            elif self.daily.isChecked():
                self.type = 2
            self.Data = self.getData(url, self.startD, self.endD, self.type)
            return
        except:
            self.Output.appendPlainText('获取数据时发生未知错误！')

    def checkFileName(self):
        if os.path.exists(self.filename + '.xlsx'):
            self.filename += '0'
            self.checkFileName()

    def geBtn(self):
        try:
            if not self.Data:
                self.Output.setPlainText('请先获取数据！')
                return
            self.name = self.Dataname.text()
            self.filename = self.Dataname.text()
            self.checkFileName()
            workbook = xlsxwriter.Workbook(self.filename + '.xlsx')
            worksheet = workbook.add_worksheet(self.name)
            title_format = workbook.add_format({'bold': True})
            yearly_format = workbook.add_format({'num_format': 'yyyy'})
            monthly_format = workbook.add_format({'num_format': 'yyyy-mm'})
            daily_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
            worksheet.write('A1', '时间', title_format)
            worksheet.write('B1', self.name, title_format)
            for i in range(len(self.Data)):
                worksheet.write(i + 1, 1, self.Data[i])
                if self.type == 0:
                    worksheet.write(i + 1, 0, str(datetime.datetime.strptime(self.startD, '%Y-%m-%d').date().year + i), yearly_format)
                if self.type == 1:
                    worksheet.write(i + 1, 0, '{0}-{1}'.format(
                        str(datetime.datetime.strptime(self.startD, '%Y-%m-%d').date().year),
                        str(datetime.datetime.strptime(self.startD, '%Y-%m-%d').date().month + i)), monthly_format)
                if self.type == 2:
                    worksheet.write(i + 1, 0, '{0}-{1}-{2}'.format(
                        str(datetime.datetime.strptime(self.startD, '%Y-%m-%d').date().year),
                        str(datetime.datetime.strptime(self.startD, '%Y-%m-%d').date().month),
                        str(datetime.datetime.strptime(self.startD, '%Y-%m-%d').date().day + i)), daily_format)
            workbook.close()
            self.Output.appendPlainText('成功生成' + self.filename + '.xlsx')
            return
        except:
            self.Output.appendPlainText('生成时发生未知错误！')

    def okBtn(self):
        self.gdBtn()
        self.geBtn()
        return




app = QApplication(sys.argv)
myWin = MyWindow()
myWin.GetData.clicked.connect(myWin.gdBtn)
myWin.Generate.clicked.connect(myWin.geBtn)
myWin.oneKey.clicked.connect(myWin.okBtn)
myWin.show()
sys.exit(app.exec_())
