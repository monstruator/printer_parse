from PyQt5 import QtCore, QtGui, QtWidgets

import sys
import os
from PyQt5.Qt import *
import win32print
import win32api
import time
import argparse

from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_FastPrint(object):
    def setupUi(self, FastPrint):
        FastPrint.setObjectName("FastPrint")
        FastPrint.resize(220, 161)
        self.centralwidget = QtWidgets.QWidget(FastPrint)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 30, 81, 16))
        self.label.setObjectName("label")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(20, 50, 113, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(20, 110, 75, 23))
        self.pushButton.setObjectName("pushButton")
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setEnabled(True)
        self.checkBox.setGeometry(QtCore.QRect(20, 80, 141, 18))
        self.checkBox.setObjectName("checkBox")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(20, 10, 221, 16))
        self.label_2.setObjectName("label_2")
        FastPrint.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(FastPrint)
        self.statusbar.setObjectName("statusbar")
        FastPrint.setStatusBar(self.statusbar)

        self.retranslateUi(FastPrint)
        QtCore.QMetaObject.connectSlotsByName(FastPrint)

    def retranslateUi(self, FastPrint):
        _translate = QtCore.QCoreApplication.translate
        FastPrint.setWindowTitle(_translate("FastPrint", "[AB] FastPrint (pages)"))
        self.label.setText(_translate("FastPrint", "Выбор страниц"))
        self.pushButton.setText(_translate("FastPrint", "Печать"))
        self.checkBox.setText(_translate("FastPrint", "Двусторонняя печать"))
        self.label_2.setText(_translate("FastPrint", "Файл не определен"))



class mywindow(QtWidgets.QMainWindow):
    GHOSTSCRIPT_PATH = "C:\\Program Files\\WinFast\\gs\\bin\\gswin64c.exe"
    GSPRINT_PATH = "C:\\Program Files\\WinFast\\Ghostgum\\gsview\\gsprint.exe"
    find_printer = 0
    right_printer = ''
    rp = ' '
    parser = None
    namespace = None
    duplex =  " "

    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_FastPrint()
        self.ui.setupUi(self)
        self.ui.checkBox.stateChanged.connect(self.checkedc)
        self.ui.pushButton.clicked.connect(self.to_printer)  

        self.parser = self.createParser()
        self.namespace = self.parser.parse_args()
        #self.ui.pushButton.setAutoDefault(True)  
        #self.ui.lineEdit.returnPressed.connect(self.ui.pushButton.click)
        if self.namespace.name:
            self.ui.label_2.setText("")
            self.find_printer = 0
        all_printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        for i in all_printers:
            if self.rp in i:
                print("OK")
                self.right_printer = i
                self.find_printer = 1
        if self.find_printer == 0:
            self.ui.label_2.setText("Принтер не найден")#self.namespace.name
        


    def keyPressEvent(self, event):
        key = event.key()
        print('Нажали: ' , str(event.key()), QtCore.Qt.Key_Return)  
        #print(app.arguments())
        if key == QtCore.Qt.Key_Return or key == QtCore.Qt.Key_Enter: 
            print("Enter")
            self.to_printer()
        if key == QtCore.Qt.Key_Escape:
            print("Esc")
            self.close()
            
    def createParser(self):
        parser = argparse.ArgumentParser()
        parser.add_argument('name', nargs='?')
        return parser

    def to_printer(self):
        print("to priner")
        if self.namespace.name:
            print(self.namespace.name)
            str1 = str(self.namespace.name)

            if self.find_printer == 1:
                printer_name = self.right_printer
                printer_srt = '" -printer "' + printer_name + '" '
                file_str = '"' + str1 + '"'
                pages = self.ui.lineEdit.text()
                if pages:
                    pages =  ' -sPageList=' + pages + ' ' 
                else:
                    pages = " "
                Str = '-ghostscript "' + self.GHOSTSCRIPT_PATH + printer_srt + self.duplex + ' -dBATCH -dNOPAUSE ' + pages + file_str + " -color" 
                print(Str)
                win32api.ShellExecute(0, 'open', self.GSPRINT_PATH, Str, '.', 0)
        self.close()

    def checkedc(self, checked):
        if checked:
            print("checked")
            self.duplex =  " -duplex_vertical "
        else:
            print("not checked")
            self.duplex =  " "
        self.show()
 
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1" 
app = QtWidgets.QApplication(sys.argv)
application = mywindow()
application.show()
 
sys.exit(app.exec())