from PyQt5 import QtCore, QtGui, QtWidgets
from fast_gui import Ui_FastPrint  # импорт нашего сгенерированного файла
import sys
import os
from PyQt5.Qt import *
import win32print
import win32api
import time
import argparse
 
class mywindow(QtWidgets.QMainWindow):
    GHOSTSCRIPT_PATH = "C:\\Program Files\\WinFast\\gs\\bin\\gswin64c.exe"
    GSPRINT_PATH = "C:\\Program Files\\WinFast\\Ghostgum\\gsview\\gsprint.exe"
    find_printer = 0
    right_printer = ''
    rp = ' '
    parser = None
    namespace = None
    duplex =  " -dDuplex=false"

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
            self.duplex =  " -dDuplex=true"
        else:
            print("not checked")
            self.duplex =  " -dDuplex=false"
        self.show()
 
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1" 
app = QtWidgets.QApplication(sys.argv)
application = mywindow()
application.show()
 
sys.exit(app.exec())