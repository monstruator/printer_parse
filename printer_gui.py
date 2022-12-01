from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QTableWidgetItem
from wind import Ui_MainWindow  # импорт нашего сгенерированного файла
from screeninfo import get_monitors
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
 

class BrowserHandler(QtCore.QObject):
    running = False
    newText = QtCore.pyqtSignal(str, object)


 
class mywindow(QtWidgets.QMainWindow):
    url = 'http://192.168.0.73/scan.htm'
    printer_name = 'Xerox WorkCentre 7345'
    driver = None
    data_files = []
    n_file = 0 #номер файла на мечать

    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        

        # for m in get_monitors():
        #     # str(m))
        #     self.window_width = m.width/3
        #     self.window_height = m.height/2
        
        # self.setFixedWidth(self.window_width)
        # self.setFixedHeight(self.window_height)
        # self.ui.label.setGeometry(QtCore.QRect(10, 150, 300, 180)) #позиция
        # self.ui.label.setFont(QtGui.QFont('SansSerif', 20))

        # self.ui.lineEdit.setGeometry(QtCore.QRect(30, self.window_height - 70, 90, self.window_height-10))
        # self.ui.lineEdit.resize(60, 30)
        # self.ui.lineEdit.setFont(QtGui.QFont('SansSerif', 14))
        # self.ui.lineEdit.setMaxLength(3)

        options = webdriver.ChromeOptions() 
        prefs = {"download.default_directory" : "D:\Pavel"}
        options.add_experimental_option("prefs",prefs)
        options.add_argument("--headless")

        self.driver = webdriver.Chrome(executable_path="D:\\Pavel\\chromedriver\\chromedriver.exe",chrome_options=options)

        # подключение клик-сигнал к слоту btnClicked
        self.ui.pushButton.clicked.connect(self.btnClicked)
        self.ui.pushButton_2.clicked.connect(self.update_file_list)
        self.ui.pushButton_3.clicked.connect(self.print_file)
        self.ui.tableWidget.cellClicked.connect(self.tab_click)

        self.ui.tableWidget.setRowCount(7)
        self.ui.tableWidget.setColumnCount(3)
        self.ui.label_2.setVisible(False)
        self.ui.label_3.setVisible(False)

    def btnClicked(self):
        self.ui.label.setText("Ожидание принтера ...")
        res = self.connect_to_printer()
        if res == 1:
            self.ui.label.setText("Принтер обнаружен")
        else:
            self.ui.label.setText("Принтер не обнаружен")    


    def update_file_list(self):    
        print('update_file_list')
        # self.ui.tableWidget.clear()

        self.list_files()    
        print(len(self.data_files))

        if len(self.data_files) > 0:
            self.ui.tableWidget.setRowCount(len(self.data_files))
            row = 0
            for tup in self.data_files:
                col = 0
                for item in tup:
                    cellinfo = QTableWidgetItem(item)
                    self.ui.tableWidget.setItem(row, col, cellinfo)
                    col += 1

                row += 1
            self.ui.label_3.setVisible(False)
        else:
            self.ui.label_3.setVisible(True)    

    def tab_click(self):
        self.n_file = self.ui.tableWidget.currentRow() + 1


    def print_file(self):
        if (len(self.data_files)) and self.n_file > 0:
            try:
                str1 = '/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(self.n_file + 1) + ']/td[1]/input'
                check_box1 = self.driver.find_element(By.XPATH, str1).click()
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input").click()
                file1 = self.driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td/small/a").click()
                self.driver.back()
                time.sleep(2)
                iframe1 = self.driver.find_element(By.NAME,"RF")
                self.driver.switch_to.frame(iframe1)
                check_box1 = self.driver.find_element(By.XPATH, str1).click()
                #return 1
            except Exception as ex:
                print(ex)
                #return 0

            self.ui.label_2.setVisible(False)
        else:
            self.ui.label_2.setVisible(True)

    
    def connect_to_printer(self):
        try:
            print('START CONNECT')
            self.driver.get(url=self.url)
            #phtml = driver.page_source
            time.sleep(1) #2
            #print(phtml)
            print('search')
            iframe = self.driver.find_element(By.NAME,"NF")
            self.driver.switch_to.frame(iframe)
            self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
            self.driver.switch_to.default_content()
            iframe1 = self.driver.find_element(By.NAME,"RF")
            self.driver.switch_to.frame(iframe1)
            return 1
        except Exception as ex:
            print("ERROR CONNECT")
            #print(ex)
            return 0  


    def list_files(self):
        try:
            l = self.driver.find_elements(By.XPATH,"/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr")
            for el in l:
                if len(el.text) != 0:
                    if 'Документ номер Имя документа' not in el.text:
                        #print(el.text)
                        self.data_files.append(el.text)     
        except Exception as ex:
            print(ex)
            return None


os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
app = QtWidgets.QApplication([])
# app.setStyle('Windows')
application = mywindow()
application.show()

 
sys.exit(app.exec())