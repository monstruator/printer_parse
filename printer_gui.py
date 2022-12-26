from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QTableWidgetItem, QSystemTrayIcon, QSpacerItem, QSizePolicy, QMenu, QAction, QStyle, qApp, QFileDialog
#from wind import Ui_MainWindow  # импорт нашего сгенерированного файла
#from screeninfo import get_monitors
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select

import time
import os
import glob
import configparser
#---------------------------------------------------------------
from reportlab.pdfgen import canvas
import reportlab
from reportlab.lib.pagesizes import letter
import PIL
from PIL import Image
import sys
#----------------------------------------------------------------- 
#from tif_to_pdf2 import convert_with_auto_rotate
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(668, 451)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(10, 360, 141, 41))
        self.pushButton.setObjectName("pushButton")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(30, 5, 411, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(170, 360, 141, 41))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(470, 160, 171, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(20, 50, 431, 271))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        #self.tableWidget.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)
        #self.tableWidget.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setEnabled(True)
        self.label_2.setGeometry(QtCore.QRect(470, 330, 181, 20))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setEnabled(True)
        self.label_3.setGeometry(QtCore.QRect(230, 320, 121, 20))
        self.label_3.setObjectName("label_3")
        self.radioButton = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton.setGeometry(QtCore.QRect(460, 130, 101, 18))
        self.radioButton.setObjectName("radioButton")
        self.radioButton_2 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_2.setGeometry(QtCore.QRect(460, 100, 101, 18))
        self.radioButton_2.setChecked(True)
        self.radioButton_2.setObjectName("radioButton_2")
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setGeometry(QtCore.QRect(460, 10, 161, 20))
        self.checkBox.setChecked(False)
        self.checkBox.setObjectName("checkBox")
        self.checkBox_2 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_2.setGeometry(QtCore.QRect(460, 70, 191, 18))
        self.checkBox_2.setChecked(False)
        self.checkBox_2.setObjectName("checkBox_2")
        self.checkBox_3 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_3.setGeometry(QtCore.QRect(460, 40, 161, 18))
        self.checkBox_3.setChecked(False)
        self.checkBox_3.setObjectName("checkBox_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(590, 120, 61, 16))
        self.label_4.setObjectName("label_4")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(330, 380, 41, 21))
        self.lineEdit.setInputMethodHints(QtCore.Qt.ImhPreferNumbers)
        self.lineEdit.setInputMask("")
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(380, 380, 71, 23))
        self.pushButton_4.setObjectName("pushButton_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(330, 360, 131, 16))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(466, 360, 101, 16))
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(466, 380, 161, 16))
        self.label_7.setObjectName("label_7")
        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QtCore.QRect(630, 370, 31, 23))
        self.pushButton_5.setObjectName("pushButton_5")
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(470, 310, 181, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.pushButton_6 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_6.setGeometry(QtCore.QRect(470, 260, 171, 31))
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_7 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_7.setGeometry(QtCore.QRect(470, 200, 171, 31))
        self.pushButton_7.setObjectName("pushButton_7")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 668, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "Подключиться к принтеру"))
        self.label.setText(_translate("MainWindow", "TextLabel"))
        self.pushButton_2.setText(_translate("MainWindow", "Обновить список файлов"))
        self.pushButton_3.setText(_translate("MainWindow", "Скачать файл"))
        self.label_2.setText(_translate("MainWindow", "Файл не выбран"))
        self.label_3.setText(_translate("MainWindow", "Файлы не найдены"))
        self.radioButton.setText(_translate("MainWindow", "Горизонтальная"))
        self.radioButton_2.setText(_translate("MainWindow", "Вертикальная"))
        self.checkBox.setText(_translate("MainWindow", "удалять в почтовом ящике"))
        self.checkBox_2.setText(_translate("MainWindow", "удалять tif при конвертировании"))
        self.checkBox_3.setText(_translate("MainWindow", "конвертировать в pdf"))
        self.label_4.setText(_translate("MainWindow", "ориентация"))
        self.pushButton_4.setText(_translate("MainWindow", "Сохранить"))
        self.label_5.setText(_translate("MainWindow", "Номер почтового ящика"))
        self.label_6.setText(_translate("MainWindow", "Папка назначения"))
        self.label_7.setText(_translate("MainWindow", "TextLabel"))
        self.pushButton_5.setText(_translate("MainWindow", "..."))
        self.pushButton_6.setText(_translate("MainWindow", "Открыть папку"))
        self.pushButton_7.setText(_translate("MainWindow", "Удалить файл"))



def convert_with_auto_rotate(tiff, vert = 1): #по умолчанию vert = 1 (вертикальная ориентация картинки)
    try: 
        print(tiff)
        if vert != 0 and vert != 1:
            vert = 0
        img = PIL.Image.open(tiff)
        width, height = img.size
        scale = width / height
        scale = scale * 0.9
        if (vert == 1 and scale > 1) or (vert == 0 and scale < 1):
            rotate_img = 1
        else:
            rotate_img = 0             
        out = tiff.replace('.tif','.pdf')
        if rotate_img == 1:
            outPDF = canvas.Canvas(out, pageCompression=1, pagesize=(height, width), bottomup=1)
        if rotate_img == 0:
            outPDF = canvas.Canvas(out, pageCompression=1, pagesize=(width, height), bottomup=1)
        for page in range(img.n_frames):
            img.seek(page)
            if rotate_img == 1:
                rotate_image = img.rotate(90, expand=True)
                imgPage = reportlab.lib.utils.ImageReader(rotate_image)
                outPDF.drawImage(imgPage, 0, 0,  height, width)
            if rotate_img == 0:
                imgPage = reportlab.lib.utils.ImageReader(img)
                outPDF.drawImage(imgPage, 0, 0,  width, height)    
            if page < img.n_frames:
                outPDF.showPage()
        outPDF.save()
        img.close()
        return 1
    except Exception as ex:
        print(ex)
        return 0

class BrowserHandler(QtCore.QObject): #поток для длительных операций
    running = False
    toProgressBar = QtCore.pyqtSignal(int)
    toInit = QtCore.pyqtSignal(str, object, int)
    toError = QtCore.pyqtSignal(str)
    driver = None    
    doc_dir = None
    printer = 0
    str1 = '' #для передачи в главное окно
    n_box_mail = 1 #номер почтового ящика
    box_name = "None"
    data_files = []
    del_in_mail = 1 #удалять файл после скачивания
    del_tif = 1 #удалять tif при конвертации
    convert_to_tif = 1 #конвертировать в pdf
    vert = 1 #ориентация выходного документа
    ip = 'http://192.168.0.128/'
    printers = [
        {'url' : 'scan.htm',  'name' : 'Xerox WorkCentre 73xx'},
        {'url' : 'prop.htm', 'name' : 'Xerox WorkCentre 1xx'},
        {'url' : 'prtscn.htm', 'name' : 'Xerox WorkCentre 73(72)'},
        {'url' : 'scan.htm', 'name' : 'Xerox WorkCentre 5325'},
        {'url' : 'scan.htm', 'name' : 'Xerox WorkCentre 7120'},
        {'url' : 'scan/index.php', 'name' : 'Xerox WorkCentre 72xx'},
    ]
    
    def __init__(self, doc_dir, ip, n_box_mail, box_name):
        super(BrowserHandler, self).__init__() #
        print('start thread ', doc_dir)
        try:
            start_time = time.time()
            self.n_box_mail = n_box_mail
            self.ip = ip
            self.doc_dir = doc_dir
            self.box_name = box_name
            options = webdriver.ChromeOptions() 
            prefs = {"download.default_directory" : self.doc_dir}
            options.add_experimental_option("prefs",prefs)
            #--headless--headless--headless
            #options.add_argument("--headless")
            self.driver = webdriver.Chrome(chrome_options=options)
            self.driver.set_page_load_timeout(9)
            print(" 1------ %s sec -------" %( time.time() - start_time))
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка запуска драйвера")
            
    
    def run(self):
        try:
            start_time = time.time()
            print("IP=",self.ip)
            self.toProgressBar.emit(0)

            res = self.connect_to_printer()

            if res > 0:
                self.toProgressBar.emit(30)
                res1 = self.choice_mail_box()
                self.toProgressBar.emit(40)
                if res1 == 1:
                    self.list_files()
                print(self.printers[self.printer-1]['name'])
                self.str1 = self.printers[self.printer-1]['name']
                print('Обнаружен принтер ', self.printers[self.printer-1]['name'])
            else:
                print('Не обнаружен принтер ', pr['name'])
                self.str1 = "Принтер не обнаружен"
            
            self.toProgressBar.emit(100)        
            self.toInit.emit(self.str1, self.data_files, self.printer)
            print(" 2------ %s sec -------" %( time.time() - start_time))
        except:
            self.toError.emit("Ошибка поиска принтера")
        
 
    def connect_to_printer(self):
        try:
            print('START CONNECT')
            #self.driver.get(url=self.ip + printer['url'])
            self.driver.get(url=self.ip)
            print("TITLE: ", self.driver.title)
            if "5325" in self.driver.title:
                self.printer = 4
            if "7120" in self.driver.title:
                self.printer = 5
            if "XEROX WORKCENTRE PRO" in self.driver.title:
                self.printer = 6
                self.driver.get(url=self.ip + self.printers[self.printer-1]['url'])
                iframe = self.driver.find_element(By.NAME,"theHeader")
                self.driver.switch_to.frame(iframe)
                name = self.driver.find_element(By.ID,"productName")
                if "7535" in name.text:
                    self.printers[self.printer-1]['name'] = "Xerox WorkCentre 75xx"
                self.driver.switch_to.default_content()
                return self.printer
            self.driver.get(url=self.ip + self.printers[self.printer-1]['url'])
            iframe = self.driver.find_element(By.NAME,"NF")
            self.driver.switch_to.frame(iframe)
            self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
            self.driver.switch_to.default_content()
            iframe1 = self.driver.find_element(By.NAME,"RF")
            self.driver.switch_to.frame(iframe1)
            return self.printer
        except Exception as ex:
            print("ERROR CONNECT to ", printer['name'])
            print(ex)
            #self.toError.emit("Ошибка поиска принтера")
            return 0  
            
    def choice_mail_box(self):
        print('choice_mail_box for ', self.printer, self.box_name)
        try:
            if self.printer == 5:

                list_mail_box = self.driver.find_elements(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td/table/tbody/tr/td/table")
                count = 0
                lst = list_mail_box[0].text.split("\n")
                for el in lst:
                    if self.box_name in el:
                        break
                    count = count + 1
            if self.printer == 1 or self.printer == 4  or self.printer == 5:
                mail_box_num = self.driver.find_element(By.NAME, "list{}".format(count)).click()
            if self.printer == 2:
                self.driver.find_element(By.XPATH, "/html/body/form[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/small/input").click()
                self.driver.find_element(By.XPATH, "/html/body/form[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/small/input").clear()
                QtCore.QThread.msleep(1000)
                self.driver.find_element(By.XPATH, "/html/body/form[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/small/input").send_keys(self.box_name)
                mail_box_num = self.driver.find_element(By.XPATH, "/html/body/form[1]/p[3]/small/input").click()
            if self.printer == 3:
                self.driver.find_element(By.XPATH, "/html/body/form[1]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/small/input").click()
                self.driver.find_element(By.XPATH, "/html/body/form[1]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/small/input").clear()
                QtCore.QThread.msleep(500)
                self.driver.find_element(By.XPATH, "/html/body/form[1]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/small/input").send_keys(self.box_name)
                QtCore.QThread.msleep(1000)
                mail_box_num = self.driver.find_element(By.XPATH, "/html/body/form[1]/p[2]/small/input").click() 
            if self.printer == 6:
                iframe = self.driver.find_element(By.NAME,"theTree")
                self.driver.switch_to.frame(iframe) #//*[@id="n1"]/a
                str1 = """//*[@id="u0"]"""
                mail_boxes = self.driver.find_elements(By.XPATH, str1)
                count = 0
                boxes = mail_boxes[0].text.split("\n")
                for el in boxes:
                    count = count + 1
                    if self.box_name in el:
                        break
                print("COUNT ", count, len(boxes))
                if count == len(boxes):
                    self.toError.emit("Почтовый ящик не найден")
                    return 0
                else:
                    str1 = """//*[@id="n""" + str(count) + """"]/a"""
                    mail_box_num = self.driver.find_element(By.XPATH, str1).click() 
            return 1
        except Exception as ex:
            self.toError.emit("Ошибка смены ящика")
            print(ex)
            return 0
            
    def list_files(self):
        self.data_files = []
        count = 0
        try:
            self.toError.emit("Обновление списка файлов ")
            if self.printer == 1 or self.printer == 4  or self.printer == 5:
                l = self.driver.find_elements(By.XPATH,"/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr")
                self.driver.find_elements(By.CSS_SELECTOR,"body > form:nth-child(4) > table:nth-child(4) > tbody > tr > td > table > tbody > tr:nth-child(2) > td > table > tbody > tr ")
            if self.printer == 2:
                l = self.driver.find_elements(By.XPATH,"/html/body/form[4]/table/tbody/tr[2]/td/table/tbody/tr")
            if self.printer == 3:
                l = self.driver.find_elements(By.XPATH,"/html/body/form[3]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr")
            if self.printer == 6:
                self.driver.switch_to.default_content()
                iframe1 = self.driver.find_element(By.NAME,"theContent")
                self.driver.switch_to.frame(iframe1)
                l = self.driver.find_elements(By.XPATH,"/html/body/form/div[1]/div[2]/table/tbody/tr")
                self.toProgressBar.emit(50)
                for el in l:
                    if not "Нет файлов для отображения" in el.text:
                        self.data_files.append(el.text)
                        self.toProgressBar.emit(50 + int(50*count/len(l)))
                return    
            self.toProgressBar.emit(50)
            
            for el in l:
                self.toProgressBar.emit(50 + int(50*count/len(l)))
                count = count + 1
                if len(el.text) != 0:
                    if ('Документ номер Имя документа' not in el.text) and ('Документов нет' not in el.text) and ('Номер документа' not in el.text):
                        #print(el.text)
                        self.data_files.append(el.text) 

        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка обновления списка")
            return None
            
    def update_file(self, box_name):  
        try:
            self.box_name = box_name
            self.toProgressBar.emit(0)
            self.driver.refresh()
            if self.printer == 6:
                self.toProgressBar.emit(30)
                res1 = self.choice_mail_box()
                self.toProgressBar.emit(70)
                if res1 == 1:
                    self.list_files()
                self.toProgressBar.emit(90)
                self.toInit.emit(self.str1, self.data_files, self.printer)
                self.toProgressBar.emit(100)
                return
            self.toProgressBar.emit(10)
            iframe = self.driver.find_element(By.NAME,"NF")
            self.driver.switch_to.frame(iframe)
            self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
            self.toProgressBar.emit(30)
            self.driver.switch_to.default_content()
            iframe1 = self.driver.find_element(By.NAME,"RF")
            self.driver.switch_to.frame(iframe1)
            self.toProgressBar.emit(40)
            self.choice_mail_box()
            self.list_files()
            self.toInit.emit(self.str1, self.data_files, self.printer)
            self.toProgressBar.emit(100)
        except:
            self.toError.emit("Ошибка обновления ящика")
   
    def load_file(self, n_files, del_in_mail, del_tif, convert_to_tif, vert, current_files_names):
        self.vert = vert #ориентация выходного документа
        self.del_in_mail = del_in_mail          #удалять файл после скачивания
        self.del_tif = del_tif                  #удалять tif при конвертации
        self.convert_to_tif = convert_to_tif    #конвертировать в pdf
        self.n_files = n_files #номера файлов
        self.current_files_names = current_files_names #имена файлов
        print(self.n_files, self.current_files_names, len(self.n_files), len(self.current_files_names))
        
        self.toProgressBar.emit(0)
        
        if self.printer == 3:
            self.load72()
            return
        if self.printer == 2:
            self.load133()
            return
        if self.printer == 6:
            self.load7225()
            return
        try:
            str1 = ''
            #выбор формата скачивания
            if self.printer == 5:
                select = Select(self.driver.find_element(By.XPATH, "/html/body/form[4]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td[1]/small/select"))
                if self.convert_to_tif == 1:
                    select.select_by_value('PDF')    
                else:
                    select.select_by_value("TIFJPG") 
                
            for el in self.n_files: #отметить все файлы на скачивание
                if self.printer == 1 or self.printer == 4 or self.printer == 5:
                    str1 = '/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(el + 2) + ']/td[1]/input'
                check_box1 = self.driver.find_element(By.XPATH, str1).click()
            print('2')
            
            #нажимаем ссылки на скачивание
            if self.printer == 1:
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input").click()
                file1 = self.driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td/small/a").click()                   
            if self.printer == 4:
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[5]/td[2]/input").click()
                count = 1
                for el in self.n_files:
                    if len(self.n_files) == 1:
                        file1 = self.driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td/small/a").click()
                    else:
                        file1 = self.driver.find_element(By.XPATH,"/html/body/table[" + str(count) + "]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td/small/a").click()
                        QtCore.QThread.msleep(1000)
                        count = count + 1
            if self.printer == 5:
                count = 1
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[10]/td[2]/input").click()
                for el in self.n_files:
                    if len(self.n_files) == 1:
                        file1 = self.driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td/small/a").click()
                    else:
                        file1 = self.driver.find_element(By.XPATH,"/html/body/table[" + str(count) + "]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td/small/a").click()
                        QtCore.QThread.msleep(1000)
                        count = count + 1
            self.toProgressBar.emit(10)
            QtCore.QThread.msleep(1000)
            if self.printer == 4:
                time.sleep(3)
            self.driver.get(url=self.ip + self.printers[self.printer-1]['url'])
            self.toProgressBar.emit(20)
            iframe = self.driver.find_element(By.NAME,"NF")
            self.driver.switch_to.frame(iframe)
            self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
            self.driver.switch_to.default_content()
            iframe1 = self.driver.find_element(By.NAME,"RF")
            self.driver.switch_to.frame(iframe1)
            self.choice_mail_box()
            self.toProgressBar.emit(30)
            
            if self.del_in_mail == 1: #если надо удалять файл после скачивания              
                for el in self.n_files:
                    if self.printer == 1 or self.printer == 4 or self.printer == 5:
                        str1 = '/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(el + 2) + ']/td[1]/input'
                    if self.printer == 2:
                        str1 = '/html/body/formнуж[4]/table/tbody/tr[2]/td/table/tbody/tr[' + str(el + 2) + ']/td[1]/input'
                    check_box1 = self.driver.find_element(By.XPATH, str1).click()
                
                time.sleep(2) #3
                
                if self.printer == 1  or self.printer == 4 or self.printer == 5:
                    button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/div/small/input[2]").click()
                if self.printer == 2:    
                    button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/div/input").click()
               
                QtCore.QThread.msleep(800) #2
                obj = self.driver.switch_to.alert
                mes = obj.text
                print(mes)
                QtCore.QThread.msleep(1800) #2
                obj.accept() #подтвердить удаление
                print('accept')
                QtCore.QThread.msleep(1800) #3
                if self.printer == 4:
                    time.sleep(5)
                self.driver.refresh()
                print('refresh')
                QtCore.QThread.msleep(1100) #1
                self.toProgressBar.emit(40)
                iframe = self.driver.find_element(By.NAME,"NF")
                print('NF')
                self.driver.switch_to.frame(iframe)
                print('SWITCH')
                self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
                print('mail box')
                self.driver.switch_to.default_content()
                print('def content')
                iframe1 = self.driver.find_element(By.NAME,"RF")
                print('RF')
                self.driver.switch_to.frame(iframe1)
                print('Frame')
                self.toProgressBar.emit(45)
                self.choice_mail_box()
                print('choise_mail')
                self.list_files() 
                print('list')
                self.toInit.emit(self.str1, self.data_files, self.printer)
            # file_find = 0
            #поиск скачанного файла
            # str2 = self.current_files_names[0] + '.tif'
            # os.chdir(self.doc_dir) #переключиться на папку с документами
            # for file in glob.glob(str2):
                # print('Нашли файл tiff', file)
                # file_find = 1
                # if self.convert_to_tif == 1: 
                    # convert_with_auto_rotate(str2, self.vert) #конвертация в pdf
                    # str2 = self.current_files_names[0] + '.pdf'
                    # if self.del_tif == 1:
                        # os.remove(file)
            # if  file_find == 0:
                # str2 = self.current_files_names[0] + '.pdf'
                # for file in glob.glob(str2):
                    # print('Нашли файл pdf', file)
                    # file_find = 1

            # self.toInit.emit(self.str1, self.data_files, self.printer)
            # if file_find ==1:
                # self.toError.emit("Файл " + str2 + " скачан")
            # else:
                # self.toError.emit("Ошибка скачивания файла")
            self.toError.emit("Файлы скачаны")
            self.toProgressBar.emit(100)
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка скачивания файла")              
    #--------------------------------------------- LOAD 7225 ------------------------        
    def load7225(self):
        try:   
            percent = 100 / len(self.n_files)
            progress = 0
            count = 0
            for el in self.n_files: #по очереди качаем все файлы
                self.toError.emit("Скачивание файла " + self.current_files_names[count])
                #print("Скачивание файла ", self.current_files_names[count])
                if self.del_in_mail == 1: #если надо удалять файл после скачивания
                    number = el+1-count
                else:
                    number = el+1
                str1 = "body > form > div.boundingBox.normalSpacing > div.boxBody.tableBody > table > tbody > tr:nth-child(" + str(number) + ") > td:nth-child(5) > button"
                self.driver.find_element(By.CSS_SELECTOR, str1).click()
                QtCore.QThread.msleep(500)
                if self.del_in_mail == 1: #если надо удалять файл после скачивания
                    str1 = "body > form > div.boundingBox.normalSpacing > div.boxBody.tableBody > table > tbody > tr:nth-child(" + str(number) + ") > td:nth-child(5) > select"
                    select = Select(self.driver.find_element(By.CSS_SELECTOR, str1))
                    select.select_by_value('dlt')    
                    str1 = "body > form > div.boundingBox.normalSpacing > div.boxBody.tableBody > table > tbody > tr:nth-child(" + str(number) + ") > td:nth-child(5) > button"
                    self.driver.find_element(By.CSS_SELECTOR, str1).click()
                    QtCore.QThread.msleep(500) 
                    obj = self.driver.switch_to.alert
                    mes = obj.text
                    QtCore.QThread.msleep(1000) #2
                    obj.accept() #подтвердить удаление
                progress = progress + percent
                self.toProgressBar.emit(int(progress))
                count = count + 1
            if self.del_in_mail == 1: #если надо удалять файл после скачивания
                #self.choice_mail_box()
                self.list_files() 
                self.toInit.emit(self.str1, self.data_files, self.printer)
            self.toProgressBar.emit(100)
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка скачивания файла")
    #--------------------------------------------- LOAD 133 ------------------------        
    def load133(self):
        try:   
            percent = 100 / len(self.n_files)
            progress = 0
            count = 0
            for el in self.n_files: #по очереди качаем все файлы
                self.toError.emit("Скачивание файла " + self.current_files_names[count])
                str1 = '/html/body/form[4]/table/tbody/tr[2]/td/table/tbody/tr[' + str(el + 2) + ']/td[1]/input'
                check_box1 = self.driver.find_element(By.XPATH, str1).click()  
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/center/small/input").click()#вызвать
                file1 = self.driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr[3]/td[2]/small/a").click()
                QtCore.QThread.msleep(800)
           
                self.driver.get(url=self.ip + self.printers[self.printer-1]['url'])
                iframe = self.driver.find_element(By.NAME,"NF")
                self.driver.switch_to.frame(iframe)
                self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
                self.driver.switch_to.default_content()
                iframe1 = self.driver.find_element(By.NAME,"RF")
                self.driver.switch_to.frame(iframe1)
                self.choice_mail_box()
                
                file_find = 0
                #поиск скачанного файла
                str2 = self.current_files_names[count] + '.tif'
                os.chdir(self.doc_dir) #переключиться на папку с документами
                for file in glob.glob(str2):
                    print('Нашли файл tiff', file)
                    file_find = 1
                    if self.convert_to_tif == 1: 
                        convert_with_auto_rotate(str2, self.vert) #конвертация в pdf
                        str2 = self.current_files_names[count] + '.pdf'
                        if self.del_tif == 1:
                            os.remove(file)
              
                progress = progress + percent
                self.toProgressBar.emit(int(progress))
                count = count + 1
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка скачивания файла")    
            
    #--------------------------------------------- LOAD 72 ------------------------        
    def load72(self):
        try:
            select = Select(self.driver.find_element(By.XPATH, "/html/body/form[3]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td/small/select"))
            if self.convert_to_tif == 1:
                select.select_by_value('PDF')    
            else:
                select.select_by_value("TIFJPG")             
            percent = 100 / len(self.n_files)
            progress = 0
            count = 0
            for el in self.n_files: #по очереди качаем все файлы
                self.toError.emit("Скачивание файла " + self.current_files_names[count])
                count = count + 1
                str1 = '/html/body/form[3]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(el + 2) + ']/td[1]/input'
                check_box1 = self.driver.find_element(By.XPATH, str1).click()
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[3]/center/small/input").click() #вызвать
                file1 = self.driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td[2]/small/a").click()
                QtCore.QThread.msleep(800)        
                self.driver.get(url=self.ip + self.printers[self.printer-1]['url'])   
                iframe = self.driver.find_element(By.NAME,"NF")
                self.driver.switch_to.frame(iframe)
                self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
                self.driver.switch_to.default_content()
                iframe1 = self.driver.find_element(By.NAME,"RF")
                self.driver.switch_to.frame(iframe1)
                self.choice_mail_box()
                progress = progress + percent
                self.toProgressBar.emit(int(progress))
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка скачивания файла")
            
    #---------------------------------------- DELETE FILES---------------------------------------------
    def delete_file(self, n_files, del_in_mail, del_tif, convert_to_tif, vert, current_files_names):
        self.n_files = n_files
        self.current_files_names = current_files_names
        self.toProgressBar.emit(0)
        if self.printer == 3:
            self.delete72()
            return
        if self.printer == 2:
            self.delete133()
            return
        if self.printer == 6:
            self.delete7225()
            return
        try:
            str1 = ''
            print('1')
            for el in self.n_files:
                if self.printer == 1 or self.printer == 4 or self.printer == 5:
                    str1 = '/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(el + 2) + ']/td[1]/input'
                check_box1 = self.driver.find_element(By.XPATH, str1).click()
            time.sleep(2) #3
            
            if self.printer == 1 or self.printer == 4 or self.printer == 5:
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/div/small/input[2]").click()
            if self.printer == 2:    
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/div/input").click()
            if self.printer == 3:    
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[3]/div/input").click()
               
            QtCore.QThread.msleep(800) #2
            obj = self.driver.switch_to.alert
            mes = obj.text
            print(mes)
            QtCore.QThread.msleep(1700) #2
            obj.accept() #подтвердить удаление
            print('accept')
            time.sleep(2) #3
            if self.printer == 4:
                time.sleep(5)
            
            self.driver.get(url=self.ip + self.printers[self.printer-1]['url'])
            print('refresh')
            QtCore.QThread.msleep(1100) #1
            self.toProgressBar.emit(40)
            iframe = self.driver.find_element(By.NAME,"NF")
            print('NF')
            self.driver.switch_to.frame(iframe)
            print('SWITCH')
            self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
            print('mail box')
            self.driver.switch_to.default_content()
            print('def content')
            iframe1 = self.driver.find_element(By.NAME,"RF")
            print('RF')
            self.driver.switch_to.frame(iframe1)
            print('Frame')
            self.toProgressBar.emit(45)
            self.choice_mail_box()
            print('choise_mail')
            self.list_files() 
            print('list')
            self.toInit.emit(self.str1, self.data_files, self.printer)
            self.toError.emit("Файл удален")
            self.toProgressBar.emit(100)
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка удаления файла")
           
    #---------------------------------------DELETE 7225-------------------------------------        
    def delete7225(self):    
        try:
            percent = 100 / len(self.n_files)
            progress = 0
            count = 0
            for el in self.n_files: #по очереди удаляем все файлы
                self.toError.emit("Удаление файла " + self.current_files_names[count])
                number = el+1-count
                
                str1 = "body > form > div.boundingBox.normalSpacing > div.boxBody.tableBody > table > tbody > tr:nth-child(" + str(number) + ") > td:nth-child(5) > select"
                select = Select(self.driver.find_element(By.CSS_SELECTOR, str1))
                select.select_by_value('dlt')    
                str1 = "body > form > div.boundingBox.normalSpacing > div.boxBody.tableBody > table > tbody > tr:nth-child(" + str(number) + ") > td:nth-child(5) > button"
                self.driver.find_element(By.CSS_SELECTOR, str1).click()
                QtCore.QThread.msleep(500) 
                obj = self.driver.switch_to.alert
                mes = obj.text
                QtCore.QThread.msleep(1000) #2
                obj.accept() #подтвердить удаление
                
                progress = progress + percent
                self.toProgressBar.emit(int(progress))
                count = count + 1
            if self.del_in_mail == 1: #если надо удалять файл после скачивания
                #self.choice_mail_box()
                self.list_files() 
                self.toInit.emit(self.str1, self.data_files, self.printer)
            self.toProgressBar.emit(100)
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка скачивания файла")    
    #---------------------------------------DELETE 133-------------------------------------        
    def delete133(self):    
        try:
            percent = 100 / len(self.n_files)
            progress = 0
            count = 0
            for el in self.n_files: #по очереди удаляем все файлы
                self.toError.emit("Удаление файла " + self.current_files_names[count])
                str1 = '/html/body/form[4]/table/tbody/tr[2]/td/table/tbody/tr[' + str(el + 2 - count) + ']/td[1]/input'
                count = count + 1
                check_box1 = self.driver.find_element(By.XPATH, str1).click()
                time.sleep(2) #3
    
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/div/input").click()
                
               
                QtCore.QThread.msleep(800) #2
                obj = self.driver.switch_to.alert
                mes = obj.text
                print(mes)
                QtCore.QThread.msleep(1700) #2
                obj.accept() #подтвердить удаление
                print('accept')
                time.sleep(2) #3

                self.driver.refresh()
                print('refresh')
                QtCore.QThread.msleep(1100) #1
                iframe = self.driver.find_element(By.NAME,"NF")
                print('NF')
                self.driver.switch_to.frame(iframe)
                print('SWITCH')
                self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
                print('mail box')
                self.driver.switch_to.default_content()
                print('def content')
                iframe1 = self.driver.find_element(By.NAME,"RF")
                print('RF')
                self.driver.switch_to.frame(iframe1)
                print('Frame')
                self.choice_mail_box()
                print('choise_mail')
                #self.list_files() 
                #print('list')
                #self.toInit.emit(self.str1, self.data_files, self.printer)
                #self.toError.emit("Файл " + self.current_file_name + " удален")
                
                progress = progress + percent
                self.toProgressBar.emit(int(progress))
                
            self.list_files() #обновим список файлов
            self.toInit.emit(self.str1, self.data_files, self.printer)
            self.toProgressBar.emit(100)
            
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка удаления файла")        
    #---------------------------------------DELETE 72-------------------------------------  
    def delete72(self):    
        try:
            percent = 100 / len(self.n_files)
            progress = 0
            count = 0
            for el in self.n_files: #по очереди удаляем все файлы
                self.toError.emit("Удаление файла " + self.current_files_names[count])
                str1 = '/html/body/form[3]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(el + 2 - count) + ']/td[1]/input'
                count = count + 1 #сдвинем указатель на гулку удаляемого файла
                check_box1 = self.driver.find_element(By.XPATH, str1).click()
                time.sleep(2) #3
                button1 = self.driver.find_element(By.XPATH, "/html/body/form[3]/div/input").click()
                QtCore.QThread.msleep(800) #2
                obj = self.driver.switch_to.alert
                mes = obj.text
                print(mes)
                QtCore.QThread.msleep(1700) #2
                obj.accept() #подтвердить удаление
                time.sleep(2) #3               
                self.driver.get(url=self.ip + self.printers[self.printer-1]['url'])
                QtCore.QThread.msleep(1100) #1
                iframe = self.driver.find_element(By.NAME,"NF")
                self.driver.switch_to.frame(iframe)
                self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
                self.driver.switch_to.default_content()
                iframe1 = self.driver.find_element(By.NAME,"RF")
                self.driver.switch_to.frame(iframe1)
                self.choice_mail_box()
                #self.toInit.emit(self.str1, self.data_files, self.printer)
                progress = progress + percent
                self.toProgressBar.emit(int(progress))
            self.list_files() #обновим список файлов
            #self.toInit.emit(self.str1, self.data_files, self.printer)
            
            self.toProgressBar.emit(100)
        except Exception as ex:
            print(ex)
            self.toError.emit("Ошибка удаления файла")
            
            
#--------------------------------M Y W I N D O W----------------------
class mywindow(QtWidgets.QMainWindow):
    printer = 0
    driver = None
    data_files = []
    n_file = 0 #номер файла на мечать
    current_files_name = []
    vert = 1 #ориентация выходного документа
    del_in_mail = 1 #удалять файл после скачивания
    del_tif = 1 #удалять tif при конвертации
    convert_to_tif = 1 #конвертировать в pdf
    doc_dir = "D:\Pavel\Download" #куда сохранять файлы
    ip = 'http://192.168.0.128/'
    trayIcon = None
    my_dir = None
    need_pdf = 1
    n_box_mail = 1
    box_name = "None"

    emit_start = QtCore.pyqtSignal()
    emit_connect = QtCore.pyqtSignal()
    emit_update = QtCore.pyqtSignal(str)
    emit_mail_update = QtCore.pyqtSignal()
    emit_load_file = QtCore.pyqtSignal(object, int, int, int, int, object)
    emit_delete_file = QtCore.pyqtSignal(object, int, int, int, int, object)
    
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        
        self.my_dir = os.getcwd() #текущая директория
        
        self.trayIcon = QSystemTrayIcon(QIcon('free-icon-printer.png'))
        self.trayIcon.setToolTip('Обзор сканера')
        self.trayIcon.activated.connect(self.systemIcon)
        
        show_action = QAction("Показать", self)
        quit_action = QAction("Выход", self)
        hide_action = QAction("Скрыть", self)
        show_action.triggered.connect(self.show)
        hide_action.triggered.connect(self.hide)
        quit_action.triggered.connect(qApp.quit)
        tray_menu = QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(quit_action)
        self.trayIcon.setContextMenu(tray_menu)
        self.trayIcon.show()

        self.readConfig() #читаем конфиг и устанавливаем галочки

        # подключение клик-сигнал к слоту btnClicked
        self.ui.pushButton.clicked.connect(self.btnClicked)
        self.ui.pushButton_2.clicked.connect(self.button_update)
        self.ui.pushButton_3.clicked.connect(self.load_file)
        self.ui.pushButton_4.clicked.connect(self.mail_update)
        self.ui.pushButton_5.clicked.connect(self.getDirectory)
        self.ui.pushButton_6.clicked.connect(self.openDir)
        self.ui.pushButton_7.clicked.connect(self.delete_file)
        
        self.ui.tableWidget.cellClicked.connect(self.tab_click)
        self.ui.radioButton.toggled.connect(self.onClickedHorz)
        self.ui.radioButton_2.toggled.connect(self.onClickedVert)
        self.ui.checkBox.stateChanged.connect(self.changeTitle)
        self.ui.checkBox_2.stateChanged.connect(self.changeTitle_2)
        self.ui.checkBox_3.stateChanged.connect(self.changeTitle_3)
        
        self.ui.tableWidget.setColumnCount(4)
        self.ui.label_2.setText("Поиск принтера...")
        self.ui.label.setText("Поиск принтера...")
        self.ui.label_3.setVisible(False)
        self.ui.lineEdit.setText(str(self.box_name))          
         
        if self.del_in_mail == 1:
            self.ui.checkBox.setChecked(True)
        if self.del_tif == 1:
            self.ui.checkBox_2.setChecked(True)
        if self.convert_to_tif == 1:
            self.ui.checkBox_3.setChecked(True)
            
        self.ui.label_7.setText(self.doc_dir)
        
        self.thread = QtCore.QThread()
        self.browserHandler = BrowserHandler(self.doc_dir, self.ip, self.n_box_mail, self.box_name)
        self.browserHandler.moveToThread(self.thread)
        self.browserHandler.toProgressBar.connect(self.inProgressBar)
        self.browserHandler.toInit.connect(self.initComplete)
        self.browserHandler.toError.connect(self.funcError)
        self.thread.started.connect(self.browserHandler.run)
        self.thread.start()
        self.dis_gui()
        #-----------------
        self.emit_connect.connect(self.browserHandler.run)
        self.emit_update.connect(self.browserHandler.update_file)
        self.emit_load_file.connect(self.browserHandler.load_file)
        self.emit_delete_file.connect(self.browserHandler.delete_file)
         
    #-----------------------------------------------------------------------------
    @QtCore.pyqtSlot(int)
    def inProgressBar(self, dig):
        self.ui.progressBar.setValue(dig)
        # if dig == 100:
            # self.ui.label_2.setText("")
        
    @QtCore.pyqtSlot(str, object, int)
    def initComplete(self, lab, listFiles, printer):
       # print('receive init printer=', printer)   
        self.printer = printer
        self.ui.label.setText(lab)
        self.data_files = listFiles
        self.update_file_list()
        
        if  self.printer == 3:
            self.ui.checkBox.setVisible(False)
            self.ui.checkBox.setChecked(False)
            self.ui.checkBox_2.setVisible(False)
            self.ui.checkBox_2.setChecked(False)
            self.ui.radioButton.setVisible(False)
            self.ui.radioButton_2.setVisible(False)
            self.ui.label_4.setVisible(False)
        if  self.printer == 2:
            self.ui.checkBox.setVisible(False)
            self.ui.checkBox.setChecked(False)
        if  self.printer == 5:
            self.ui.checkBox_2.setVisible(False)
            self.ui.checkBox_2.setChecked(False)
            self.ui.radioButton.setVisible(False)
            self.ui.radioButton_2.setVisible(False)
            self.ui.label_4.setVisible(False)
        if  self.printer == 6:
            self.ui.checkBox_2.setVisible(False)
            self.ui.checkBox_2.setChecked(False)
            self.ui.checkBox_3.setVisible(False)
            self.ui.checkBox_3.setChecked(False)
            self.ui.radioButton.setVisible(False)
            self.ui.radioButton_2.setVisible(False)
            self.ui.label_4.setVisible(False)
        self.en_gui()
        
    @QtCore.pyqtSlot(str)
    def funcError(self, str1):
        print('receive error ', str1)
        self.ui.label_2.setText(str1)
        self.en_gui()
    
    def dis_gui(self):
        self.ui.pushButton.setEnabled(False)           
        self.ui.pushButton_2.setEnabled(False)    
        self.ui.pushButton_3.setEnabled(False) 
        self.ui.pushButton_4.setEnabled(False) 
        self.ui.pushButton_5.setEnabled(False)
        self.ui.pushButton_7.setEnabled(False)        

    def en_gui(self):
        self.ui.pushButton.setEnabled(True)
        #if self.need_pdf == 1:
        #self.ui.checkBox_2.setEnabled(True)    
        #self.ui.checkBox_3.setEnabled(True)
        #else:
        #    self.ui.checkBox_2.setChecked(False)
        #    self.ui.checkBox_3.setChecked(False)
        #    self.ui.checkBox_2.setEnabled(False)
        #    self.ui.checkBox_3.setEnabled(False)
        self.ui.pushButton_2.setEnabled(True)    
        self.ui.pushButton_3.setEnabled(True)
        self.ui.pushButton_4.setEnabled(True)
        self.ui.pushButton_5.setEnabled(True)
        self.ui.pushButton_7.setEnabled(True)
        
    
    def systemIcon(self, reason):
        if reason == 3:
            self.show()

    def getDirectory(self):
        dirlist = QFileDialog.getExistingDirectory(self,"Выбрать папку",".")
        #self.plainTextEdit.appendHtml("<br>Выбрали папку: <b>{}</b>".format(dirlist))
        str1 = '/'
        str2 = '\\'
        dirlist = dirlist.replace(str1, str2)
        print(dirlist)
        self.doc_dir = dirlist
        self.ui.label_7.setText(self.doc_dir)
        self.crudConfig()
        self.dis_gui()
        self.ui.label.setText("Перезапустите программу")
     
    def openDir(self):
        os.system('explorer')
        os.startfile(self.doc_dir)
        
    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.trayIcon.showMessage("Обзор сканера", "Я пока буду в трее", QSystemTrayIcon.Information,2000)


    def mail_update(self):
        self.box_name = self.ui.lineEdit.text()
        self.crudConfig()
        self.ui.tableWidget.clear()
        self.ui.label_2.setText("Смена почтового ящика...")
        self.dis_gui()
        self.emit_update.emit(self.box_name)
        # try:
            # temp = int(shost)
            # if temp > 0 and temp < 1000:
                # self.n_box_mail = temp
                # self.crudConfig()
                # self.ui.tableWidget.clear()
                # self.ui.label_2.setText("Обновление номера почтового ящика...")
                # self.emit_update.emit(self.n_box_mail)
                # self.dis_gui()
            # else:
                # self.ui.lineEdit.clear()
        # except:
            # print("Введено не число")
            # self.ui.lineEdit.clear()
                       
    def crudConfig(self):
        path = self.my_dir + "/settings.ini"
        if not os.path.exists(path):
            self.createConfig()
        config = configparser.ConfigParser()
        config.read(path)
    
        config.set("Settings", "del_in_mail", str(self.del_in_mail))
        config.set("Settings", "del_tif", str(self.del_tif))
        config.set("Settings", "convert_to_tif", str(self.convert_to_tif))
        config.set("Settings", "ip", str(self.ip)) 
        config.set("Settings", "doc_dir", self.doc_dir)
        config.set("Settings", "n_box_mail", str(self.n_box_mail))
        config.set("Settings", "box_name", self.box_name)

        with open(path, "w") as config_file:
            config.write(config_file)
        
    def createConfig(self):
        path = self.my_dir + "/settings.ini"
        config = configparser.ConfigParser()
        config.add_section("Settings")
        config.set("Settings", "del_in_mail", str(self.del_in_mail))
        config.set("Settings", "del_tif", str(self.del_tif))
        config.set("Settings", "convert_to_tif", str(self.convert_to_tif))
        config.set("Settings", "ip", str(self.ip))
        config.set("Settings", "doc_dir", self.doc_dir)
        config.set("Settings", "n_box_mail", str(self.n_box_mail))
        config.set("Settings", "box_name", self.box_name)
        
        with open(path, "w") as config_file:
            config.write(config_file)
            
    def readConfig(self):
        path = self.my_dir + "/settings.ini"
        if not os.path.exists(path):
            self.createConfig()
        config = configparser.ConfigParser()
        config.read(path)
        
        self.del_in_mail = int(config.get("Settings", "del_in_mail"))
        self.del_tif = int(config.get("Settings", "del_tif"))
        self.convert_to_tif = int(config.get("Settings", "convert_to_tif"))
        self.ip = config.get("Settings", "ip")
        self.doc_dir = config.get("Settings", "doc_dir")
        self.n_box_mail = int(config.get("Settings", "n_box_mail"))
        self.box_name = config.get("Settings", "box_name")
        print("read ", self.del_in_mail, self.del_tif, self.convert_to_tif, self.ip)
        
    def changeTitle(self, state):
        if state == Qt.Checked:
            self.del_in_mail = 1
        else:
            self.del_in_mail = 0
        self.crudConfig()
    
    def changeTitle_2(self, state):
        if state == Qt.Checked:
            self.del_tif = 1
        else:
            self.del_tif = 0
        self.crudConfig()
                
    def changeTitle_3(self, state):
        if state == Qt.Checked:
            self.convert_to_tif = 1
        else:
            self.convert_to_tif = 0
        self.crudConfig() 
    
    def onClickedHorz(self):
        if self.ui.radioButton.isChecked():
            self.vert = 0
        
    def onClickedVert(self):
        if self.ui.radioButton_2.isChecked():
            self.vert = 1
    
    def btnClicked(self): #подключиться к принтеру
        self.ui.tableWidget.clear()
        self.ui.label.setText('Поиск принтера...')
        self.ui.label_2.setText("Поиск принтера...")
        self.emit_connect.emit()
        self.dis_gui()
        
    def button_update(self):
        self.ui.tableWidget.clear()
        self.ui.label_2.setText("Обновление списка файлов...")
        self.emit_update.emit(self.box_name)
        self.dis_gui()
        
    
    def update_file_list(self):    
        print('update_file_list')
        self.ui.tableWidget.clear()
        if  self.printer == 6:
            self.ui.tableWidget.setHorizontalHeaderLabels(['имя','дата','время', 'размер'])
            if len(self.data_files) > 0:
                self.ui.tableWidget.setRowCount(len(self.data_files))
                row = 0
                for tup in self.data_files:
                    col = 0
                    file_info = tup.split()
                    for item in file_info:
                        cellinfo = QTableWidgetItem(item)
                        self.ui.tableWidget.setItem(row, col, cellinfo)
                        col += 1
                    row += 1
                self.ui.label_3.setVisible(False)
            else:
                self.ui.label_3.setVisible(True)
            return
            
        self.ui.tableWidget.setHorizontalHeaderLabels(['номер','имя','дата','время'])
   
        if len(self.data_files) > 0:
            self.ui.tableWidget.setRowCount(len(self.data_files))
            row = 0
            for tup in self.data_files:
                col = 0
                file_info = tup.split()
                for item in file_info:
                    cellinfo = QTableWidgetItem(item)
                    self.ui.tableWidget.setItem(row, col, cellinfo)
                    col += 1
                row += 1
            self.ui.label_3.setVisible(False)
        else:
            self.ui.label_3.setVisible(True)    

    def tab_click(self):
        pass
        #items = self.ui.tableWidget.selectedItems()
        #self.n_file = self.ui.tableWidget.currentRow() + 1
        #tup = self.data_files[self.n_file - 1]
        #t_split = tup.split()
        #self.current_file_name = 'img0' + t_split[0]

    def load_file(self): #скачать файл
        items = self.ui.tableWidget.selectedItems()
        n_files = []
        current_files_names = []
        
        for item in items:
            n_files.append(item.row()) #
            tup = self.data_files[item.row()]
            t_split = tup.split()
            current_files_names.append('img0' + t_split[0])
            
        n_files = list(set(n_files)) #список выделенных строк         
        #print(n_files,current_files_names)
        if (len(n_files)):
            #self.dis_gui()
            self.ui.label_2.setText("Скачивание файла...")
            self.emit_load_file.emit(n_files, self.del_in_mail, self.del_tif, self.convert_to_tif, self.vert, current_files_names)
        else:
            self.ui.label_2.setText("Файл не выбран.")
            self.ui.label_2.setVisible(True) 
            
        #if (len(self.data_files)) and self.n_file > 0: # файлы есть и конкретный выбран
        #    self.dis_gui()
        #    self.ui.label_2.setText("Скачивание файла...")
        #    self.emit_load_file.emit(self.n_file, self.del_in_mail, self.del_tif, self.convert_to_tif, self.vert, self.current_file_name)
        #else:
        #    self.ui.label_2.setText("Файл не выбран.")
        #    self.ui.label_2.setVisible(True)

    def delete_file(self): #удалить файл
        items = self.ui.tableWidget.selectedItems()
        n_files = []
        current_files_names = []
        
        for item in items:
            n_files.append(item.row()) #
            tup = self.data_files[item.row()]
            t_split = tup.split()
            current_files_names.append('img0' + t_split[0])
            
        n_files = list(set(n_files)) #список выделенных строк         
        #print(n_files,current_files_names)
        if (len(n_files)):
            #self.dis_gui()
            self.ui.label_2.setText("Удаление файла...")
            self.emit_delete_file.emit(n_files, self.del_in_mail, self.del_tif, self.convert_to_tif, self.vert, current_files_names)
        else:
            self.ui.label_2.setText("Файл не выбран.")
            self.ui.label_2.setVisible(True) 
        # if (len(self.data_files)) and self.n_file > 0: # файлы есть и конкретный выбран
            # self.dis_gui()
            # self.ui.label_2.setText("Удаление файла...")
            # self.emit_delete_file.emit(self.n_file, self.del_in_mail, self.del_tif, self.convert_to_tif, self.vert, self.current_file_name)
        # else:
            # self.ui.label_2.setText("Файл не выбран.")
            # self.ui.label_2.setVisible(True)


os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
app = QtWidgets.QApplication([])
# app.setStyle('Windows')
application = mywindow()
application.show()

 
sys.exit(app.exec())