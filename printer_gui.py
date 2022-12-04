from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QTableWidgetItem
from wind import Ui_MainWindow  # импорт нашего сгенерированного файла
#from screeninfo import get_monitors
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
import time
import os
import glob
import configparser
#----------------------------------------------------------------- 
from tif_to_pdf2 import convert_with_auto_rotate


class BrowserHandler(QtCore.QObject):
    running = False
    newText = QtCore.pyqtSignal(str, object)
 
class mywindow(QtWidgets.QMainWindow):
    
    printer = 0
    driver = None
    data_files = []
    n_file = 0 #номер файла на мечать
    current_file_name = ''
    vert = 1 #ориентация выходного документа
    del_in_mail = 1
    del_tif = 1
    convert_to_tif = 1
    ip = 'http://192.168.0.128/'
    
    printers = [
        {'url' : 'scan.htm',  'name' : 'Xerox WorkCentre 73xx'},
        {'url' : 'prop.htm', 'name' : 'Xerox WorkCentre 1xx'}
    ]

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
        #--headless--headless--headless
        options.add_argument("--headless")

        self.driver = webdriver.Chrome(executable_path="D:\\Pavel\\chromedriver\\chromedriver.exe",chrome_options=options)
        self.driver.set_page_load_timeout(5)
        # подключение клик-сигнал к слоту btnClicked
        self.ui.pushButton.clicked.connect(self.btnClicked)
        self.ui.pushButton_2.clicked.connect(self.button_update)
        self.ui.pushButton_3.clicked.connect(self.load_file)
        self.ui.tableWidget.cellClicked.connect(self.tab_click)
        self.ui.radioButton.toggled.connect(self.onClickedHorz)
        self.ui.radioButton_2.toggled.connect(self.onClickedVert)
        self.ui.checkBox.stateChanged.connect(self.changeTitle)
        self.ui.checkBox_2.stateChanged.connect(self.changeTitle_2)
        self.ui.checkBox_3.stateChanged.connect(self.changeTitle_3)
        
        self.ui.tableWidget.setRowCount(7)
        self.ui.tableWidget.setColumnCount(4)
        self.ui.label_2.setVisible(False)
        self.ui.label_3.setVisible(False)

        count1 = 0
        for pr in self.printers:
            print('Поиск принтера ', pr['name'])
            res = self.connect_to_printer(pr)
            count1 = count1 + 1
            print('count ', count1)
            
            if res == 1:
                self.ui.label.setText(pr['name'])
                self.printer = count1
                self.choice_mail_box()
                self.update_file_list()
                break
            else:
                print('Не обнаружен принтер ', pr['name'])
                self.ui.label.setText("Принтер не обнаружен ")
         
        self.readConfig() #читаем конфиг и устанавливаем галочки
        if self.del_in_mail == 1:
            self.ui.checkBox.setChecked(True)
        if self.del_tif == 1:
            self.ui.checkBox_2.setChecked(True)
        if self.convert_to_tif == 1:
            self.ui.checkBox_3.setChecked(True)
            
    #-----------------------------------------------------------------------------
    def crudConfig(self):
        path = "settings.ini"
        if not os.path.exists(path):
            self.createConfig(path)
        config = configparser.ConfigParser()
        config.read(path)
    
        config.set("Settings", "del_in_mail", str(self.del_in_mail))
        config.set("Settings", "del_tif", str(self.del_tif))
        config.set("Settings", "convert_to_tif", str(self.convert_to_tif))
        config.set("Settings", "ip", str(self.ip)) 

        with open(path, "w") as config_file:
            config.write(config_file)
        
    def createConfig(self):
        path = "settings.ini"
        config = configparser.ConfigParser()
        config.add_section("Settings")
        config.set("Settings", "del_in_mail", str(self.del_in_mail))
        config.set("Settings", "del_tif", str(self.del_tif))
        config.set("Settings", "convert_to_tif", str(self.convert_to_tif))
        config.set("Settings", "ip", str(self.ip))
        
        with open(path, "w") as config_file:
            config.write(config_file)
            
    def readConfig(self):
        path = "settings.ini"
        if not os.path.exists(path):
            createConfig(path)
        config = configparser.ConfigParser()
        config.read(path)
        
        self.del_in_mail = int(config.get("Settings", "del_in_mail"))
        self.del_tif = int(config.get("Settings", "del_tif"))
        self.convert_to_tif = int(config.get("Settings", "convert_to_tif"))
        self.ip = config.get("Settings", "ip")
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
            print('Horiz')
            self.vert = 0
        
    def onClickedVert(self):
        if self.ui.radioButton_2.isChecked():
            print('Vert')   
            self.vert = 1
    
    def btnClicked(self): #подключиться к принтреру
        count1 = 0
        for pr in self.printers:
            print('Поиск принтера ', pr['name'])
            res = self.connect_to_printer(pr)
            count1 = count1 + 1
            print('count ', count1)
            
            if res == 1:
                self.ui.label.setText(pr['name'])
                self.printer = count1
                self.choice_mail_box()
                self.update_file_list()
                break
            else:
                print('Не обнаружен принтер ', pr['name'])
                self.ui.label.setText("Принтер не обнаружен ")

    def button_update(self):
        self.driver.refresh()
                
        iframe = self.driver.find_element(By.NAME,"NF")
        self.driver.switch_to.frame(iframe)
        self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
        self.driver.switch_to.default_content()
        iframe1 = self.driver.find_element(By.NAME,"RF")
        self.driver.switch_to.frame(iframe1)
        #self.connect_to_printer(self.printer)
        self.choice_mail_box()
        self.update_file_list()
        
    
    def update_file_list(self):    
        print('update_file_list')
        self.ui.tableWidget.clear()
        self.list_files()    
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
        self.n_file = self.ui.tableWidget.currentRow() + 1
        #print(self.data_files)
        #print(self.n_file)
        tup = self.data_files[self.n_file - 1]
        t_split = tup.split()
        self.current_file_name = 'img0' + t_split[0]
        #print(self.current_file_name)

    def load_file(self): #скачать файл
        if (len(self.data_files)) and self.n_file > 0: # файлы есть и конкретный выбран
            try:
                #print(self.printer)
                if self.printer == 1:
                    str1 = '/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(self.n_file + 1) + ']/td[1]/input'
                if self.printer == 2:
                    str1 = '/html/body/form[4]/table/tbody/tr[2]/td/table/tbody/tr[' + str(self.n_file + 1) + ']/td[1]/input'
                check_box1 = self.driver.find_element(By.XPATH, str1).click()
                if self.printer == 1:
                    button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input").click()
                    file1 = self.driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td/small/a").click()
                if self.printer == 2:    
                    button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/center/small/input").click()
                    file1 = self.driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr[3]/td[2]/small/a").click()                   
                
                #self.driver.back() #вернуться на страницу с файлами
                #time.sleep(2)
                #iframe1 = self.driver.find_element(By.NAME,"RF")
                #self.driver.switch_to.frame(iframe1)
                
                time.sleep(2)
                self.driver.refresh()
                
                iframe = self.driver.find_element(By.NAME,"NF")
                self.driver.switch_to.frame(iframe)
                self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
                self.driver.switch_to.default_content()
                iframe1 = self.driver.find_element(By.NAME,"RF")
                self.driver.switch_to.frame(iframe1)
                #self.connect_to_printer(self.printer)
                self.choice_mail_box()
                self.update_file_list()
                
                #
                if self.del_in_mail == 1: #если надо удалять файл после скачивания              
                    if self.printer == 1:
                        str1 = '/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(self.n_file + 1) + ']/td[1]/input'
                    if self.printer == 2:
                        str1 = '/html/body/form[4]/table/tbody/tr[2]/td/table/tbody/tr[' + str(self.n_file + 1) + ']/td[1]/input'
                    check_box1 = self.driver.find_element(By.XPATH, str1).click()
                    time.sleep(3)
                    button1 = self.driver.find_element(By.XPATH, "/html/body/form[4]/div/input").click()
                    
                    time.sleep(2)
                    #alert = Alert(self.driver)
                    #time.sleep(2)
                    #alert.accept()
                    obj = self.driver.switch_to.alert
                    mes = obj.text
                    print(mes)
                    time.sleep(2)
                    obj.accept()
                    
                    time.sleep(1)
                    self.driver.refresh()
                    
                    iframe = self.driver.find_element(By.NAME,"NF")
                    self.driver.switch_to.frame(iframe)
                    self.driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
                    self.driver.switch_to.default_content()
                    iframe1 = self.driver.find_element(By.NAME,"RF")
                    self.driver.switch_to.frame(iframe1)
                    #self.connect_to_printer(self.printer)
                    self.choice_mail_box()
                    self.update_file_list()
                                
                file_find = 0
                #поиск скачанного файла
                str2 = self.current_file_name + '.tif'
                for file in glob.glob(str2):
                    print('Нашли файл tiff', file)
                    file_find = 1
                    if self.convert_to_tif == 1: 
                        convert_with_auto_rotate(str2, self.vert) #конвертация в pdf
                        str2 = self.current_file_name + '.pdf'
                        if self.del_tif == 1:
                            os.remove(file)
                if  file_find == 0:
                    str2 = self.current_file_name + '.pdf'
                    for file in glob.glob(str2):
                        print('Нашли файл pdf', file)
                        file_find = 1
                        
                #check_box1 = self.driver.find_element(By.XPATH, str1).click() #снять выделение
                #self.connect_to_printer(self.printer)
                #self.driver.get(response.url)
                #time.spleep(1)
                
                
                if file_find ==1:
                    self.ui.label_2.setText("Файл " + str2 + " скачан")
                else:
                    self.ui.label_2.setText("Файл не скачан1")
                    
                    
                self.ui.label_2.setVisible(True)
                #return 1
            except Exception as ex:
                print(ex)
                self.ui.label_2.setText("Файл не скачан2")
                self.ui.label_2.setVisible(True)
                #return 0
        else:
            self.ui.label_2.setText("Файл не выбран.")
            self.ui.label_2.setVisible(True)

    
    def connect_to_printer(self, printer):
        try:
            print('START CONNECT')
            self.driver.get(url=self.ip + printer['url'])
            time.sleep(1) #2
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

    def choice_mail_box(self):
        print('choice_mail_box for ', self.printer)
        try:
            if self.printer == 1:
                mail_box_num = self.driver.find_element(By.NAME, "list{}".format(num)).click()
            if self.printer == 2:
                mail_box_num = self.driver.find_element(By.XPATH, "/html/body/form[1]/p[3]/small/input").click()
            return 1
        except Exception as ex:
            print(ex)
            return 0
    
    
    def list_files(self):
        self.data_files = []
        try:
            if self.printer == 1:
                l = self.driver.find_elements(By.XPATH,"/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr")
            if self.printer == 2:
                l = self.driver.find_elements(By.XPATH,"/html/body/form[4]/table/tbody/tr[2]/td/table/tbody/tr")
            
            for el in l:
                #print(el.text)
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