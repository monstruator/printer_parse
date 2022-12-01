from selenium import webdriver
from selenium.webdriver.common.by import By
import time

# url = "https://www.instagram.com/"


#http://192.168.0.73/box/4ndmTE52NNNLRNX3IvT9nnlhnn4seNt68I4xQ3fieI8v27GvNnYv9OQGOiC7M81Vx969ch27PM5j97XB/img01644.pdf

def connect_to_printer(url, driver):
    try:
        print('START CONNECT')
        driver.get(url=url)
        #phtml = driver.page_source
        time.sleep(1) #2
        #print(phtml)
        print('search')
        iframe = driver.find_element(By.NAME,"NF")
        driver.switch_to.frame(iframe)
        mail_box = driver.find_element(By.LINK_TEXT,'Почтовый ящик').click()
        driver.switch_to.default_content()
        iframe1 = driver.find_element(By.NAME,"RF")
        driver.switch_to.frame(iframe1)
        return 1
    except Exception as ex:
        print("ERROR CONNECT")
        #print(ex)
        return 0

    
def choice_mail_box(driver, num):
    try:
        mail_box_num = driver.find_element(By.NAME, "list{}".format(num)).click()
        return 1
    except Exception as ex:
        print(ex)
        return 0
    
def download_file(driver, num):
    try:
        str1 = '/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[' + str(num + 1) + ']/td[1]/input'
        check_box1 = driver.find_element(By.XPATH, str1).click()
        button1 = driver.find_element(By.XPATH, "/html/body/form[4]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input").click()
        file1 = driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td/small/a").click()
        driver.back()
        time.sleep(2)
        iframe1 = driver.find_element(By.NAME,"RF")
        driver.switch_to.frame(iframe1)
        check_box1 = driver.find_element(By.XPATH, str1).click()
        return 1
    except Exception as ex:
        print(ex)
        return 0


def list_mail_box(driver):
    try:
        l = driver.find_elements(By.XPATH,"/html/body/form[1]/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr")
        #/html/body/form[1]/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody
        count = 0
        mail_box_list = []
        mail_box_id = 0
        for el in l:
            if len(el.text) != 0:
                if ('(Не использ.)' not in el.text) and ('Номер почтового ящика' not in el.text):
                    #mail_box_id = int(el.text[0:3])
                    mail_box_list.append(el.text)
                    #print(el.text, mail_box_id)
            count = count + 1  
            if count > 10:
                break
                #mail_box_id = mail_box_id + 1
        #print('Всего найдено ящиков:', count)
        return mail_box_list
    except:
        return None
  

def list_files(driver):
    try:
        l = driver.find_elements(By.XPATH,"/html/body/form[4]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr")
        count = 0
        files_list = []
        file_id = 0
        for el in l:
            if len(el.text) != 0:
                if 'Документ номер Имя документа' not in el.text:
                    #print(el.text)
                    files_list.append(el.text)     
        return files_list
    except Exception as ex:
        print(ex)
        return None
  
def main():

    url = "http://192.168.0.73/scan.htm"

    options = webdriver.ChromeOptions() 
    prefs = {"download.default_directory" : "D:\Pavel"}
    options.add_experimental_option("prefs",prefs)
    options.add_argument("--headless")

    driver = webdriver.Chrome(executable_path="D:\\Pavel\\chromedriver\\chromedriver.exe",chrome_options=options)
    
    res = connect_to_printer(url, driver)
    time.sleep(1)
  
    if res == 0:
        print('Не удалось получить доступ к принтеру')
        driver.close()
        driver.quit()
        return
    
    mail_box_list = list_mail_box(driver)
    mail_box_index_list = []
    if mail_box_list:
        print('Всего найдено ящиков:', len(mail_box_list))
        for el in mail_box_list:
            mail_box_index_list.append(int(el[0:3]))
        print(mail_box_index_list)
    else:
        print('Ящики не найдены')
        
    res = 0
    
    while res == 0:
        print('Выберите номер почтового ящика или введите "quit" для выхода')
        num = input()
        if num == 'quit':
            driver.close()
            driver.quit()
            return
        try:
            num = int(num)
        except:
            print('Не удалось определить номер почтового ящика')
            continue
        if num not in mail_box_index_list:
            print('Номер ящика неактивен')
            continue
        res = choice_mail_box(driver, num)
        if res == 0:
            print('Не удалось получить доступ к почтовому ящику')
            continue
    
    
    while True:  
        res = list_files(driver)
        if res:
            print('Выберите номер файла для скачивания от 1 до ',len(res),' или введите "quit" для выхода')
            num = input()
            if num == 'quit':
                break
            try:
                num = int(num)
            except:
                print('Не удалось определить номер файла')
                continue
            if num>len(res) or num<0:
                print('введите корректный номер файла')
                continue    
            download_file(driver, num)
        else:
            print('файлы в почтовом ящике не найдены')
        
    
    driver.close()
    driver.quit()
    
if __name__ == '__main__':
    main()