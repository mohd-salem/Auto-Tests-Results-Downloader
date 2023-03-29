from selenium import webdriver
from time import sleep
from datetime import datetime
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
import random
import os
from pyperclip import copy as clipCopy
from sys import argv
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
import csv
import sys
import time
from io import BytesIO
import datetime
import openpyxl
from os import fsync
import sys 
import easygui
options = webdriver.ChromeOptions()
options.add_argument("--disable-notifications")
options.add_argument("--disable-gpu")
options.add_argument("--silent")
#options.add_argument("--headless")  
options.add_argument("--start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
import pytesseract as tess
tess.pytesseract.tesseract_cmd = r'Tesseract-OCR\tesseract.exe'
from PIL import Image
import urllib.request
driver = webdriver.Chrome('chromedriver', options=options)

def connect():
    global driver
    global getpdf
    while (1):
        try:
            driver.get('URL')
            print("Login to the system")
            driver.find_element_by_xpath('//*[@id="O5B_id-inputEl"]').send_keys("*****")
            driver.find_element_by_xpath('//*[@id="O93_id-inputEl"]').clear()
            driver.find_element_by_xpath('//*[@id="O93_id-inputEl"]').send_keys("***")
            driver.find_element_by_xpath('//*[@id="O5F_id-inputEl"]').send_keys("******")
            while True :
                captchaFun()
        except:
            sleep(3)
            print('Logged in')
        break
    sleep(8)
    driver.find_element_by_xpath('//*[@id="O1F7_id"]').click()
    sleep(8)
    getpdf()



def captchaFun() :
            loc=driver.find_element_by_xpath('//*[@id="O8B_id"]/img')
            # image = driver.get_screenshot_as_png() 
            src = loc.get_attribute('src')
            urllib.request.urlretrieve(src, "captcha.webp")
            # im = Image.open(BytesIO(image))
            # im = im.crop((910, 520  , 1080, 560))
            # im.save('captcha.png')
            img = Image.open('captcha.webp').convert('RGB')
            text = tess.image_to_string(img)
            print('text is ' + text)
            driver.find_element_by_xpath('//*[@id="O87_id-inputEl"]').send_keys(text)
            sleep(2)
            if Alert :
                try:
                    driver.switch_to.alert.accept()
                    captchaFun()
                except NoAlertPresentException:
                    print("exception hanlded")

            

def getpdf():
    datetimeF = datetime.datetime.now().strftime('%Y-%m-%d _ %H')
    new_dir = ".//" + datetimeF
    if not os.path.exists(new_dir):
        os.makedirs(new_dir)
    container = driver.find_elements_by_xpath("//div[@class='x-panel x-tabpanel-child x-panel-default']")
    table_id = container[2].find_elements_by_tag_name('table')
    for x in range (0,len(table_id)) :
        try:
            table_id = container[2].find_elements_by_tag_name('table')
            table_id[x].click()
            name_path = table_id[x].get_attribute('id')
            first_Name = driver.find_element_by_xpath('//*[@id="'+str(name_path)+'"]//tbody//tr//td[5]//div').text
            last_Name = driver.find_element_by_xpath('//*[@id="'+str(name_path)+'"]//tbody//tr//td[6]//div').text
            print(first_Name +' '+ last_Name)
            driver.find_element_by_xpath('//*[@id="O2AE_id"]').click()
            sleep(8)
            src = driver.find_elements_by_xpath('//iframe[@class="x-component x-abs-layout-item x-component-default"]')[1].get_attribute('src')
            fullfilename = os.path.join(datetimeF, str(first_Name)+" "+str(last_Name)+" "+str(x)+".pdf")
            urllib.request.urlretrieve(src,fullfilename ) 
            try :
                close_button = driver.find_element_by_xpath("/html/body/div[6]/div[1]/div/div/div[4]")
                close_button.click()
            except:
                close_button1 = driver.find_element_by_xpath("/html/body/div[8]/div[1]/div/div/div[4]")
                close_button1.click()
            print('Download has been Completed ')
            print('#*'*50)
            sleep(2)
            try:
                driver.find_element_by_xpath('//*[@id="O2BE_id"]').click()
                sleep(3)
                srcN = driver.find_elements_by_xpath('//iframe[@class="x-component x-abs-layout-item x-component-default"]')[1].get_attribute('src')
                fullfilename = os.path.join(datetimeF, str(first_Name)+" "+str(last_Name)+" "+"e nabız"+".pdf")
                urllib.request.urlretrieve(srcN, fullfilename)
                
                print(' E-nabiz has been downloaded ')
                print('*'*50)
                sleep(2)
                close_button = driver.find_element_by_xpath("/html/body/div[6]/div[1]/div/div/div[4]")
                close_button.click()
            except :
                driver.find_element_by_xpath('/html/body/div[6]/div[2]/div[2]/div/div/a[1]').click()
                print('  E-nabiz not found ')
                print('-'*25) 
        except :
            print('Eror while downloading -_-')
            try:
                close_button = driver.find_element_by_xpath("/html/body/div[6]/div[1]/div/div/div[4]")

                close_button.click()
            except:
                try:
                    close_button = driver.find_element_by_xpath("/html/body/div[8]/div[1]/div/div/div[4]")
                    close_button.click()
                except:
                    print ('No Windows to close it')
                print ('No Windows to close it')
            print('-'*50)







def downloader() :
    all_Patients =[]
    container = driver.find_elements_by_xpath("//div[@class='x-panel x-tabpanel-child x-panel-default']")
    table_id = container[2].find_elements(By.TAG_NAME, "table")
    for row in table_id :
        try:
            name_path = row.get_attribute('id')
            first_Name = driver.find_element_by_xpath('//*[@id="'+str(name_path)+'"]//tbody//tr//td[5]//div').text
            last_Name = driver.find_element_by_xpath('//*[@id="'+str(name_path)+'"]//tbody//tr//td[6]//div').text
            full_Name =[first_Name,last_Name]
            #print(full_Name)
            if full_Name in all_Patients :
                #print('patient exsits')
                continue
            else:
                all_Patients.append (full_Name)
        except:
            continue
    print(all_Patients)
    print ('there is '+ str( len(all_Patients)) + ' patients today')


        


def oldpdf():
    
    low = input("Enter the lowest value ")
    high = input("Enter the highest value ")
    low=int(low)
    high=int(high)
    
    
    for x in range (low,high):
        try:
            v = "gridview-1039"
            xpat = '//*[@id="'
            xpat += str(v)     
            xpat += '-record-'
            xpat += str(x)
            xpat += '"]'
            print (xpat)
            driver.find_element_by_xpath(xpat).click()
            driver.find_element_by_xpath('//*[@id="O25D_id"]').click()
            sleep(8)
            src = driver.find_element_by_xpath('/html/body/div[11]/div[2]/div/div/span/div/div[1]/div/span/div/iframe').get_attribute('src')
            z= str(xpat)
            z+= "/td[5]/div"
            print(z)
            y=driver.find_element_by_xpath(z).text
            print(y)
            t= str(xpat)
            t+= "/td[6]/div"
            print(t)
            r=driver.find_element_by_xpath(t).text
            print(y,' ',r)
            urllib.request.urlretrieve(src, str(y)+" "+str(r)+" "+str(x)+".pdf")
            print(src)
            driver.find_element_by_xpath('/html/body/div[11]/div[1]/div/div/div/div[4]/img').click()
            sleep(5)
            print('Download has been Completed ')
            print('-'*50)
            sleep(2)
            try:
                driver.find_element_by_xpath('//*[@id="O26D_id-btnIconEl"]').click()
                sleep(5)
                srcN = driver.find_element_by_xpath('/html/body/div[11]/div[2]/div/div/span/div/div[1]/div/span/div/iframe').get_attribute('src')
                
                print(srcN)
                urllib.request.urlretrieve(srcN, str(y)+" "+str(r)+" "+"e nabız"+".pdf")
                print(' E-nabiz has been downloaded ')
                print('-'*50)
                sleep(2)
                driver.find_element_by_xpath('/html/body/div[11]/div[1]/div/div/div/div[4]/img').click()
            except :
                driver.find_element_by_xpath('//*[@id="button-1005-btnIconEl"]').click()
                print(' E-nabiz not found ')
                print('-'*50)
    
        except :
            print('Eror while downloading -_- ')
            print('-'*50)
def gettestbefore():
    low = input("Enter the lowest value ")
    high = input("Enter the highest value ")
    low=int(low)
    high=int(high)
    
    
    for x in range (low,high):
        v = "gridview-1039"
        xpat = '//*[@id="'
        xpat += str(v)     
        xpat += '-record-'
        xpat += str(x)
        xpat += '"]'
        print (xpat)
        driver.find_element_by_xpath(xpat).click()
        z=0
        while True :
            try:
                u="/html/body/div[8]/div/span/div/div[1]/div/div/div[2]/div[2]/div/span/div/fieldset[2]/div/span/div/div[2]/div/div/div[2]/div[1]/div/span/div/div/div[2]/div/table/tbody/tr["
                u+= str (z)
                u+= "]/td[1]/div"
                print(u)
                g=driver.find_element_by_xpath(u).text
                print (g)
                z=z+1
            except:
                break
        
    #vk = openpyxl.Workbook()
    #sh = vk.active
    #sh.title="test"
    #sh ['A4'].value=g
    #vk.save("tests.xlsx")    
def gettestafter():
    low = input("Enter the lowest value ")
    high = input("Enter the highest value ")
    low=int(low)
    high=int(high)
    
    
    for x in range (low,high):
        v = "gridview-1039"
        xpat = '//*[@id="'
        xpat += str(v)     
        xpat += '-record-'
        xpat += str(x)
        xpat += '"]'
        print (xpat)
        driver.find_element_by_xpath(xpat).click()
        z=0
        while True :
            try:
                u="/html/body/div[8]/div/span/div/div[1]/div/div/div[2]/div[2]/div/span/div/fieldset[2]/div/span/div/div[2]/div/div/div[2]/div[2]/div/span/div/div/div[2]/div/table/tbody/tr["
                u+= str (z)
                u+= "]/td[1]/div"
                g=driver.find_element_by_xpath(u).text
                print (g)
                z=z+1
            except:
                break
        
    vk = openpyxl.Workbook()
    sh = vk.active
    sh.title="test"
    sh ['A4'].value=g
    vk.save("tests.xlsx")
def refresh():
    for x in range (0,50):
        driver.find_element_by_xpath('//*[@id="O1C2_id-btnIconEl"]').click()
        sleep(300)
def window_close() : 
    driver.close()




if __name__ == '__main__' :
    print('Commands')
    print('-'*50)
    print('For login type  ')
    print('connect()')
    print('')
    print('For download The Pdfs type ')
    print('getpdf()')
    print('')
    connect()


##    app = QtWidgets.QApplication(sys.argv)
##    Form = QtWidgets.QWidget()
##    ui = Ui_Form()
##    ui.setupUi(Form)
##    Form.show()
##    sys.exit(app.exec_())    












def grd():
    datetimeF = datetime.datetime.now().strftime('%Y-%m-%d _ %H')
    new_dir = "C://Users//NoVa//OneDrive//Desktop//py//PDFs//" + datetimeF
    if not os.path.exists(new_dir):
        os.makedirs(new_dir)
    container = driver.find_elements_by_xpath("//div[@class='x-panel x-tabpanel-child x-panel-default']")
    table_id = container[2].find_elements_by_tag_name('table')
    for x in range (0,len(table_id)) :
                table_id = container[2].find_elements_by_tag_name('table')
                table_id[x].click()
                name_path = table_id[x].get_attribute('id')
                first_Name = driver.find_element_by_xpath('//*[@id="'+str(name_path)+'"]//tbody//tr//td[5]//div').text
                last_Name = driver.find_element_by_xpath('//*[@id="'+str(name_path)+'"]//tbody//tr//td[6]//div').text
                print(first_Name +' '+ last_Name)
                driver.find_element_by_xpath('//*[@id="O286_id"]').click()
                sleep(8)
                src = driver.find_elements_by_xpath('//iframe[@class="x-component x-abs-layout-item x-component-default"]')[1].get_attribute('src')
                fullfilename = os.path.join(datetimeF, str(first_Name)+" "+str(last_Name)+" "+str(x)+".pdf")
                urllib.request.urlretrieve(src,fullfilename )
                sleep(5)
                close_button = driver.find_elements_by_id("O1C38_id-btnIconEl")
                close_button[0].click()
                sleep(5)
                print('Download has been Completed ')
                print('#*'*50)
                sleep(2)
                try:
                    driver.find_element_by_xpath('//*[@id="O296_id"]').click()
                    sleep(5)
                    srcN = driver.find_elements_by_xpath('//iframe[@class="x-component x-abs-layout-item x-component-default"]')[1].get_attribute('src')
                    fullfilename = os.path.join(datetimeF, str(first_Name)+" "+str(last_Name)+" "+"e nabız"+".pdf")
                    urllib.request.urlretrieve(srcN, fullfilename)
                    
                    print(' E-nabiz has been downloaded ')
                    print('*'*50)
                    sleep(2)
                    close_button = driver.find_elements_by_id("O1C38_id-btnIconEl")
                    close_button[0].click()
                except :
                    driver.find_element_by_xpath('/html/body/div[5]/div[2]/div[2]/div/div/a[1]').click()
                    print('  E-nabiz not found ')
                    print('-'*25) 




















# class ExampleWindow(QWidget):
#     def __init__(self):
#         super().__init__()

#         self.setup()

#     def setup(self):
#         btn_quit = QPushButton('Force Quit', self)
#         btn_quit.clicked.connect(QApplication.instance().quit)
#         btn_quit.resize(btn_quit.sizeHint())
#         btn_quit.move(325, 375)
#         btn_getPdf = QPushButton('download pdf', self)
#         btn_getPdf.clicked.connect(connect)
#         # btn_getPdf.clicked.connect()
#         btn_getPdf.move(0, 375)






#         self.setGeometry(200, 200, 400, 400)
#         self.setWindowTitle('Umut PDF Downloader')

#         self.show()

    # def closeEvent(self, event: QCloseEvent):
    #     reply = QMessageBox.question(self, 'Message', 'Are you sure you want to quit?',
    #                                  QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

    #     if reply == QMessageBox.Yes:
    #         event.accept()
    #     else:
    #         event.ignore()

        
# def run():
#     app = QApplication(sys.argv)

#     ex = ExampleWindow()

#     sys.exit(app.exec())

