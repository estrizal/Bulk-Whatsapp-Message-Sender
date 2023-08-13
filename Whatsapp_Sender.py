from lib2to3.pgen2 import driver
from email import message
from logging import exception
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QDialog
from PyQt5 import QtWidgets, QtCore, QtGui
import sys
from PyQt5.uic import loadUi
import time
import multiprocessing
from multiprocessing.context import Process
from multiprocessing import process
from ctypes import c_char_p
from PyQt5.uic.uiparser import WidgetStack
#from pyrebase.pyrebase import Database
import pyrebase #pip install pyrebase4 
#run it if the error occurs then pip uninstall crypto and install crypto then rename it to Crypto (C is in capital now) in site packages and then install pycrypto
import os
import socket
#source whatsapp imports below

import xlrd    #pip install xlrd==1.2.0
from selenium import webdriver  
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import pyautogui
import pyperclip
from io import BytesIO
#import ioS
import codecs
import win32clipboard #pip install pywin32
from PIL import Image
import docxpy
import winreg
import webbrowser







register_loaded = False
Logged_in = False
Register = None

DOCUMENT_Path = ""
Excel_path = ""
Photo_or_video_Path = ""
con_code_usage = True


def update_label():
    global User_interface
    global status_of_internet
    User_interface.value = widget.currentIndex()

    if status_of_internet.value == "connected":
        try:
            if User_interface.value == 0:
                main.internet.setHidden(True)
            if User_interface.value == 1:
                Login.internet.setHidden(True)
            if User_interface.value == 2:
                Register.internet.setHidden(True)
                pass
        except Exception:
            pass
    elif status_of_internet.value == "not":
        try:
            if User_interface.value == 0:
                main.internet.setVisible(True)
            if User_interface.value == 1:
                Login.internet.setVisible(True)
            if User_interface.value == 2:
                Register.internet.setVisible(True)
        except Exception:
            pass



class Main(QMainWindow): #,FROM_MAIN):
    global DOCUMENT_Path
    global Excel_path
    global Photo_or_video_Path
    global Logged_in
    global pho_or_doc
    global con_code_usage


    def __init__(self,parent=None):
        super(Main,self).__init__(parent)
        loadUi("Main_UI.ui",self)
        
        #self.setupUi(self)
        widget.setFixedSize(1109,721)
        self.sign_out.clicked.connect(self.gotologin)
        #self.whatshesaid_2.setText("<font size=12 color='#4ac8eb'>"+ "listening" +"</font>")
        self.internet.setHidden(True)
        self.START.clicked.connect(self.start_sending)
        self.Message.clicked.connect(self.document_browse)
        self.Excel.clicked.connect(self.excel_browse)
        self.Attachment.clicked.connect(self.Photo_browse)
        self.START.clicked.connect(self.start_sending)
        self.Rem_code.clicked.connect(self.rem_code)
        self.TO_USE.clicked.connect(self.USE_IT)
        self.Fix.clicked.connect(self.Fix_it)
        self.checkbox.stateChanged.connect(lambda:self.checked())        

        try:
            file = open(r"C:\ProgramData\con_code.txt", 'r')
            con_code_Content = file.read()
            self.country_code.setPlainText(con_code_Content)
        except:
            pass

    def checked(self):
        global con_code_usage
        if self.checkbox.isChecked() == True:
            self.label_4.setHidden(True)
            self.country_code.setHidden(True)
            self.Rem_code.setHidden(True)
            con_code_usage = False
        else:
            self.label_4.setHidden(False)
            self.country_code.setHidden(False)
            self.Rem_code.setHidden(False)
            con_code_usage = True

    def Fix_it(self):
        webbrowser.open('https://sites.google.com/view/estriadi-portfolio-1/bulk-whatsapp-sender/common-fixes?authuser=2')
    
    def USE_IT(self):
        webbrowser.open('https://www.youtube.com/watch?v=RT-ZnvpnIRU')  # Go to example.com
        
    def gotologin(self):
        global Logged_in

        if Logged_in == False:
            Login = login()
            widget.addWidget(Login)

        my_shine = open(r"C:\ProgramData\Whatsapp_sender_Logged_in.txt", "w")
        my_shine.write("N")
        my_shine.close()

        widget.setCurrentIndex(widget.currentIndex() +1)


    def document_browse(self):
        global DOCUMENT_Path
        fname = QFileDialog.getOpenFileName(self,"Open Docx", "Documents", "document files (*.docx)")
        DOCUMENT_Path = fname[0]
        self.Message.setStyleSheet("background-color : #00BFFF")
        #print(fname[0])

    def excel_browse(self):
        global Excel_path
        fname = QFileDialog.getOpenFileName(self,"Open xlsx", "Excel File With Phonenumbers", "excel files (*.xlsx)")
        Excel_path = fname[0]
        self.Excel.setStyleSheet("background-color : #00BFFF")

    def Photo_browse(self):
        global Photo_or_video_Path
        global pho_or_doc
        jojo = self.option.currentText()
        if jojo == "Photo/video":
            pho_or_doc = 0
        else:
            pho_or_doc = 1

        if pho_or_doc == 0:
            fname = QFileDialog.getOpenFileName(self,"Open Photo/video", "Photo Or Video")

        if pho_or_doc == 1:
            fname = QFileDialog.getOpenFileName(self,"Open file to send", "File to send")
        
        Photo_or_video_Path = fname[0]
        self.Attachment.setStyleSheet("background-color : #00BFFF")


    def start_sending(self):
      global DOCUMENT_Path
      global Excel_path
      global Photo_or_video_Path
      global con_code_usage

      def Read_docx(filepath):
        text = docxpy.process(filepath)
        return text

      try:
       i_Cant = False

       try:
           messege = str(Read_docx(DOCUMENT_Path))
       except:
           self.Error.setText("<font size=12 color='#ff0000'>"+"Message docx is empty or can't be read"+"</font>")
           self.Error.setHidden(False)
           i_Cant = True
           self.Message.setStyleSheet("background-color : #FF6347")



       try:
           wait = self.wait_starting_chat.toPlainText()
           wait = int(wait)
           self.wait_starting_chat.setStyleSheet("background-color : white")
            

       except:
           self.Error.setText("<font size=12 color='#ff0000'>"+"only type numbers in the waiting time field"+"</font>")
           self.Error.setHidden(False)
           i_Cant = True
           self.wait_starting_chat.setStyleSheet("background-color : #FF6347")

       if con_code_usage == True:
        try:
            con_code = self.country_code.toPlainText()
            con_code = int(con_code)
            self.country_code.setStyleSheet("background-color : white")

        except:
            self.Error.setText("<font size=12 color='#ff0000'>"+"Country Code cant be an alphabet."+"</font>")
            self.Error.setHidden(False)
            i_Cant = True
            self.country_code.setStyleSheet("background-color : #FF6347")
       else:
           pass

       try:
           loc = (Excel_path)
           wr=xlrd.open_workbook(loc) #opening excel file
           shee1 = wr.sheet_by_index(0) #sheet number
           num = shee1.cell_value(0,0) #cell value, ghs columb and row to get phone number
           num2 = str(num)  
           num2 = shee1.cell_value(0,0)
           num2 = str(num2)
           num2 = num2[:-2]

       except:
           self.Error.setText("<font size=12 color='#ff0000'>"+"excelfile does not have no. in A1 of 1st sheet by index"+"</font>")
           self.Error.setHidden(False)
           i_Cant = True
           self.Excel.setStyleSheet("background-color : #FF6347")




       if DOCUMENT_Path == '':
           self.Error.setText("<font size=12 color='#ff0000'>"+"Please browse the docx containing message to send"+"</font>")
           self.Error.setHidden(False)

       elif Excel_path == '':
           self.Error.setText("<font size=12 color='#ff0000'>"+"Please browse the xlsx file with phonenumbers to send"+"</font>")
           self.Error.setHidden(False)
           self.Excel.setStyleSheet("background-color : #FF6347")



       elif i_Cant == True:
           print('i_cant')
       else:
           
        if Photo_or_video_Path == '':
            self.Error.setText("<font size=12 color='#ff0000'>"+"No photo/video browsed, only message will be sent"+"</font>")
            self.Error.setHidden(False)
            print('no photo')









        self.Error.setHidden(True)

        wait = self.wait_starting_chat.toPlainText()
        wait = int(wait)

        con_code = self.country_code.toPlainText()
        con_code = int(con_code)


        def send_to_clipboard(clip_type, data):
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(clip_type, data)
            win32clipboard.CloseClipboard()

        #loc = (r'C:\Users\My Computer\Desktop\contacts.xls')   # THIS IS THE LOCATION OF THE XCEL FILE CONTAINING ALL THE CONTACTS
        loc = (Excel_path)

        ghs = 0  # THIS IS THE VARIABLE TO DEFINE THE COLUMB
        nsh = 0  # THIS VARIABLE IS USED A THE VALUE OF ROWS
        hf = 0   # THIS IS THE VARIABLE USED SO THAT THE WHILE LOOP CAN BE RAN MANY TIMES
        path = 1 # THIS IS THE VARIABLE ..... YOU WONT UNDERSTANT SO I AM NOT WRITING
        npf = 1  # THIS IS THE VARIABLE TO DEFINE HOW MANY TABS ARE OPENED RIGHT NOW + 1

        browser = webdriver.Edge('msedgedriver.exe') #webdriver.Chrome('D:\\chromedriver.exe')  # THIS IS THE LOCATION OF YOUR WEBDRIVER
        
        wr=xlrd.open_workbook(loc) #opening excel file
        shee1 = wr.sheet_by_index(0) #sheet number
        num = shee1.cell_value(ghs,nsh) #cell value, ghs columb and row to get phone number
        num2 = str(num)  #TYPECASTING A FLOAT VALUE IN A STRING SO THAT IT CAN BE ADDED TO THE URL

        messege = str(Read_docx(DOCUMENT_Path))
        pyperclip.copy(messege)
        browser.get('https://web.whatsapp.com/')
        load_whatsapp = WebDriverWait(browser, 1000000000000).until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))) #  //*[@id="app"]/div/div/div[3]/header/div[2]/div/span/div[3]/div  sent by subsciber
        time.sleep(3)
        while hf == hf:
            try:
                pyperclip.copy(messege)
                num2 = shee1.cell_value(ghs,nsh)
                num2 = str(num2)
                num2 = num2[:-2]
                if con_code_usage == True:
                    patanahi = 'https://web.whatsapp.com/send/?phone='+str(con_code)+num2+'&text&type=phone_number&app_absent=0'   #'https://wa.me/'+str(con_code)+num2
                if con_code_usage == False:
                    patanahi = 'https://web.whatsapp.com/send/?phone='+"+"+num2+'&text&type=phone_number&app_absent=0'

                browser.execute_script("window.open('');")
                browser.switch_to.window(browser.window_handles[npf])
                browser.get(patanahi)        
                print(patanahi)
                time.sleep(13)
                try:    
                        
                        myElem = WebDriverWait(browser, wait).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p'))) #//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p
                        typenum = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')  # MESSEGE WALA TEXT BOX TO FIND KR RHE HAI  //*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2] old sending message
                        time.sleep(0.8)
                        time.sleep(0.5)
                        typenum.click()

                        pyautogui.hotkey('ctrl', 'v')
                        typenum.send_keys(Keys.ENTER)
                        time.sleep(1)

                        pipicon = browser.find_element_by_xpath('//div[@title="Attach"]')
                        pipicon.click()
                        time.sleep(1)
                        filepath = Photo_or_video_Path
                        wait_for_photos_to_Open = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@accept="*"]')))
                        if pho_or_doc == 0:
                            phots = browser.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
                            phots.send_keys(filepath)
                        else:
                            phots = browser.find_element_by_xpath('//input[@accept="*"]')
                            phots.send_keys(filepath)

                        to_wait_for_send_button = myElem = WebDriverWait(browser, 1000000).until(EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]')))
                        sendbutton = browser.find_element_by_xpath('//span[@data-icon="send"]')# //*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]/p
                        sendbutton.click()
                        time.sleep(15)


                except Exception as identifier:
                    print(identifier)
                    time.sleep(3)
            except Exception as identifier:
                print(identifier)

            npf = npf+1
            if npf == 3:
                browser.switch_to.window(browser.window_handles[0])
                time.sleep(0.5)
                browser.close()
                browser.switch_to.window(browser.window_handles[0])
                time.sleep(0.5)
                browser.close()
                npf = npf-2
                browser.switch_to.window(browser.window_handles[0])
                time.sleep(0.5)
                
            else:
                pass    
            print("Sent till row no." + str(ghs))
            ghs = ghs+1
            hf=hf+1
            path = path+1

      except Exception as A:
          print(A)



    def rem_code(self):
        try:
            file = open(r"C:\ProgramData\con_code.txt", 'w+')
            con_code = self.country_code.toPlainText()
            try:
                for_an_error = int(con_code)
                file.write(con_code)

            except:
                print("The user tried to put an alphabet in number field")
                
        except:
            pass



class login(QDialog):
    global register_loaded
    global Register
    def __init__(self):
        super(login,self).__init__()
        loadUi("Log_in.ui",self)
        
        widget.setFixedSize(897,617)
        self.Register_button.clicked.connect(self.switchtoregister)
        self.Login_button.clicked.connect(self.login_to_firebase)
        self.internet.setHidden(True)


    def switchtoregister(self):
        global register_loaded
        global Register
        if register_loaded == False:
            Register = register()
            widget.addWidget(Register)
        if register_loaded == True:
            pass

        self.Email_input.setPlainText("")
        self.Password_input.setPlainText("")
        widget.setCurrentIndex(widget.currentIndex() +1)


    def login_to_firebase(self):
        
        global auth
        global Firebase

        email = self.Email_input.toPlainText()
        password = self.Password_input.toPlainText()

        email = email.replace('\n','').replace(' ','')


        password = password.replace('\n','').replace(' ','')
        try:
            login = auth.sign_in_with_email_and_password(email, password)
            my_file = open(r"C:\ProgramData\Whatsapp_sender_Logged_in.txt", "w")
            email_credentials = email
            my_file.write(email_credentials +"\n")
            my_file.write(password)
            my_file.close()
            self.Email_input.setPlainText("")
            self.Password_input.setPlainText("")

            widget.setCurrentIndex(widget.currentIndex() - 1)

            widget.setFixedSize(1109,721)
        except Exception as a:
            print(a)
            self.Hidden_label.setText("<font size=4 color='#ff0000'>"+"wrong or NOT registered email or WRONG password"+"</font>")



class register(QDialog):
    global register_loaded
    global auth
    global Firebase
    def __init__(self):
        super(register,self).__init__()
        loadUi("Register_user.ui",self)
        
        widget.setFixedSize(897,617)
        self.Login_Button.clicked.connect(self.gotologin)
        self.Register_button.clicked.connect(self.register_firebase)
        self.internet.setHidden(True)
    def gotologin(self):

        global register_loaded
        register_loaded = True
        self.Email_user_input.setPlainText("")
        self.Password_user_input.setPlainText("")
        self.Confirm_Passoword_user_input.setPlainText("")
        widget.setCurrentIndex(widget.currentIndex() -1)
    def register_firebase(self):
        global auth
        global Firebase
        email = self.Email_user_input.toPlainText()
        email = email.replace('\n','').replace(' ','')

        password = self.Password_user_input.toPlainText()
        password = password.replace('\n','').replace(' ','')

        confirm_pass = self.Confirm_Passoword_user_input.toPlainText()
        confirm_pass = confirm_pass.replace('\n','').replace(' ','')

        if password == confirm_pass:
            try:
                user = auth.create_user_with_email_and_password(email,password)
                self.Hidden_thing.setText("<font size=5 color='#55ff00'>"+"you are registered, plaease login"+"</font>")
            except Exception as identifier:
                print(identifier)
                self.Hidden_thing.setText("<font size=5 color='#ff0000'>"+"wrong or registered email or weak password"+"</font>")
        else:
            self.Hidden_thing.setText("<font size=5 color='#ff0000'>"+"Passwords not matched"+"</font>")



def Update_the_app(status_of_internet,User_interface):
    time.sleep(8)
    if status_of_internet.value=="connected":
        file = open(r"C:\ProgramData\WhatsApp_version.txt", 'r')
        current_version = file.read()
        file.close()
        print(current_version)
        database = Firebase.database()
        Server_Version = database.get()
        Server_version = str(Server_Version.val()).replace("OrderedDict", "").replace("(", "").replace(")" ,"").replace("[" ,"").replace("]" ,"").replace("'" ,"").replace("Beta" ,"").replace("," ,"").replace("Yup" ,"").replace(" " ,"").replace("Version","")
        print(Server_version)
        if current_version != Server_version:
            os.startfile('Whatsapp Sender updaterr.bat')
            os.system("TASKKILL /F /IM Sanchar_Whatsapp_Sender.exe")
        else:
            def get_registry_value(path, name="", start_key = None):
                if isinstance(path, str):
                    path = path.split("\\")
                if start_key is None:
                    start_key = getattr(winreg, path[0])
                    return get_registry_value(path[1:], name, start_key)
                else:
                    subkey = path.pop(0)
                with winreg.OpenKey(start_key, subkey) as handle:
                    assert handle
                    if path:
                        return get_registry_value(path, name, handle)
                    else:
                        desc, i = None, 0
                        while not desc or desc[0] != name:
                            desc = winreg.EnumValue(handle, i)
                            i += i
                        return desc[1]



            def update_driver():
                os.startfile('Whatsapp Sender Driver Updater.bat')


            MS_VERSION = get_registry_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Edge\BLBeacon","version")
            print('your edge version is')
            print(MS_VERSION)

        try:
            file = open(r"C:\ProgramData\msedge_version.txt", 'r')
            current_version = file.read()
            current_version = current_version.replace("\n","").replace(" ","")
            print(current_version)
            if str(current_version) == str(MS_VERSION):
                if os.path.isfile( 'msedgedriver.exe'):
                    print("driver is upto date and file is in the right place")
                else:
                    print("the software might have been reinstalled, thus msedge is not there. Downloading file")
                    update_driver()

            
            else:
                update_driver()
        except Exception as i:
            print(i)
            update_driver()




        

def Internet(status_of_internet,User_interface):
    time.sleep(3)
    while True:
        time.sleep(1)
        try:
            socket.create_connection(("1.1.1.1", 53))
            status_of_internet.value = "connected"
        except Exception:
            status_of_internet.value = "not"



firebaseConfig = {

    "apiKey": "AIzaSyBxCHXdgQAkbGygN5LAQuVCnDUFGSv3Dc8",
    "authDomain": "whatsapp-sender-7a064.firebaseapp.com",
    "projectId": "whatsapp-sender-7a064",
    "storageBucket": "whatsapp-sender-7a064.appspot.com",
    "messagingSenderId": "766117771099",
    "appId": "1:766117771099:web:22d45ad330e41d32ff4c7e",
    "measurementId": "G-BGM7ZM4CEX",
    "databaseURL": "https://whatsapp-sender-7a064-default-rtdb.firebaseio.com"
}

Firebase = pyrebase.initialize_app(firebaseConfig)
auth = Firebase.auth()
storage = Firebase.storage()
database = Firebase.database()

flags = QtCore.Qt.WindowFlags(QtCore.Qt.FramelessWindowHint)





if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = QtWidgets.QApplication(sys.argv)
    try:
        file1 = open(r"C:\ProgramData\Whatsapp_sender_Logged_in.txt","r")
    except Exception:
        file1 = open(r"C:\ProgramData\Whatsapp_sender_Logged_in.txt","w+")
    login_state = file1.read()

    widget=QtWidgets.QStackedWidget()
    main = Main()
    widget.addWidget(main)
    #widget.setFixedSize(1381,801)
    widget.show()
    to_check = open(r"C:\ProgramData\Whatsapp_sender_Logged_in.txt")
    content = to_check.read()
    to_check.close()

    if content == "N":
        Login = login()
        widget.addWidget(Login)
        widget.setCurrentIndex(widget.currentIndex() +1)
    else:
        try:
            to_check3 = open(r"C:\ProgramData\Whatsapp_sender_Logged_in.txt")
            string_list = to_check3.readlines()
            to_check3.close()
            email_id = str(string_list[0])
            email_id = email_id.replace("\n","")
            print(email_id)
            #print(email_id)
            password_cre = str(string_list[1])
            logggggin = auth.sign_in_with_email_and_password(email_id,password_cre)
        except Exception as I:
            print(I)
            to_check2 = open(r"C:\ProgramData\Whatsapp_sender_Logged_in.txt", "w+")
            to_check2.write("N")
            to_check2.close()
            Login = login()
            widget.addWidget(Login)
            widget.setCurrentIndex(widget.currentIndex() +1)


    manager = multiprocessing.Manager()
    User_interface = manager.Value(c_char_p, 0)
    status_of_internet = manager.Value(c_char_p, "PRESS HOTKEY")


    p1 = Process(target=Internet, args=(status_of_internet, User_interface))
    p1.start()

    p2 = Process(target=Update_the_app, args=(status_of_internet, User_interface))
    p2.start()



    timer = QtCore.QTimer()
    timer.timeout.connect(update_label)
    timer.start(100)

    sys.exit(app.exec_())
