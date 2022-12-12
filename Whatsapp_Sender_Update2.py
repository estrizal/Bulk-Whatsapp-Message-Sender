from logging import exception
import time
import pyrebase
from zipfile import ZipFile
#import struct
import shutil
import os
import socket
def checking():
    try:
        #socket.create_connection(('Google.com', 80))
        socket.create_connection(("1.1.1.1", 53))
        return True
    except Exception:
        return False


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
storage = Firebase.storage()
database = Firebase.database()
file = open(r"C:\ProgramData\WhatsApp_version.txt", 'r')
current_version = file.read()
file.close()
to_check = checking()
if to_check == True:
    Server_Version = database.get()
    Server_version = str(Server_Version.val()).replace("OrderedDict", "").replace("(", "").replace(")" ,"").replace("[" ,"").replace("]" ,"").replace("'" ,"").replace("Beta" ,"").replace("," ,"").replace("Yup" ,"").replace(" " ,"").replace("Version","")
    print(Server_version)
    Patch_server_version = float(Server_version)
    patch_currect_version = float(current_version)


    patch_checker = Patch_server_version-patch_currect_version
    if patch_checker < 0.099 and patch_checker > 0:
        #if patch_checker < 0.1:
     try:
        if current_version != Server_version:
            print("Whatsapp Sender PATCH, installing update DO NOT CLOSE THIS")
            Path_on_cloud = "Update/Whatsapp_Sender_patch.zip"
            print('download command sent')
            storage.child(Path_on_cloud).download('','Whatsapp_Sender_p.zip')
            file = "Whatsapp_Sender_p.zip"
            print("download of whatsapp sender patch")
            time.sleep(5)
            with ZipFile(file,'r') as zip:
                zip.printdir()
                print("extracting...")
                zip.extractall()
                print("Extracted")
                #os.popen("Delete.cmd")

                filen = open(r"C:\ProgramData\WhatsApp_version.txt", 'w')
                filen.write(Server_version)
                filen.close()
                print("Update of whatsapp Sender completed, kindly close this window")


     except Exception as I:
        print(I)
     
        '''
        Path_on_cloud = "Update/Whatsapp_Sender.exe"
        storage.child(Path_on_cloud).download('./','Whatsapp_Sender.exe')
        filen = open(r"C:\ProgramData\WhatsApp_version.txt", 'w')
        filen.write(Server_version)
        filen.close()
        print("Update of whatsapp Sender completed, kindly close this window")
        '''
    
    else:
     try:
        if current_version != Server_version:
            print("Whatsapp Sender new update available, installing update DO NOT CLOSE THIS")
            Path_on_cloud = "Update/Whatsapp_Sender.zip"
            print('download command sent')
            storage.child(Path_on_cloud).download('','Whatsapp_Sender_U.zip')
            file = "Whatsapp_Sender_U.zip"
            print("download of whatsapp sender new version completed")
            time.sleep(5)
            with ZipFile(file,'r') as zip:
                zip.printdir()
                print("extracting...")
                zip.extractall()
                print("Extracted")
                #os.popen("Delete.cmd")
                root_src_dir = './Whatsapp_Sender'
                root_dst_dir = './'
                for src_dir, dirs, files in os.walk(root_src_dir):
                    dst_dir = src_dir.replace(root_src_dir, root_dst_dir, 1)
                    if not os.path.exists(dst_dir):
                        os.makedirs(dst_dir)
                    for file_ in files:
                        src_file = os.path.join(src_dir, file_)
                        dst_file = os.path.join(dst_dir, file_)
                        if os.path.exists(dst_file):
                            # in case of the src and dst are the same file
                            if os.path.samefile(src_file, dst_file):
                                continue
                            os.remove(dst_file)
                        shutil.move(src_file, dst_dir)
                        filen = open(r"C:\ProgramData\WhatsApp_version.txt", 'w')
                        filen.write(Server_version)
                        filen.close()
                        print("Update of whatsapp Sender completed, kindly close this window")


     except Exception as I:
        print(I)
else:
    print("please connect to internet")


