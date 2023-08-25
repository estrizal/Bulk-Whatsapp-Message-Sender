import winreg
from zipfile import ZipFile
import os.path
import sys
is_64 = sys.maxsize > 2**32
if is_64 == True:
    print("your compouter is 64 bit")
else:
    print("Your computer is 32 bit")


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


MS_VERSION = get_registry_value(r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Edge\BLBeacon","version")
print('your edge version is')
print(MS_VERSION)

def update_driver(MS):
    print("downloading the edgedriver for this version")
    import urllib.request
    if is_64 == False:
        urllib.request.urlretrieve("https://msedgedriver.azureedge.net//"+str(MS)+"/edgedriver_win32.zip", "edgedriver_win32.zip")
    
    if is_64 == True:
        try:
            urllib.request.urlretrieve("https://msedgedriver.azureedge.net//"+str(MS)+"/edgedriver_win64.zip", "edgedriver_win32.zip")
        except:
            urllib.request.urlretrieve("https://msedgedriver.azureedge.net//"+str(MS)+"/edgedriver_win32.zip", "edgedriver_win32.zip")


    with ZipFile('edgedriver_win32.zip','r') as zip:
        zip.printdir()
        print("extracting...")
        zip.extractall()
        print("Extracted")

    filen = open(r"C:\ProgramData\msedge_version.txt", 'w+')
    filen.write(MS_VERSION)
    print("ready to go!")

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
            update_driver(MS_VERSION)

    
    else:
        update_driver(MS_VERSION)
except Exception as i:
    print(i)
    update_driver(MS_VERSION)
