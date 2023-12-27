# Whatsapp-message-sender-open-source


*MOST IMPORTANT NOTE* :- This whatsapp sender has stopped working, as the whatsapp is probably checking some advance pattern detectors to detect automation. Though, it might still work in some countries. but it's not working in india atleast. as of now. 

Here's a working demo of the time when this code used to work:- https://www.youtube.com/watch?v=RT-ZnvpnIRU&t=3s

YOU CAN DOWNLOAD THE EXECUTABLE FILE OF THIS PROJECT FROM HERE:- https://sites.google.com/view/estriadi-portfolio-1/bulk-whatsapp-sender?authuser=0
I have just converted python code to exe and then use inno setup compiler to make an installer for distribution.
This project sends whatsapp messages to phone numbers stored in the first column of an excel file.



Technical stuff down here:) :-

* Just run the "INSTALL LIBRARIES.cmd" file to install all the libraries

* fist step will be to make a txt file named 'WhatsApp_version.txt' and put it in C:\ProgramData if you cant see this folder, then enable show hidden folder option in your folder settings in control panel. 

* This whatsapp sender.py is the main file of this project. The project uses selenium library to automate the process of sending whatsapp messages.

this app first checks the version of your microsoft edge from your registry and checks if the currect version of webdriver matchs, if not it runs the edge update.exe file (which is just the converted version of edge_update.py) and download the latest version.

there is also an updater to update the app whenever I launch one. Firebase is used for authentication and launching update. 

In this project I have used multiprocessing library, to run multiple functions at onece.e

I Know the code a LOTTTT MESSYY and Without comments too.. But I believe in your abilities.
