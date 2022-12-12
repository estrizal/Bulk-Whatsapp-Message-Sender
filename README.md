# Whatsapp-message-sender-open-source

YOU CAN DOWNLOAD THE EXECUTABLE FILE OF THIS PROJECT FROM HERE:- https://sites.google.com/view/estriadi-portfolio-1/bulk-whatsapp-sender?authuser=0
I have just converted python code to exe and then use inno setup compiler to make an installer for distribution.

This project sends whatsapp messages to phone numbers stored in the first column of an excel file.

Technical stuff down here:) :-

This whatsapp sender.py is the main file of this project. The project uses selenium library to automate the process of sending whatsapp messages.

If you have worked with selenium, you would know that you have to download a webdriver for the suitable version of your browser. so this app first checks the version of your microsoft edge from your registry and checks if the currect version of webdriver matchs, if not it runs the edge update.exe file (which is just the converted version of edge_update.py) and download the latest version.

there is also an updater to update the app whenever I launch one. Firebase is used for authentication and launching update. 

In this project I have used multiprocessing library, to run multiple functions at onece.e

I Know the code a LOTTTT MESSYY and Without comments too, but I am lazy enough to not care XD
