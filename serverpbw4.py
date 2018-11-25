import socketserver
import subprocess
import os
import pyautogui
import time

class MyTCPHandler(socketserver.BaseRequestHandler):
    """
    The RequestHandler class for our server.

    It is instantiated once per connection to the server, and must
    override the handle() method to implement communication to the
    client.
    """

    def handle(self):
        # self.request is the TCP socket connected to the client
        self.data = self.request.recv(1024).strip()
        print("{} wrote:".format(self.client_address[0]))
        print(self.data)
        global ff
        if self.data == b'Internetexplorer':    
            print('opening internet explorer')
            subprocess.Popen(r'"C:\Program Files\Internet Explorer\IEXPLORE.EXE" www.extron.com')  #may need to change your file location
        elif self.data == b'PowerPoint':
            print('opening powerpoint')
            subprocess.Popen(r'"C:\Program Files (x86)\Microsoft Office\Office15\POWERPNT.EXE" /s Control.ppsx') # may need to change your location
        elif self.data == b'Firefox':
            print('opening firefox')
            subprocess.Popen(r'"C:\Program Files (x86)\Mozilla Firefox\firefox.exe" www.google.com') # may need to change your location
            time.sleep(.1)
            pyautogui.click()
        elif self.data == b'FirefoxSMP351':
            print('opening firefoxsmp351')
            subprocess.Popen(r'"C:\Program Files (x86)\Mozilla Firefox\firefox.exe" http://192.168.254.51') # may need to change your location
            time.sleep(.1)
            pyautogui.click()
        elif self.data == b'FirefoxOCGallery':
            print('opening firefoxocgallery')
            subprocess.Popen(r'"C:\Program Files (x86)\Mozilla Firefox\firefox.exe" http://192.168.254.2:8080/engage/ui/index.html') # may need to change your location
            time.sleep(.1)
            pyautogui.click()
        elif self.data == b'FirefoxOCScheduler':
            print('opening firefoxocscheduler')
            subprocess.Popen(r'"C:\Program Files (x86)\Mozilla Firefox\firefox.exe" http://192.168.254.2:8080/admin/index.html#/scheduler') # may need to change your location
            time.sleep(.1)
            pyautogui.click()
        elif self.data == b'FirefoxOCRecordings':
            print('opening firefoxocrecordings')
            subprocess.Popen(r'"C:\Program Files (x86)\Mozilla Firefox\firefox.exe" http://192.168.254.2:8080/admin/index.html#/recordings') # may need to change your location
            time.sleep(.1)
            pyautogui.click()
        elif self.data == b'VLCArchive':
            print('opening vlc')
            subprocess.Popen(r'"C:\Program Files (x86)\VideoLAN\VLC\vlc.exe" "C:\VLC\Archive.xspf"') # may need to change your location
            time.sleep(.1)
        elif self.data == b'VLCConfidence':
            print('opening vlc')
            subprocess.Popen(r'"C:\Program Files (x86)\VideoLAN\VLC\vlc.exe" "C:\VLC\Confidence.xspf"') # may need to change your location
            time.sleep(.1)
        elif self.data == b'CloseFF':
            print('closing firefox')
            os.system("taskkill /F /IM firefox.exe")
        elif self.data == b'CloseFFGently':
            print('closing firefox gently')
            pyautogui.press('alt')
            time.sleep(.5)
            pyautogui.press('f')
            time.sleep(.5)
            pyautogui.press('x')
            time.sleep(.5)
        elif self.data == b'CloseVLC':
            print('closing vlc')
            os.system("taskkill /F /IM vlc.exe")
        elif self.data == b'ClosePP':
            print('closing powerpoint')
            os.system("taskkill /F /IM powerpnt.exe")
        elif self.data == b'CloseIE':
            print('closing internet explorer')
            os.system("taskkill /F /IM iexplore.exe")
        elif self.data == b'CloseVLC':
            print('closing vlc')
            os.system("taskkill /F /IM vlc.exe")
        elif self.data == b'NextPP':
            pyautogui.press('right')
        elif self.data == b'PrevPP':
            pyautogui.press('left')
        elif self.data == b'PPStart':
            pyautogui.press('1')
            time.sleep(.1)
            pyautogui.press('enter')
        elif self.data == b'PP1':
            pyautogui.press('2')
            time.sleep(.1)
            pyautogui.press('enter')
        elif self.data == b'PP2':
            pyautogui.press('3')
            time.sleep(.1)
            pyautogui.press('enter')
        elif self.data == b'PP3':
            pyautogui.press('4')
            time.sleep(.1)
            pyautogui.press('enter')
        elif self.data == b'PP4':
            pyautogui.press('5')
            time.sleep(.1)
            pyautogui.press('enter')
        elif self.data == b'PP5':
            pyautogui.press('6')
            time.sleep(.1)
            pyautogui.press('enter')
        elif self.data == b'CloseFileManager':
            pyautogui.press('alt')
            time.sleep(.1)
            pyautogui.press('f')
            time.sleep(.1)
            pyautogui.press('c')
            time.sleep(.1)
            pyautogui.press('enter')
        
        # just send back the same data, but upper-cased
        self.request.sendall(self.data.upper())

if __name__ == "__main__":
    HOST, PORT = "", 1977   # change 1977 to whatever port you would like to use

    # Create the server, binding to localhost on port specified
    server = socketserver.TCPServer((HOST, PORT), MyTCPHandler)

    # Activate the server; this will keep running until you
    # interrupt the program with Ctrl-C
    server.serve_forever()
