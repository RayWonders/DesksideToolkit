import sys
import subprocess
import ctypes
import os
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from BiosSledgehammer import *

app = QtWidgets.QApplication(sys.argv)

class Ui(QtWidgets.QMainWindow):

    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('DesksideToolkit.ui', self)
        self.show()
        self.setWindowTitle("Robb's Deskside Toolkit :) ")
        textEdit = QTextEdit() 
        textEdit.setLineWrapMode(QTextEdit.WidgetWidth)

    #Slot signal for Headset repairs
    def headsetconfwin(self):
        self.w = headsetconfwin()
        self.w.show()
        

    #slot window for Windows repairs
    def windowsconfwin(self):
        self.wcw = windowsconfwin()
        self.wcw.show()
        

    #Slot for Manually installed apps
    def manconfwin(self):
        self.mcw = manconfwin()
        self.mcw.show()
        

    #Slot for BIOSSledgehammer start window
    def biossledge(self):
        self.bsh = biossledge()
        self.bsh.show()
        

    #Slot for Driver selection
    def driverconfwin(self):
        self.dcw = driverconfwin()
        self.dcw.show()

    #Slot for stores (WIP)
    def projectsconfwin(self):
        self.scw = projectsconfwin()
        self.scw.show()
    
    #Slot for macOS (WIP)
    def macconfwin(self):
        self.mcw = macconfwin()
        self.mcw.show()

    #Slot for Suite (WIP)
    def suiteconfwin(self):
        self.oscw = suiteconfwin()
        self.oscw.show()
        

#window to activate DISM/SFC repair w checkbox
class windowsconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(380, 200)
        self.setWindowTitle("Windows 10 repair (Fingers Crossed)")

        #checkbox for Restart function
        self.winchk = QCheckBox(self)
        self.winchk.setGeometry(30, 140, 350, 55)
        self.winchk.setText('restart when completed? (recommended)')
        self.winchk.setStyleSheet('font-size:15px')
        self.winchk.setChecked(False)

        #"Start" Button for SFC/DISM
        self.winstart = QPushButton(self)
        self.winstart.setGeometry(115, 25, 160, 100)
        self.winstart.setText('Start')
        self.winstart.setStyleSheet('font-size:25px')
        self.winstart.clicked.connect(self.run_sfc_scannow)

        #Connect signal so program knows if checkbox state changes and what script to run in conclusion       
        self.winchk.stateChanged.connect(self.chkboxchange)

    def chkboxchange(self):
        if self.winchk.isChecked():
            self.winstart.clicked.disconnect() # Disconnect the clicked signal from all its connections
            # Connect the clicked signal to the run_sfc_scannow_shutdown function
            self.winstart.clicked.connect(self.run_sfc_scannow_shutdown)
        else:
            self.winstart.clicked.disconnect() # Disconnect the clicked signal from all its connections
            # Connect the clicked signal to the run_sfc_scannow function
            self.winstart.clicked.connect(self.run_sfc_scannow)

    #SFC /Scannow && Dism /RestoreHealth script (no shutdown)
    def run_sfc_scannow(self):
        ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", " /C sfc /scannow && DISM /Online /Cleanup-Image /Restorehealth ", None, 1)

    #SFC /Scannow && Dism /RestoreHealth script WITH SHUTDOWN
    def run_sfc_scannow_shutdown(self):
        ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", " /C sfc /scannow && DISM /Online /Cleanup-Image /Restorehealth && shutdown /f /r /t 0 ", None, 1)
        
# Confirmation Window for Headset repairs
class headsetconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(400, 250)
        self.setWindowTitle("Meijer Headset Repair")

        # Start Button
        self.startbtn = QPushButton(self)
        self.startbtn.setGeometry(150, 115, 100, 100)
        self.startbtn.setText('Start')
        self.startbtn.setStyleSheet('font-size:30px')

        # Start Button Event
        self.startbtn.clicked.connect(self.script1)
        
        # Save Your Work Label
        self.savelbl = QLabel(self)
        self.savelbl.setGeometry(15, 0, 450, 55)
        self.savelbl.setText('Save any work, and continue when ready')
        self.savelbl.setStyleSheet('font-size:20px')

        #Label warning user of shutdown/restart
        self.restartlbl = QLabel(self)
        self.restartlbl.setGeometry(27, 35, 350, 55)
        self.restartlbl.setText('(Your Device will restart when you select "Start")')
        self.restartlbl.setStyleSheet('font-size:16px')

    #Driver Refresh Script Variable
    def script1(self):
        ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", "/C pnputil /delete-driver usbxhci.inf /uninstall /force /reboot", None, 1)
        

#Confirmation window for Manually installed Applications
class manconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(400, 300)
        self.setWindowTitle("Manually install Apps")

        #Power BI Button for app links
        self.powerbibtn = QPushButton(self)
        self.powerbibtn.setGeometry(15, 15, 150, 35)
        self.powerbibtn.setText('Power Bi Desktop')
        self.powerbibtn.setStyleSheet('font-size:12px')
        # PowerBI Button Event
        self.powerbibtn.clicked.connect(self.powerbidownload)

        #ZoOM button for zoom Link
        self.zoombtn = QPushButton(self)
        self.zoombtn.setGeometry(15, 60, 150, 35)
        self.zoombtn.setText('ZooM')
        self.zoombtn.setStyleSheet('font-size:12px')
        # ZooM Button Event
        self.zoombtn.clicked.connect(self.zoomdownload)

        #Button for chrome Link
        self.chromebtn = QPushButton(self)
        self.chromebtn.setGeometry(15, 105, 150, 35)
        self.chromebtn.setText('Chrome')
        self.chromebtn.setStyleSheet('font-size:12px')
        # Chrome Button Event
        self.chromebtn.clicked.connect(self.chromedownload)

        #Button for Google Earth Pro link
        self.gepbtn = QPushButton(self)
        self.gepbtn.setGeometry(15, 150, 150, 35)
        self.gepbtn.setText('Google Earth Pro')
        self.gepbtn.setStyleSheet('font-size:12px')
        # Google Earth Pro Button Event
        self.gepbtn.clicked.connect(self.googleearthprodownload)

    #Zoom process to download zoom w URL (purpose of "import Browser")
    def zoomdownload(self):
        url = "https://zoom.us/client/5.12.8.10232/ZoomInstallerFull.exe?archType=x64"
        self.zdown = QProcess()
        self.zdown.startDetached("cmd.exe", ["/c", "start", "", url])

    #PowerBI download Process w URL
    def powerbidownload(self):
        url = "https://download.microsoft.com/download/8/8/0/880BCA75-79DD-466A-927D-1ABF1F5454B0/PBIDesktopSetup_x64.exe"
        self.bidown = QProcess()
        self.bidown.startDetached("cmd.exe", ["/c", "start", "", url])
    
    def chromedownload(self):
        url = "https://www.google.com/chrome/thank-you.html?statcb=0&installdataindex=empty&defaultbrowser=0#"
        self.chromedown = QProcess()
        self.chromedown.startDetached("cmd.exe", ["/c", "start", "", url])

    #Google Earth Pro (gep) Download
    def googleearthprodownload(self):
        url = "https://www.google.com/earth/versions/download-thank-you/?usagestats=0"
        self.gepdown = QProcess()
        self.gepdown.startDetached("cmd.exe", ["/c", "start", "", url])

class driverconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(400, 300)
        self.setWindowTitle("Drivers for tings :)")

    #Nvidia T1000 Button for aDriver download
        self.t100btn = QPushButton(self)
        self.t100btn.setGeometry(15, 15, 150, 35)
        self.t100btn.setText('Nvidia T1000 Driver')
        self.t100btn.setStyleSheet('font-size:12px')
        # T1000 Button Event
        self.t100btn.clicked.connect(self.t1000download)

    #Nvidia T1000 Driver download process 
    def t1000download(self):
        url = "https://us.download.nvidia.com/Windows/Quadro_Certified/528.49/528.49-quadro-rtx-desktop-notebook-win10-win11-64bit-international-dch-whql.exe"
        self.t1000down = QProcess()
        self.t1000down.startDetached("cmd.exe", ["/c", "start", "", url])

#BIOSSledgehammer start window
class biossledge(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(380, 300)
        self.setWindowTitle("Update my BIOS!!!")
        
        #label during countdown
        self.bioslbl = QLabel(self)
        self.bioslbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.bioslbl.setGeometry(5, 30, 370, 55)
        self.bioslbl.setAlignment(Qt.AlignCenter)

        # Set the font of the label to a larger, bold style
        font = QFont("Arial", 10,)
        self.bioslbl.setFont(font)
        self.bioslbl.setText('BIOS update will begin momentarily...')
        self.bioslbl.setWordWrap(True)

        #consent checkbox that device will NOT be turned off
        self.bioscheckbox = QCheckBox(self)
        self.bioscheckbox.setGeometry(15, 225, 500, 55)
        self.bioscheckbox.setText('I will NOT turn off whilst BIOS updates. (MUST be checked)')
        self.bioscheckbox.setStyleSheet('font-size:12px')
        self.bioscheckbox.setChecked(False)
        self.bioscheckbox.stateChanged.connect(self.start_countdown)
        
        # Set the countdown timer to 5 seconds
        self.countdown_timer = QTimer(self)
        self.countdown_timer.setInterval(1000)  # 1 second interval
        self.countdown_timer.timeout.connect(self.update_countdown)
        self.countdown_value = 5

        # Set up the UI
        self.label = QLabel(str(self.countdown_value), self)
        self.label.setAlignment(Qt.AlignCenter)

        # Set the font of the label
        font = QFont("Open Sans", 20)
        self.label.setFont(font)
        self.label.setGeometry(175, 120, 40, 40)

    def start_countdown(self):
        if self.bioscheckbox.isChecked():
            self.countdown_timer.start()
        else:
            self.countdown_timer.stop()
            
    def update_countdown(self):
         #Decrement the countdown value
        self.countdown_value -= 1

         #Update the label with the new countdown value
        self.label.setText(str(self.countdown_value))

         #If the countdown value is zero, stop the timer and run the program
        if self.countdown_value == 0:
            self.countdown_timer.stop()
             #Run your program here (maybe)
            self.is_admin()
        else:
            return 0

    #Function to verify device is using Windows and has 64bit Powershell, followed by running BiosSledgehammer.ps1 as admin
    def is_admin(self):
        os.environ["PS_PART_PATH"] = "WindowsPowerShell\v1.0\powershell.exe" 
        os.environ["PS_EXE"] = "C:\Windows\System32\%PS_PART_PATH%" 
        os.environ["PS_EXE_SYSNATIVE"] = "c:\windows\sysnative\%PS_PART_PATH%"

        if os.path.exists(os.environ["PS_EXE_SYSNATIVE"]):
            os.environ["PS_EXE"] = os.environ["PS_EXE_SYSNATIVE"]

        #File Path Variables        
        cwd = os.getcwd()
        file_path = os.path.join(cwd, 'BiosSledgehammer', 'BiosSledgehammer.ps1')

        #if/else to request admin and run BiosSledge
        if ctypes.windll.shell32.IsUserAnAdmin() == 0:
            # Request admin privilege
            ctypes.windll.shell32.ShellExecuteW(None, "runas", "powershell.exe", f"-ExecutionPolicy Bypass -File {file_path} -WaitAtEnd", None, 1)
        else:
            # Run script as admin
            subprocess.run(['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', file_path, '-WaitAtEnd'], shell=True)

#Stores stuff
class projectsconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(400, 300)
        self.setWindowTitle("Store stuffs (work in progess) :)")


#macOS stuff
class macconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(400, 300)
        self.setWindowTitle("work in progress :)")


#Sweet stuff :>
class suiteconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(400, 300)
        self.setWindowTitle("Fingers Crossed :)")

        #Office Repair/update Button
        self.officeupdatebtn = QPushButton(self)
        self.officeupdatebtn.setGeometry(15, 15, 150, 35)
        self.officeupdatebtn.setText('Office Update')
        self.officeupdatebtn.setStyleSheet('font-size:12px')

        # Office Repair button event
        self.officeupdatebtn.clicked.connect(self.officeupdatescript)

         #Office Repair/update script
    def officeupdatescript(self):
        ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", "/C cd C:\Program Files\Common Files\microsoft shared\ClickToRun && OfficeC2RClient.exe /update user ", None, 1)

window = Ui()
sys.exit(app.exec_())