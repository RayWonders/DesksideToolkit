import sys
import webbrowser
import subprocess
import ctypes
import os
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *

app = QtWidgets.QApplication(sys.argv)

class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('DesksideToolkit.ui', self)
        self.show()
        self.setWindowTitle("Robb's Deskside Toolkit")

    #Slot signal for Headset repairs
    def headsetconfwin(self):
        self.w = headsetconfwin()
        self.w.show()
        self.hide()

    #slot window for Windows repairs
    def windowsconfwin(self):
        self.wcw = windowsconfwin()
        self.wcw.show()
        self.hide()

    #Slot for Manually installed apps
    def manconfwin(self):
        self.mcw = manconfwin()
        self.mcw.show()
        self.hide()

    #Slot for BIOSSledgehammer start window
    def biossledge(self):
        self.bsh = biossledge()
        self.bsh.show()
        self.hide

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
        #if self.p is None: No process Running
        self.p = QProcess()
        self.p.finished.connect(self.script1_finished) #clean up once complete.
        self.p.start ("Headsetrepair.bat")

    def script1_finished(self):
        self.p = None

#Confirmation window for Manually installed Applications
class manconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(400, 300)
        self.setWindowTitle("Manually install Apps")

        #Test Button for app links
        self.powerbibtn = QPushButton(self)
        self.powerbibtn.setGeometry(15, 15, 150, 35)
        self.powerbibtn.setText('Power Bi Desktop')
        self.powerbibtn.setStyleSheet('font-size:12px')
        # PowerBI Button Event
        self.powerbibtn.clicked.connect(self.powerbidownload)

        #Test button for zoom Link
        self.zoombtn = QPushButton(self)
        self.zoombtn.setGeometry(15, 60, 150, 35)
        self.zoombtn.setText('ZooM')
        self.zoombtn.setStyleSheet('font-size:12px')
        # ZooM Button Event
        self.zoombtn.clicked.connect(self.zoomdownload)

    #Zoom process to download zoom w URL (purpose of "import Browser")
    def zoomdownload(self):
        #if self.p is None: No process Running
        self.zdown = QProcess()
        self.zdown.start (webbrowser.open_new_tab("https://zoom.us/client/5.12.8.10232/ZoomInstallerFull.exe?archType=x64"))

    def zoomdownload_finished(self):
        self.zdown = None

    #PowerBI download Process w URL
    def powerbidownload(self):
        #if self.p is None: No process Running
        self.bidown = QProcess()
        self.bidown.start (webbrowser.open_new_tab("https://download.microsoft.com/download/8/8/0/880BCA75-79DD-466A-927D-1ABF1F5454B0/PBIDesktopSetup_x64.exe"))

    def powerbidownload_finished(self):
        self.bidown = None

#BIOSSledgehammer start window
class biossledge(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(380, 200)
        self.setWindowTitle("Update my BIOS!!!")

        #label during countdown
        self.bioslbl = QLabel(self)
        self.bioslbl.setGeometry(5, 0, 370, 55)
        self.bioslbl.setAlignment(Qt.AlignCenter)

        # Set the font of the label to a larger, bold style
        font = QFont("Arial", 12,)
        self.bioslbl.setFont(font)
        self.bioslbl.setText('BIOS update will begin momentarily...')

        #consent checkbox that device will NOT be turned off
        self.bioscheckbox = QCheckBox(self)
        self.bioscheckbox.setGeometry(10, 135, 500, 55)
        self.bioscheckbox.setText('I will NOT turn off whilst BIOS updates. (MUST be checked)')
        self.bioscheckbox.setStyleSheet('font-size:12px')
        self.bioscheckbox.setChecked(False)
        #self.bioscheckbox.stateChanged.connect(self.run_bs)
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
        font = QFont("Open Sans", 24)
        self.label.setFont(font)
        self.label.setGeometry(168, 75, 40, 40)

    def start_countdown(self, state):
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
             #Run yourdww program here (maybe)
            self.is_admin()
        else:
            return 0

    def is_admin(self):
        ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", "cmd /k " + os.getcwd() + "\DesksideToolkit\BiosSledgehammer\RunVisable.bat", None, 1)
        
"""
    def bs_script1(self):        
        self.bs_script = QProcess()
        self.bs_script.finished.connect(self.bs_script1_finished) #clean up once complete.
        self.bs_script.start("BiosSledgehammer/RunVisable.bat")

    def bs_script1_finished(self):
        self.bs_script = None
"""
window = Ui()
sys.exit(app.exec_())