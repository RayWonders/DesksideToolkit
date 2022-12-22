import sys
import webbrowser
import subprocess
import ctypes
from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5.QtCore import QProcess, Qt
from PyQt5.QtWidgets import QMainWindow, QLabel, QPushButton, QCheckBox

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

#window to activate DISM/SFC repair w checkbox
class windowsconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(380, 200)
        self.setWindowTitle("Windows 10 repair (Fingers Crossed)")

        #checkbox for Restart function
        self.winchk = QCheckBox(self)
        self.winchk.setGeometry(10, 140, 350, 55)
        self.winchk.setText('restart when completed? (recommended)')
        self.winchk.setStyleSheet('font-size:15px')
        self.winchk.setChecked(False)

        #"Start" Button for SFC/DISM
        self.winstart = QPushButton(self)
        self.winstart.setGeometry(115, 25, 160, 100)
        self.winstart.setText('Start')
        self.winstart.setStyleSheet('font-size:25px')
        self.winstart.clicked.connect(self.winscript)

        self.winchk.stateChanged.connect(self.chkboxchange)

    def chkboxchange(self):
        if self.winchk.isChecked():
            self.winstart.clicked.disconnect() # Disconnect the clicked signal from all its connections
            # Connect the clicked signal to the winscript2 function
            self.winstart.clicked.connect(self.run_sfc_scannow)
        else:
            self.winstart.clicked.disconnect() # Disconnect the clicked signal from all its connections
            # Connect the clicked signal to the winscript function
            self.winstart.clicked.connect(self.winscript)

    #SFC /Scannow && Dism /RestoreHealth script
    def winscript(self):
        self.winfix = QProcess()
        self.winfix.finished == None
        self.winfix.start('adminask.bat', ['sfc /scannow && DISM /Online /Cleanup-Image /Restorehealth'])

    #SFC /Scannow && Dism /RestoreHealth script WITH RESTARTS
    def run_sfc_scannow(self):
        def is_admin():
            try:
                return ctypes.windll.shell32.IsUserAnAdmin()
            except:
                return False

        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", " /C sfc /scannow && DISM /Online /Cleanup-Image /Restorehealth && shutdown /f /r /t 0 ", None, 1)
        else:
            subprocess.call(["cmd.exe", "/C", "sfc", "/scannow"])

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

#Confirmation window for Manially installed Applications
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

window = Ui()
sys.exit(app.exec_())