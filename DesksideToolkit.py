import sys
import webbrowser
from PyQt5 import QtWidgets, uic
from PyQt5 import QtWidgets
from PyQt5.QtCore import QProcess 
from PyQt5.QtWidgets import QMainWindow, QLabel, QPushButton, QCheckBox

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

class windowsconfwin(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(380, 200)
        self.setWindowTitle("Windows 10 repair (Fingers Crossed)")

        self.winchk = QCheckBox(self)
        self.winchk.setGeometry(55, 140, 450, 55)
        self.winchk.setText('restart when completed? (recommended)')
        self.winchk.setStyleSheet('font-size:15px')

        self.winstart = QPushButton(self)
        self.winstart.setGeometry(115, 25, 160, 100)
        self.winstart.setText('Start')
        self.winstart.setStyleSheet('font-size:30px')

        if self.winchk.isChecked() == True:
                self.winstart.clicked.connect(self.winscript2)
        else:
                self.winstart.clicked.connect(self.winscript)

        #SFC /Scannow && Dism /RestoreHealth script
    def winscript(self):
        self.winp = QProcess()
        self.winp.finished.connect(self.winscript_finished)
        self.winp.start ("winrepair.bat")

    def winscript_finished(self):
        self.winp = None

        #SFC /Scannow && Dism /RestoreHealth script with Restart

    #def winscript2(self):
        #self.winpwr = QProcess()
        #self.winpwr.finished.connect(self.winscript2_finished)
        #self.winpwr.start ("winrepairrestart.bat")

    def winscript2_finished(self):
        self.winpwr = None

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

            #Driver Refresh Script Variable
    def zoomdownload(self):
        #if self.p is None: No process Running
        self.zdown = QProcess()
        self.zdown.start (webbrowser.open_new_tab("https://zoom.us/client/5.12.8.10232/ZoomInstallerFull.exe?archType=x64"))

    def zoomdownload_finished(self):
        self.zdown = None

    def powerbidownload(self):
        #if self.p is None: No process Running
        self.bidown = QProcess()
        self.bidown.start (webbrowser.open_new_tab("https://download.microsoft.com/download/8/8/0/880BCA75-79DD-466A-927D-1ABF1F5454B0/PBIDesktopSetup_x64.exe"))

    def powerbidownload_finished(self):
        self.bidown = None

app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()