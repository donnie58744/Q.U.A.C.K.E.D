import shutil
import pystray,sys,os, json, re, getpass, win32com.client, requests
from PIL import Image
from PyQt6 import uic, QtTest
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt6.QtCore import Qt, QThread, pyqtSignal, pyqtSlot, QObject, QProcess
from PyQt6.QtGui import QIcon
from time import sleep
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
from discord import Webhook, SyncWebhook,File
from libs.CalculateBitrate import calcBitrate
import signal

dir_path = os.path.dirname(os.path.realpath(__file__))

class main():
    username = ''
    avatar = ''
    listenerThreadRunning = True
    filename=''

    def getConfig(self, key):
        with open(dir_path + '/files/config.json') as f:
            data = json.load(f)

        return data[key]
        
    def writeConfig(self, key, value):
        with open(dir_path + '/files/config.json') as f:
            data = json.load(f)
            if key in data:
                del data[key]
                cacheDict = dict(data)
                cacheDict.update({key:value})
                with open(dir_path + '/files/config.json', 'w') as f:
                    json.dump(cacheDict, f, indent=4)
    
    def stringToBool(self, string):
        match(str(string).lower()):
            case 'true':
                return True
            case 'false':
                return False

    def deleteWindowsStartupShortcut(self):
        USER_NAME = getpass.getuser()
        startupFolder = fr"C:\Users\{USER_NAME}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        os.remove(startupFolder+'/Q.U.A.C.K.E.D.lnk')

    def createWindowsStartupShortcut(self):
        USER_NAME = getpass.getuser()
        dir_path = os.path.dirname(os.path.realpath(__file__))
        print(dir_path)

        # pythoncom.CoInitialize() # remove the '#' at the beginning of the line if running in a thread.
        startupFolder = fr"C:\Users\{USER_NAME}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" # path to where you want to put the .lnk
        path = os.path.join(startupFolder, 'Q.U.A.C.K.E.D.lnk')
        target = dir_path+'/Q.U.A.C.K.E.D.exe'

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WindowStyle = 7 # 7 - Minimized, 3 - Maximized, 1 - Normal
        shortcut.save()

    def getDiscordUserInfo(self, discordId):
        try:
            headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0',
            }
            r = requests.get(f'https://discordlookup.mesavirep.xyz/v1/user/{discordId}', headers=headers)
            x = json.loads(r.text)
            id = str(x["id"])
            username = str(x["tag"]).split('#',1)[0]
            icon = x["avatar"]['link']
            return [id,username,icon]
        except Exception as e:
            pass

    def previewClip(self, clip):
        os.startfile(clip)

class listener(QObject):
    def __init__(self, signal_to_emit):
        super().__init__()
        self.signal_to_emit = signal_to_emit
        self.p = None

    @pyqtSlot()
    def executeThread(self):
        try:
            patterns = ["*"]
            ignore_patterns = None
            ignore_directories = True
            case_sensitive = True
            my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)
            my_event_handler.on_created = self.on_created
            path = rf"{main().getConfig(key='capture_folder')}"
            go_recursively = True
            self.my_observer = Observer()
            self.my_observer.schedule(my_event_handler, path, recursive=go_recursively)
            self.my_observer.start()
            while main.listenerThreadRunning:
                print('Thread Running')
                sleep(1)
            else:
                self.my_observer.stop()
        except Exception as e:
            self.my_observer.stop()
            try:
                self.my_observer.join()
            except Exception as e:
                pass
            pass
        
    def on_created(self, event):
        self.my_observer.stop()
        if (main.listenerThreadRunning):
            main.listenerThreadRunning = False
            main.filename = str(event.src_path)
            self.signal_to_emit.emit('', '','mainScreen',str(event.src_path))

class HandbrakeCLI(QObject):
    def __init__(self, signal_to_emit):
        super().__init__()
        self.signal_to_emit = signal_to_emit
        self.p = None
        self.lock = False

    @pyqtSlot()
    def executeThread( self ):
        print(main.filename.split('.',1)[0])
        newFileName = main.filename.split('.',1)[0] + '_Compressed.mp4'
        
        handbrakeCommand = ["-i", rf'{main.filename}', "-o", rf'{newFileName}', "--vb", str(calcBitrate().solveEquation(filename=main.filename))]
        print(handbrakeCommand)
        self.signal_to_emit.emit('', '','compressionScreen','')
        main.filename = newFileName
        if self.p is None:  # No process running.
            self.signal_to_emit.emit('console', 'Beep Boop, Handbrake Starting...','updateGui','')
            self.p = QProcess()  # Keep a reference to the QProcess (e.g. on self) while it's running.
            self.pid = self.p.processId()
            self.p.finished.connect(self.process_finished)  # Clean up once complete.
            customCommand = str(main().getConfig(key='handbrakeCommand')).split(',')
            print(handbrakeCommand+customCommand)
            self.p.start(dir_path+"/files/HandBrakeCLI.exe",handbrakeCommand+customCommand)
            self.p.readyReadStandardOutput.connect(self.handle_stdout)
            
    def handle_stdout(self):
        data = self.p.readAllStandardOutput()
        pattern = re.compile(r"(\d+(\.\d+) ?%)", re.IGNORECASE)
        stdout = bytes(data).decode("utf8")
        # regex match for % complete and ETA
        matches = pattern.findall(stdout)

        if matches:
            if (float(str(matches[0][0]).replace('%','').replace(' ','')) <=99 and self.lock == False):
                task = 'Encoding Task 1 of 2: '
            else :
                self.lock = True
                task = 'Encoding Task 2 of 2: '
            
            print(float(str(matches[0][0]).replace('%','').replace(' ','')))
            self.signal_to_emit.emit('console',  task+str(matches[0][0]),'updateGui','')

    def quit(self):
        self.p = None
        self.lock=False
        try:
            os.kill(self.pid, signal.CTRL_C_EVENT) #SIGINT is CTRL-
        except:
            pass

    def process_finished(self):
        self.p.close()
        self.p = None
        self.lock=False
        self.signal_to_emit.emit('', '','shareScreen','')

class TrayIcon(QObject):
    def __init__(self, signal_to_emit):
        super().__init__()
        self.signal_to_emit = signal_to_emit

    def create_image(self):
        # Generate an image and draw a pattern
        image = Image.open(dir_path+'/res/duckbutton.png')
        return image

    @pyqtSlot()
    def executeThread(self):
        # In order for the icon to be displayed, you must provide an icon
        self.icon = pystray.Icon(
            name='Q.U.A.C.K.E.D',
            icon=self.create_image(),title='Q.U.A.C.K.E.D')
        self.icon.menu=pystray.Menu(
            pystray.MenuItem("Manual Mode", self.manualMode),
            pystray.MenuItem("Settings", self.settingsScreen),
            pystray.MenuItem("Quit", self.quitProgram)
        )

        try:
            # To finally show you icon, call run
            self.icon.run_detached()
        except Exception as e:
            self.icon.stop()

    def quit(self):
        self.icon.icon=''
        self.icon.stop()

    def manualMode(self):
        self.quit()
        self.signal_to_emit.emit('', '','manualMode','')
    
    def settingsScreen(self):
        self.signal_to_emit.emit('', '','settingsScreen','')

    def quitProgram(self):
        self.quit()
        QtTest.QTest.qWait(700)
        self.signal_to_emit.emit('', '','close','')
        
class Ui(QMainWindow):
    mainThreadSig = pyqtSignal(str,str,str,str)

    def __init__(self):
        super().__init__()
        app.aboutToQuit.connect(self.shutdown)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setFixedSize(self.size())
        self.setWindowIcon(QIcon(dir_path+'/res/duckbutton.png'))
        # Center
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        
        # Start HandBrake Thread
        self.handbrakeThread = HandbrakeCLI(self.mainThreadSig)
        self.handbrakeFoo = QThread(self) 
        self.handbrakeThread.moveToThread(self.handbrakeFoo)
        
        # Start listner Thread
        self.listenerThread = listener(self.mainThreadSig)
        self.listenerFoo = QThread(self) 
        self.listenerThread.moveToThread(self.listenerFoo)

        # Start Tray Thread
        self.trayThread = TrayIcon(self.mainThreadSig)
        self.trayFoo = QThread(self) 
        self.trayThread.moveToThread(self.trayFoo)

        # Connect the thread to the functions
        self.listenerFoo.started.connect(self.listenerThread.executeThread)
        self.handbrakeFoo.started.connect(self.handbrakeThread.executeThread)
        self.trayFoo.started.connect(self.trayThread.executeThread)
        print(str(main().getConfig(key='handbrakeCommand')))
        # Load and set user info from config
        main.username = main().getConfig(key='username')
        main.avatar = main().getConfig(key='avatar_url')

        self.mainThreadSig.connect(self.threadReciver)
        self.startProgram()

    def manualMode(self, request):
        if (request == 'tray'):
            self.trayFoo.terminate()
            main.listenerThreadRunning = False
            self.listenerFoo.terminate()
        dlg = QFileDialog(self)
        dlg.setFileMode(QFileDialog.FileMode.ExistingFile)
        dlg.setDirectory(main().getConfig(key='capture_folder'))
        dlg.setNameFilter("All files (*.*);; MP4 File (*.mp4);; MKV File (*.mkv)")

        if (dlg.exec()):
            dlgFilename = dlg.selectedFiles()
            main.filename = dlgFilename[0]
            if (request == 'tray'):
                self.trayFoo.start()
            self.mainScreen(filename=dlgFilename[0], label='<html><head/><body><p><span style=" font-size:12pt; font-weight:600;">Manual Mode</span></p></body></html>')
        else:
            if (request == 'tray'):
                #Pop window over everything
                main.listenerThreadRunning = True
                self.listenerFoo.start()
                self.trayFoo.start()

    def shutdown(self):
        self.handbrakeThread.quit()
        self.handbrakeFoo.terminate()
        main.listenerThreadRunning = False
        self.listenerFoo.terminate()
        self.trayThread.quit()
        self.trayFoo.terminate()
        app.quit()
        os.system('taskkill /IM "' + "HandBrakeCLI.exe" + '" /F')
        os._exit(os.X_OK)

    @pyqtSlot(str,str,str,str)
    def threadReciver(self, label, text, request, filename):
        if (request == 'updateGui'):
            getattr(self,label).setText(text)
        elif (request == 'close'):
            self.shutdown()
        elif (request == 'mainScreen'):
            self.mainScreen(filename=filename)
        elif (request == 'compressionScreen'):
            self.compressionScreen()
        elif (request == 'shareScreen'):
            self.shareScreen()
        elif (request == 'manualMode'):
            self.manualMode(request='tray')
        elif (request == 'settingsScreen'):
            self.settingsScreen()

        elif (request=='elementVisible'):
            getattr(self,label).setVisible(bool(text))
            
    def yesBtnClicked(self):
        self.yesBtn.setVisible(False)
        self.noBtn.setVisible(False)
        self.console.setVisible(True)
        main.listenerThreadRunning = False
        self.listenerFoo.terminate()
        # Start Compression
        self.handbrakeFoo.start()
    
    def noBtnClicked(self):
        main.filename = ''
        main.listenerThreadRunning = True
        self.listenerFoo.start()
        self.hide()

    def startProgram(self):
        self.trayFoo.start()
        main.listenerThreadRunning = True
        self.listenerFoo.start()
        uic.loadUi(dir_path+'/files/mainMenu.ui', self)
        self.versionLabel.setText(main().getConfig(key='version'))
        self.show()
        QtTest.QTest.qWait(3500)
        self.hide()

    def mainScreen(self,filename=None, label=None):
        self.listenerFoo.terminate()
        uic.loadUi(dir_path+'/files/main.ui', self)
        if(label != None):
            self.mainLabel.setText(label)
        self.console.setVisible(False)
        self.filenameLabel.setText(f'<html><head/><body><p><span style=" font-size:11pt;">{filename}</span></p></body></html>')
        self.yesBtn.clicked.connect(self.yesBtnClicked)
        self.noBtn.clicked.connect(self.noBtnClicked)
        self.previewRawClipBtn.clicked.connect(lambda: main().previewClip(clip=main.filename))
        self.differntClipBtn.clicked.connect(self.manualMode)
        #Pop window over everything
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.show()
        self.setWindowFlags(self.windowFlags() & ~ Qt.WindowType.WindowStaysOnTopHint)
        self.show()
        self.activateWindow()

        # Check if the yes btn is enabled
        if (not main().stringToBool(string=main().getConfig(key='yesBtnEnabled'))):
            self.yesBtnClicked()

    def compressionScreen(self):
        self.mainLabel.setText('<html><head/><body><p><span style=" font-size:12pt; font-weight:600;">Compressing...</span></p></body></html>')
        self.console.setVisible(True)
        self.previewRawClipBtn.setVisible(False)
        self.differntClipBtn.setVisible(False)

    def setupSettingsScreenBtns(self):
        self.closeBtn.clicked.connect(self.hide)
        self.generalBtn.clicked.connect(self.generalSettingsScreen)
        self.sharingOptionsBtn.clicked.connect(self.sharingSettingsScreen)
        self.aboutBtn.clicked.connect(self.settingsScreen)

    def settingsScreen(self):
        uic.loadUi(dir_path+'/files/settingsScreen.ui', self)
        self.versionLabel.setText(f'<html><head/><body><p><span style=" font-size:10pt;">Version: {main().getConfig(key="version")}</span></p></body></html>')
        self.setupSettingsScreenBtns()
        self.show()
        self.activateWindow()

    def applyGeneralSettings(self):
        main().writeConfig(key='capture_folder', value=self.captureFolderTxtBox.text())
        main().writeConfig(key='yesBtnEnabled', value=str(self.yesBtnEnabledCheckBox.isChecked()))
        main().writeConfig(key='handbrakeCommand', value=self.handbrakeSettingTxtBox.text())

        # Check if windows startup has been enabled
        if (main().getConfig(key='windowsStartup') != str(self.windowsStartupCheckBox.isChecked()) and self.windowsStartupCheckBox.isChecked()):
            main().createWindowsStartupShortcut()
        elif (main().getConfig(key='windowsStartup') != str(self.windowsStartupCheckBox.isChecked()) and not self.windowsStartupCheckBox.isChecked()):
            main().deleteWindowsStartupShortcut()

        main().writeConfig(key='windowsStartup', value=str(self.windowsStartupCheckBox.isChecked()))

    def generalSettingsScreen(self):
        uic.loadUi(dir_path+'/files/generalSettingsScreen.ui', self)
        self.setupSettingsScreenBtns()
        self.pickFileDirBtn.clicked.connect(self.pickFile)
        self.captureFolderTxtBox.setText(main().getConfig(key='capture_folder'))
        self.applyBtn.clicked.connect(self.applyGeneralSettings)
        self.cancelBtn.clicked.connect(self.settingsScreen)
        self.yesBtnEnabledCheckBox.setChecked(main().stringToBool(string=main().getConfig(key='yesBtnEnabled')))
        self.windowsStartupCheckBox.setChecked(main().stringToBool(string=main().getConfig(key='windowsStartup')))
        self.handbrakeSettingTxtBox.setText(main().getConfig(key='handbrakeCommand'))

    def pickFile(self):
        dlg = QFileDialog(self)
        dlg.setFileMode(QFileDialog.FileMode.Directory)
        if (main().getConfig(key='capture_folder') == ''):
            USER_NAME = getpass.getuser()
            dlg.setDirectory(fr"C:\Users\{USER_NAME}\Videos\Captures")
        else:
            dlg.setDirectory(main().getConfig(key='capture_folder'))
        dlg.setNameFilter("All files (*.*);; MP4 File (*.mp4);; MKV File (*.mkv)")

        if (dlg.exec()):
            dlgFilename = dlg.selectedFiles()
            filepath = dlgFilename[0]
            self.captureFolderTxtBox.setText(str(filepath))
            return filepath

    def autoFindBtnClicked(self):
        discordUserInfo = main().getDiscordUserInfo(discordId=self.discordIDTxtBox.text())
        self.displayNameTxtBox.setText(discordUserInfo[1])
        self.avatarUrlTxtBox.setText(discordUserInfo[2])

    def sharingSettingsScreen(self):
        uic.loadUi(dir_path+'/files/sharingSettingsScreen.ui', self)
        self.setupSettingsScreenBtns()
        self.webhookUrlTxtBox.setText(main().getConfig(key='discord_webhook_url'))
        self.displayNameTxtBox.setText(main().getConfig(key='username'))
        self.avatarUrlTxtBox.setText(main().getConfig(key='avatar_url'))
        self.applyBtn.clicked.connect(lambda: main().writeConfig(key='username', value=self.displayNameTxtBox.text()))
        self.applyBtn.clicked.connect(lambda: main().writeConfig(key='avatar_url', value=self.avatarUrlTxtBox.text()))
        self.applyBtn.clicked.connect(lambda: main().writeConfig(key='discord_webhook_url', value=self.webhookUrlTxtBox.text()))
        self.cancelBtn.clicked.connect(self.settingsScreen)
        self.autoFindBtn.clicked.connect(self.autoFindBtnClicked)

    def shareScreen(self):
        self.handbrakeFoo.terminate()
        uic.loadUi(dir_path+'/files/shareScreen.ui', self)
        self.discordBtn.clicked.connect(self.discordBtnClicked)
        self.saveLaterBtn.clicked.connect(self.saveLaterBtnClicked)
        self.previewCompressedClipBtn.clicked.connect(lambda: main().previewClip(clip=main.filename))
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.show()
        self.setWindowFlags(self.windowFlags() & ~ Qt.WindowType.WindowStaysOnTopHint)
        self.show()
        self.activateWindow()

    # Have users make an account  - send info to db - get users code - whitelist them | When they login just verify the username and password

    def discordBtnClicked(self):
        uic.loadUi(dir_path+'/files/discordScreen.ui', self)
        self.sendBtn.clicked.connect(self.discordSendBtnClicked)
        self.msgTxtBox.setFocus()
        self.msgTxtBox.returnPressed.connect(self.discordSendBtnClicked)

    def discordSendBtnClicked(self):
        self.mainLabel.setText('<html><head/><body><p><span style=" font-size:12pt; font-weight:600;">Please Wait...</span></p></body></html>')
        self.sendBtn.setVisible(False)
        self.msgTxtBox.setVisible(False)
        QtTest.QTest.qWait(500)
        try:
            discord_webhook_url=main().getConfig(key='discord_webhook_url')
            webhook = SyncWebhook.from_url(discord_webhook_url)
            webhook.send(str(self.msgTxtBox.text())+"\n"+"\n"+"`Fully Compressed And Sent With Q.U.A.C.K.E.D`", file=File(rf'{main.filename}'), username=main().getConfig(key='username'), avatar_url=main().getConfig(key='avatar_url'))
            self.mainLabel.setText('<html><head/><body><p><span style=" font-size:12pt; font-weight:600;">Sent To Discord!</span></p><p><span style=" font-size:11pt;">This Window Will Automatically Close</span></p></body></html>')
            QtTest.QTest.qWait(5000)
            main.listenerThreadRunning = True
            self.listenerFoo.start()
            self.hide()
        except Exception as e:
            self.mainLabel.setText(f'<html><head/><body><p><span style=" font-size:12pt; font-weight:600;">Error!</span></p><p><span style=" font-size:11pt;">{str(e)}</span></p><p><span style=" font-size:11pt;">Going Back To Share Screen.</span></p></body></html>')
            print(e)
            QtTest.QTest.qWait(5000)
            self.shareScreen()

    def saveLaterBtnClicked(self):
        self.mainLabel.setText('<html><head/><body><p><span style=" font-size:12pt; font-weight:600;">Saved!</span></p><p><span style=" font-size:11pt;">This Window Will Automatically Close</span></p></body></html>')
        QtTest.QTest.qWait(5000)
        main.listenerThreadRunning = True
        self.listenerFoo.start()
        self.hide()

app = QApplication(sys.argv)
w = Ui()
app.exec()