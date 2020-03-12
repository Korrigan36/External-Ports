# -*- coding: utf-8 -*-
"""
Created on Wed Aug 09 12:28:54 2017

@author: v-stpurc

Based on LUA code written by David Fuhriman v-davifu 09/09/2015
Copyright 2017 Microsoft - All rights reserved
"""
import sys
import visa
from time import sleep
from datetime import datetime

from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QRadioButton, QVBoxLayout, QCheckBox, QProgressBar,
    QGroupBox, QComboBox, QLineEdit, QPushButton, QMessageBox, QInputDialog, QDialog, QDialogButtonBox, QSlider, QFileDialog)
from PyQt5.QtGui import QIcon, QPainter, QPen, QFont, QPixmap
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QCoreApplication, QObject, QRunnable, QThread, QThreadPool, pyqtSignal, pyqtSlot

sys.path.append(r"\\gameshare\ieb\yukon\Toledo\VRegs\Zach_stuff\instrument_Libraries")

#import tekScopeLib
import misc_Functions
import dutInfoDialog
import timerThread
#Global Variables
#TestSetup_Defaults_1 = {'Port': '','Channel': 5,'Amps': 5,'Duration': 5,'Trigger': 1.0} 
#TestSetup_Defaults_1['Port'] = "USB3 5V Rear Top"
#print TestSetup_Defaults_1['Port']


class MainWindow(QWidget):
    
    is_okay = None
    testList = {}
    dutList = {}
    step_Count = None
    str_load_frame_id = None

    def __init__(self):
        super(MainWindow, self).__init__()
        
        self.initUI()
        self.openInstruments()
        self.openFileNameDialog()
        self.getDutInfo()
        self.startThread()

    def quitLoop(self):
        
        self.thread.quit()
        print("Got quit signal")

    def runLoop(self, stepCount):
        self.progressBar.setValue(int(stepCount))

    def testInfo(self, testType, testID1, testID2, testID3, portName1, portName2, portName3, amps1
                 , amps2, amps3, duration1, duration2, duration3, trigger1, trigger2, trigger3, pulseT1, pulseT2, pulseT3):
        self.testTypeLineEdit.setText(testType)
        self.testID1LineEdit.setText(testID1)
        self.testID2LineEdit.setText(testID2)
        self.testID3LineEdit.setText(testID3)
        self.portName1LineEdit.setText(portName1)
        self.portName2LineEdit.setText(portName2)
        self.portName3LineEdit.setText(portName3)
        self.amps1LineEdit.setText(amps1)
        self.amps2LineEdit.setText(amps2)
        self.amps3LineEdit.setText(amps3)
        self.duration1LineEdit.setText(duration1)
        self.duration2LineEdit.setText(duration2)
        self.duration3LineEdit.setText(duration3)
        self.trigger1LineEdit.setText(trigger1)
        self.trigger2LineEdit.setText(trigger2)
        self.trigger3LineEdit.setText(trigger3)
        self.pulseT1LineEdit.setText(pulseT1)
        self.pulseT2LineEdit.setText(pulseT2)
        self.pulseT3LineEdit.setText(pulseT3)

    def startThread(self):
        #We're only using one Chroma mainframe here. Config file functionality not fully implemented
        self.work = timerThread.TimerThread(self.load, self.str_load_frame_id, self.step_Count, self.dutList, self.testList) 
        self.work.timerSignal.connect(self.runLoop)
        self.work.quitSignal.connect(self.quitLoop)
        self.work.testInfoSignal.connect(self.testInfo)
        self.thread = QThread()

        self.work.moveToThread(self.thread)
        self.thread.started.connect(self.work.run)
        self.thread.finished.connect(self.close)
        self.thread.start()

    def getDutInfo(self):
        is_okay = True 
        self.dutList[0] = {} 
        self.dutList[0]['run_count'] = 1 
        
        self.step_Count = 0
        console_Count = 0
        #Go through each item in testList and check console number
        #While at it increment step count. This is the number of lines on the test sequence tab
        for key in self.testList: # scan test sequence to find how many consoles there are 
            self.step_Count = self.step_Count + 1
            if key != 0:
                console_num = self.testList[key]['Console']
                if console_num not in self.dutList: 
                    console_Count = console_Count + 1
                    self.dutList[console_num] = {} # instantiate a table for this console_num 

        #Check for sequential dutList entries was here
		
        for loop in range (1, console_Count + 1): 
            dialog = dutInfoDialog.DutInfoDialog()
            dialog.setWindowModality(Qt.ApplicationModal)
            dialog.exec_()
            self.configID, self.pcbaSN, self.productSN, self.runNotes, = dialog.returnInfo()
    			
            self.dutList[loop]['dut_config'] = self.configID
            self.dutList[loop]['dut_pcbaSn'] = self.pcbaSN 
            self.dutList[loop]['dut_productSn'] = self.productSN 
            self.dutList[loop]['dut_tag'] = "Unit_" + str(loop)
        
        #The following comment looks out of place in new object oriented implementation
        #Took out big loop repeat feature. Just once through is good for now.
       
        self.progressBar.setMaximum(self.step_Count)


    def initUI(self):
        
        self.setGeometry(300, 300, 1000, 300)
        self.setWindowTitle('XBox Power Ports Test')
        self.setWindowIcon(QIcon('xbox_icon.ico')) 

        loadLabel = QLabel(self)
        loadLabel.move(10,10)
        loadLabel.setText("Chroma Load")
  
        self.loadIDN = QLabel(self)
        self.loadIDN.setGeometry(150, 10, 1000, 10)
        self.loadIDN.setText("Chroma Load")

        self.progressBar = QProgressBar(self)
#        self.progressBar.setGeometry(500, 40, 500, 10)
        self.progressBar.setGeometry(500, 30, 500, 10)
        
        self.progressLabel = QLabel(self)
        self.progressLabel.setGeometry(700, 10, 100, 10)
        self.progressLabel.setText("Test Progress")

        testTypeLabel = QLabel(self)
        testTypeLabel.setGeometry(10, 50, 95, 20)
        testTypeLabel.setText("Test File Name")
        
        self.testTypeLineEdit = QLabel(self)
        self.testTypeLineEdit.setGeometry(10, 70, 500, 20)
        self.testTypeLineEdit.setText("")

        self.font = QFont()
        self.font.setBold(True)
        self.font.setPointSize(12)
        self.testTypeLineEdit.setFont(self.font)

        testID1 = QLabel(self)
        testID1.setGeometry(10, 90, 1000, 20)
        testID1.setText("Test ID  Port Name      Load     Duration   Trigger   Pulse Current")
        
        self.testID1LineEdit = QLabel(self)
        self.testID1LineEdit.setGeometry(10, 105, 100, 20)
        self.testID1LineEdit.setText("")

        self.portName1LineEdit = QLabel(self)
        self.portName1LineEdit.setGeometry(50, 105, 100, 20)
        self.portName1LineEdit.setText("")

        self.amps1LineEdit = QLabel(self)
        self.amps1LineEdit.setGeometry(120, 105, 100, 20)
        self.amps1LineEdit.setText("")

        self.duration1LineEdit = QLabel(self)
        self.duration1LineEdit.setGeometry(155, 105, 100, 20)
        self.duration1LineEdit.setText("")

        self.trigger1LineEdit = QLabel(self)
        self.trigger1LineEdit.setGeometry(210, 105, 100, 20)
        self.trigger1LineEdit.setText("")

        self.pulseT1LineEdit = QLabel(self)
        self.pulseT1LineEdit.setGeometry(250, 105, 100, 20)
        self.pulseT1LineEdit.setText("")

        self.testID2LineEdit = QLabel(self)
        self.testID2LineEdit.setGeometry(10, 120, 100, 20)
        self.testID2LineEdit.setText("")

        self.portName2LineEdit = QLabel(self)
        self.portName2LineEdit.setGeometry(50, 120, 100, 20)
        self.portName2LineEdit.setText("")

        self.amps2LineEdit = QLabel(self)
        self.amps2LineEdit.setGeometry(120, 120, 100, 20)
        self.amps2LineEdit.setText("")

        self.duration2LineEdit = QLabel(self)
        self.duration2LineEdit.setGeometry(155, 120, 100, 20)
        self.duration2LineEdit.setText("")

        self.trigger2LineEdit = QLabel(self)
        self.trigger2LineEdit.setGeometry(210, 120, 100, 20)
        self.trigger2LineEdit.setText("")

        self.pulseT2LineEdit = QLabel(self)
        self.pulseT2LineEdit.setGeometry(250, 120, 100, 20)
        self.pulseT2LineEdit.setText("")

        self.testID3LineEdit = QLabel(self)
        self.testID3LineEdit.setGeometry(10, 135, 100, 20)
        self.testID3LineEdit.setText("")

        self.portName3LineEdit = QLabel(self)
        self.portName3LineEdit.setGeometry(50, 135, 100, 20)
        self.portName3LineEdit.setText("")

        self.amps3LineEdit = QLabel(self)
        self.amps3LineEdit.setGeometry(120, 135, 100, 20)
        self.amps3LineEdit.setText("")

        self.duration3LineEdit = QLabel(self)
        self.duration3LineEdit.setGeometry(155, 135, 100, 20)
        self.duration3LineEdit.setText("")

        self.trigger3LineEdit = QLabel(self)
        self.trigger3LineEdit.setGeometry(210, 135, 100, 20)
        self.trigger3LineEdit.setText("")

        self.pulseT3LineEdit = QLabel(self)
        self.pulseT3LineEdit.setGeometry(250, 135, 100, 20)
        self.pulseT3LineEdit.setText("")

        progressMax = 10 #hard code dummy value for now
        self.progressBar.setMaximum(progressMax)
        self.progressBar.setMinimum(0)

        self.show()

    def openFileNameDialog(self):    
        options = QFileDialog.Options()
        #options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Select Test Sequence Setup Spreadsheet", "","Excel Files (*.xlsx)", options=options)
        if fileName:
            #testList contains a primary key that represents each Step Num row in the configuration spreadsheet.
            #Each of those keys is a list that contains keys for every column in the configuration spreadsheet.
            #In this way the configuration spreadsheet is completely replecated in testList and the file is no longer used beyond this point.
            self.testList, self.is_okay = misc_Functions.loadTestSequence(fileName)
            
        if self.is_okay != True:
            QMessageBox.about(self, "Test Configuration Failed", "Bad Juju!")
            self.close()

    def openInstruments(self):
        
        rm = visa.ResourceManager()
        instrumentList = rm.list_resources()
        print(instrumentList)
        for loopIndex in range(0, len(instrumentList)):
            try:
                inst = rm.open_resource(instrumentList[loopIndex])
                try:
                    returnString = inst.query("*IDN?")
                    print("ID: " + str(returnString))
                    if returnString.find("CHROMA") != -1:
                        print("found chroma")
                        self.str_load_frame_id = returnString
                        self.str_load_frame_id = self.str_load_frame_id.replace (',','_')
                        self.str_load_frame_id = self.str_load_frame_id.replace ('\n','')
                        self.str_load_frame_id = self.str_load_frame_id.replace ('\r','')
                        print self.str_load_frame_id
        
                        self.loadIDN.setText(returnString) 
                        self.load = inst
                        break
                except visa.VisaIOError:
                    print "Could Not Query Intrument"
            except Exception as e:
                print("Error opening instrument")
                print(e)
     
        tempString = "*RST"
        self.load.write(tempString) 
    
        os_time_initialized = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print os_time_initialized
        
    def closeEvent(self, event):
        print("Closing...")
        self.thread.quit()
        print("Finished")
        event.accept()
    

    
if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    ex = MainWindow()
#    app.exec_()  
    sys.exit(app.exec_())  
