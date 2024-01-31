#! /usr/bin/env python
from ast import Index
import sys, serial, glob, os,csv,time
from xmlrpc.client import DateTime
from openpyxl import Workbook

from xml.etree.ElementInclude import include
from PyQt6.QtWidgets import QApplication
from PyQt6.QtWidgets import QDialog
from PyQt6.QtWidgets import QDialogButtonBox
from PyQt6.QtWidgets import QFormLayout
from PyQt6.QtWidgets import QLineEdit
from PyQt6.QtWidgets import QVBoxLayout
from PyQt6.QtWidgets import QMainWindow
from PyQt6.QtWidgets import QMessageBox
from PyQt6 import QtCore, QtGui
from PyQt6.QtCore import QTime,QTimer,QDateTime

import gui as d
import functions as f
wb = Workbook()
global serOpen
global ser
global portnum

##global errorFlag
serOpen = False
errorFlag = False
portnum = ""
Rx=""
CorrC= [0.048828125,2.5,0.048828125,2.5]
ser = serial.Serial()

class Window(QMainWindow, d.Ui_Dialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.butScan.clicked.connect(self.butScan_click)
        self.butCon.clicked.connect(self.butCon_click)
        self.butZeroA1.clicked.connect(self.butZeroA1_click)
        self.butZeroB1.clicked.connect(self.butZeroB1_click)
        self.butZeroA2.clicked.connect(self.butZeroA2_click)
        self.butZeroB2.clicked.connect(self.butZeroB2_click)
        self.butZeroAll.clicked.connect(self.zeroAll_click)
        self.butRecord.clicked.connect(self.butREC_click)
        self.butClearWin.clicked.connect(self.ClearWin)
        self.butSnip.clicked.connect(self.snip_click)
        self.butCapture.clicked.connect(self.butCapture_click) ## temporary for req command
        self.butSnip.setEnabled(False)
        self.timer = QTimer()
        self.timerREC = QTimer()
        self.rec = False
        self.fileName = time.strftime("%Y%m%d_%I%M%S")+".csv"
        self.label_10.setText(self.fileName)
        self._path = os.path.dirname(os.path.realpath(__file__)) + '\\' + self.fileName   
        self.createFile()
        self.lcdENA = [False,False,False,False]
        self.offset = [0,0,0,0] 
        self.vals = ["","","",""]
        #self.timer.setInterval(100)  # msecs 100 = 1/10th sec
        #self.timer.timeout.connect(self.Tx)
        self.timerREC.timeout.connect(self.butCapture_click)

    def butREC_click(self):
        if (not self.rec):
            self.butRecord.setText("STOP REC")
            self.labelRec.setText("RECORDING ..")
            self.rec = True
            frq = self.spinRec_rate.value()
            self.timerREC.start(int(1000/frq))
            self.spinRec_rate.setEnabled(False)
        elif (self.rec):
            self.butRecord.setText("REC")
            self.labelRec.setText("")
            self.rec = False
            self.timerREC.stop()
            self.spinRec_rate.setEnabled(True)


    def ClearWin(self):
        self.textEditA.clear()
        self.textEditB.clear()

    def butScan_click(self):
        self.comboCom.clear()
        comlist=[]
        comlist.clear()
        comlist = f.COM_setup().serial_ports()
        idlist=[]
        idlist.clear()
        idlist = f.COM_setup().serial_ports()
        self.comboCom.addItems(comlist[0])
        print(comlist[0])
        if comlist[0] != []:
            self.butCon.setEnabled(True)
        else:
            self.butCon.setEnabled(False) 

    def butCon_click(self):
        global serOpen
        global portnum
        index = self.comboCom.currentIndex()
        portnum = f.COM_setup().get_port(index)
        print("HI and CONNECT " + portnum)
        if (serOpen == False):
            self.butCon.setText("DISCONNECT") 
            self.spinRate.setEnabled(False)           
            self.thread = QtCore.QThread(self)
            self.worker = Worker()
            self.worker.moveToThread(self.thread)
            self.worker.rx_frame.connect(self.showRx)
            self.thread.started.connect(self.worker.gen)            
            self.thread.start()  
            self.butZeroAll.setEnabled(True)
            self.butClearWin.setEnabled(True)
            self.butRecord.setEnabled(True)
            #timerVal = (int)(1000/(self.spinRate.value()))
            #print(timerVal)
            #self.timer.start(timerVal)
            self.butCapture.setEnabled(True)
            self.butSnip.setEnabled(True)
            return
        elif (serOpen == True):  
            self.butCon.setText("CONNECT")
            #self.timer.stop()
            serOpen = False
            ser.flush()  
            self.butCapture.setEnabled(False)
            if (ser.isOpen()):
                ser.close()
                print("COM Closed")
                self.butZeroAll.setEnabled(False)
                self.butSnip.setEnabled(False)
                #self.spinRate.setEnabled(True)  
                self.butClearWin.setEnabled(False)
                self.butRecord.setEnabled(False)  
            self.thread.finished.connect(self.worker.quit)
            self.thread.quit() 
            return

    def showRx(self,frame):
        #self.lineRx.setText(frame)
        #print(frame)
        self.vals = frame.split(";")
        #print(len(self.vals))
        for i in range(len(self.vals)):
            #print(self.vals[i]) 
            value_off= float(self.vals[i]) - self.offset[i]
            self.updateLCD(i,value_off)

    def createFile(self):
        headerA = "Signal A1;Signal B1"
        headerB = "Signal A2;Signal B2 \n"
        with open(self.fileName, 'w',newline='') as file:
            file.write(headerA+headerB)
            self.textEditA.append(headerA +"\n")
            self.textEditB.append(headerB)
            file.close()
    
    def butCapture_click(self):
        _vals =""
        with open(self.fileName, 'a',newline='') as file:
            lineA= ""
            lineB= ""
            if(self.chckA1.isChecked()):lineA += (str(self.lcdA1.value())+";")
            else: lineA +=("-;")
            if(self.chckB1.isChecked()):lineA +=(str(self.lcdB1.value())+";")
            else: lineA +=("-;")
            if(self.chckA2.isChecked()):lineB +=(str(self.lcdA2.value())+";")
            else: lineB +=("-;")
            if(self.chckB2.isChecked()):lineB +=(str(self.lcdB2.value())+";")
            else: lineB +=("-;")
            file.write(lineA+lineB)
            file.write("\n")
            self.textEditA.append(lineA)
            self.textEditB.append(lineB)
            file.close()
        return
   
 
    def updateLCD(self,channel,val):
        if (channel == 0):
            if (val == 9999):
                self.lblStateA1.setText("No Connection")
                self.lcdA1.setProperty("value",0)
                self.butZeroA1.setEnabled(False)
                self.lcdENA[channel] = False
            else:
                self.lblStateA1.setText("Connected")
                self.lcdA1.setProperty("value",round(val*self.spinCorrA1.value(),2))
                self.butZeroA1.setEnabled(True)
                self.lcdENA[channel] = True
        elif (channel == 1):
            if (val == 9999):
                self.lblStateB1.setText("No Connection")
                self.lcdB1.setProperty("value",0)
                self.butZeroB1.setEnabled(False)
                self.lcdENA[channel] = False
            else:
                self.lblStateB1.setText("Connected")
                self.lcdB1.setProperty("value",round(val*self.spinCorrB1.value(),0))
                self.butZeroB1.setEnabled(True)
                self.lcdENA[channel] = True
        elif (channel == 2):
            if (val == 9999):
                self.lblStateA2.setText("No Connection")
                self.lcdA2.setProperty("value",0)
                self.butZeroA2.setEnabled(False)
                self.lcdENA[channel] = False
            else:
                self.lblStateA2.setText("Connected")
                self.lcdA2.setProperty("value",round(val*self.spinCorrA2.value(),2))
                self.butZeroA2.setEnabled(True)
                self.lcdENA[channel] = True
        elif (channel == 3):
            if (val == 9999):
                self.lblStateB2.setText("No Connection")
                self.lcdB2.setProperty("value",0)
                self.butZeroB2.setEnabled(False)
                self.lcdENA[channel] = False
            else:
                self.lblStateB2.setText("Connected")
                self.lcdB2.setProperty("value",round(val*self.spinCorrB2.value(),0))
                self.butZeroB2.setEnabled(True)
                self.lcdENA[channel] = True
       

    def butZeroA1_click(self):
        self.offset[0] = float(self.vals[0])
    def butZeroB1_click(self):
        self.offset[1] = int(self.vals[1])
    def butZeroA2_click(self):
        self.offset[2] = float(self.vals[2])
    def butZeroB2_click(self):
        self.offset[3] = int(self.vals[3])
    def zeroAll_click(self):
        for i_id, i in enumerate(self.lcdENA):
            #print(i_id, i)
            if self.lcdENA[i_id] == True:
                self.offset[i_id] = float(self.vals[i_id])
    def snip_click(self):   
        self.printMSG()
            
    def fillCcorr(self):
        self.spinCorrA1.setValue(CorrC[0])
        self.spinCorrB1.setValue(CorrC[1])
        self.spinCorrA2.setValue(CorrC[2])
        self.spinCorrB2.setValue(CorrC[3])

    def printMSG(self):
              
        msgGen = ("Channel A1: " + str(self.lcdA1.value()) + " mm\r\n")
        msgGen += ("Channel B1: " + str(self.lcdB1.value()) + " \u00b5m\r\n")
        msgGen += ("Channel A2: " + str(self.lcdA2.value()) + " mm\r\n")
        msgGen += ("Channel B2: " + str(self.lcdB2.value()) + " \u00b5m\r\n")
        msg = QMessageBox.information(self, "Actual Values",
        msgGen,
        QMessageBox.StandardButton.Close)

class Worker(QtCore.QObject):

    rx_frame = QtCore.pyqtSignal(str)
   
    
    def quit(self):
        print("COM Closed2")
 
    def gen(self):
        global serOpen
        global ser
        print(portnum) 
        ser = serial.Serial(portnum, 9600,writeTimeout = 0, timeout = None, rtscts=False, dsrdtr=False)
        ser.setDTR(0)
        time.sleep(1)
        print("CTR COM OPENED")
        serOpen = True
        sturtup = 1
        # hex_bytes = int(ser_bytes.encode('hex'),16)
        # ser.write("HI".encode())
        window = bytearray(20)
        counter = 0
        bit_flag = 0
        messageType = 0
        CK_A = 0
        CK_B = 0
        stuff = 0
        while serOpen:                    
            if ser.inWaiting():
                ser_bytes = ser.read()
                hex_bytes = int(ser_bytes.hex(),16)
                #print(hex(hex_bytes)[2:])
                #print(ser_bytes)
                if hex_bytes == 0x7E and bit_flag == 0:
                    bit_flag = 1
                    Rx = (hex(hex_bytes)[2:] + " ")
                    #print("Frame Recon0")
                elif hex_bytes == 0xFF and bit_flag == 1:
                    bit_flag = 2
                    Rx += (hex(hex_bytes)[2:] + " ")
                    #print("Frame Recon1")
                elif (hex_bytes == 0x0 and bit_flag == 2): ## status frame 
                    bit_flag = 3
                    messageType = hex_bytes
                    Rx += (hex(hex_bytes)[2:] + " ")
                    CK_A = 0
                    CK_B = 0
                    #print("Frame Recon2")
                elif bit_flag == 3:
                    Rx += (hex(hex_bytes)[2:] + " ")
                    #print("Frame Recon3")
                    frame_len = hex_bytes
                    #print(frame_len)
                    bit_flag = 4
                elif bit_flag == 4:
                    window[counter] = hex_bytes
                    Rx += (hex(hex_bytes)[2:] + " ")
#                    print(hex(hex_bytes)[2:])
##                  CHECKSUM
                    if counter < frame_len:
                        CK_A += window[counter]
                        CK_B += CK_A
                        CK_A &= 0xff
                        CK_B &= 0xff
                    Rx = Rx.upper()
                    #END of FRAME  
                    if hex_bytes == 0x7E and counter == frame_len + 2:
                        #print (hex(CK_A))
                        #print(hex(CK_B))
                        if window[counter-2] == CK_A and window[counter-1] == CK_B:
                            #print ("CHECKSUM OK - proceed")
                            #unstuffing
                            stuff = 0
                            for i in range(0, frame_len):
                                window[i-stuff]= window[i]
                                if window[i] == 0x5D and window[i-1] == 0x7D:
                                    stuff += 1
                                    #print("unstuffing 0x7D")
                                elif window[i] == 0x5E and window[i-1] == 0x7D:
                                    window[i-1] = 0x7E
                                    stuff += 1
                                    #print("unstuffing 0x7E")
                            #status Frame
                            if (messageType == 0x00): #Status frame
                                #print(Rx)
                                A1 = int.from_bytes([window[0],window[1],window[2]], byteorder='big', signed=True)
                                A2 = int.from_bytes([window[3],window[4],window[5]], byteorder='big', signed=True)
                                B1 = int.from_bytes([window[6],window[7],window[8],window[9]], byteorder='big', signed=True)
                                B2 = int.from_bytes([window[10],window[11],window[12],window[13]], byteorder='big', signed=True)
                                str_result_array = str(A1)+";"+str(B1)+";"+str(A2)+";"+str(B2)
                                #print(str_result_array)
                                self.rx_frame.emit(str_result_array)
                        else:
                            print ("CHECKSUM ERROR")
                            print (window.hex()+" ")
                            print(frame_len)
                            print (counter)
                            print (hex(CK_A))
                            print(hex(CK_B))
                        bit_flag = 0
                        counter = 0
                        stuff = 0
                    else:
                        counter += 1 
                
                if counter >=20:
                    counter = 0
                    bit_flag = 0
                    stuff = 0

        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Window()
    win.fillCcorr()
    win.show()
    sys.exit(app.exec())
