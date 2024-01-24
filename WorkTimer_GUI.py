#===============================================================================================================
# Author:       Guglielmini Andrea
# Date:         19/12/2023
# Application:  WorKTimer GUI
# Version:      See WorkTimer.py
#===============================================================================================================
# TBD: add tool button (in place of stastitic button) with:
# Daily effort (work hour)
# Lists of TODOs
# TBD: percentage over 100 management


from PyQt5 import QtWidgets, QtCore


from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QMessageBox,QFileDialog
import datetime
from datetime import timedelta, datetime
import sys
# Local imports
from dialogs import settings
from dialogs import statistics
from DB_lib import work_project

import math

import os

class Window(QWidget):


    def __init__(self, version, project, feedback,confighand,pandashow):
        QWidget.__init__(self)
        self.layout = QGridLayout()
        self.setLayout(self.layout)
        self.version = version
        self.feedback=feedback
        self.pandashow=pandashow

        #Read data from config ini
        try:
            self.confighand = confighand
            self.configini = confighand.configreadout()
            load=self.loadini()
        except:
            load=False
        if not load:
                self.settings(True)


        self.loadtheprj(project, feedback)

        if self.isprojectrunning:
            self.actualprojectruntime = self.configini['RUN']['actualprojectruntime']
            self.actualprojectname = str(self.configini['RUN']['actualprojectname'])
        else:
            self.actualprojectname = ""
            self.actualprojectruntime=0
            self.isprojectrunning=False

        self.updatelist()


        self.nameprj = os.path.basename(project)
        self.pathprj=os.path.split(project)[0]

        self.format = '%y/%m/%d, %H:%M:%S'

        # create menu
        toolbar = WKMenu(self.container, pandashow)
        toolbaractive=True

        self.layout.addWidget(toolbar,0,0)
        menurow=1
        self.column1=0
        self.column2=1
        self.column3=2
        self.column4 = 3
        self.column5 = 4
        self.column6 = 5
        self.column7 = 6
        self.column8 = 7
        self.column9 = 8
        self.column10 = 9
        self.column11 = 10

        ### ROW 1 ###
        if toolbaractive:
            pass
        else:
            self.OpenPRJbtn = QPushButton('Open Prj', self)
            self.OpenPRJbtn.clicked.connect(self.openprj_btn)
            self.layout.addWidget(self.OpenPRJbtn, menurow, self.column1)

            self.UpdatePRJbtn = QPushButton('Update Prj', self)
            self.UpdatePRJbtn.clicked.connect(self.updateprj)
            self.layout.addWidget(self.UpdatePRJbtn, menurow, self.column2)

            self.SavePRJbtn = QPushButton('Save Prj', self)
            self.SavePRJbtn.clicked.connect(self.saveprj)
            self.layout.addWidget(self.SavePRJbtn, menurow, self.column3)
            self.SavePRJbtn.setEnabled(False)       # Not implemented yet --> how can be useful?

            self.EditTimebtn = QPushButton('Modify Start Time', self)
            self.EditTimebtn.clicked.connect(self.editstarttime)
            self.layout.addWidget(self.EditTimebtn, menurow, self.column4)
            if pandashow!=False:
                self.Statsbtn = QPushButton('Statistics', self)
                self.Statsbtn.clicked.connect(self.showstatistics)
                self.layout.addWidget(self.Statsbtn, menurow, self.column5)

            self.Settingsbtn = QPushButton('Settings', self)
            self.Settingsbtn.clicked.connect(self.settings)
            self.layout.addWidget(self.Settingsbtn, menurow, self.column6)

            self.exittoolbtn = QPushButton('Exit tool', self)
            self.exittoolbtn.clicked.connect(self.exittool)
            self.layout.addWidget(self.exittoolbtn, menurow, self.column7)

        self.numerorighe = 10
        ### ROW 2 to maxrow ###
        menurow = menurow+1
        self.maxrow = menurow+self.numerorighe

        # Init of graphics lists
        self.LineHeaders = []
        self.LineAzienda = []
        self.LineCliente = []
        self.LineNomeScheda = []
        self.LineOREp = []
        self.LinePreventivo = []
        self.LineOreLavorate = []
        self.LineOreLavorate_PB = []
        self.btnstart=[]
        self.LineStatus=[]

        self.listofelements=[]

        linerow=0
        self.multipage=False
        self.actualpage=0
        #if len(self.projecthandler.listofwork) > self.numerorighe:
        #    self.items=self.maxrow
        #else:
        #    self.items=len(self.projecthandler.listofwork)+1

        # Add name of columns
        for item in range(0,len(self.HeadersColumn)):
            self.LineHeaders.append(QLineEdit())
            self.LineHeaders[item].setText(self.HeadersColumn[item])
            self.LineHeaders[item].setStyleSheet("color: white;font: bold 16px; border: 1px solid black;background-color:  gray ")
            self.LineHeaders[item].setEnabled(False)
            self.layout.addWidget(self.LineHeaders[item], menurow, item)

        # Add percentage column name
        self.LineHeadersPB=(QLineEdit())
        self.LineHeadersPB.setText("% Time")
        self.LineHeadersPB.setStyleSheet("color: white;font: bold 16px; border: 1px solid black;background-color:  gray; text-align: center")
        self.LineHeadersPB.setEnabled(False)
        self.LineHeadersPB.setMaximumWidth(116)
        self.layout.addWidget(self.LineHeadersPB, menurow, self.column8)

        menurow = menurow + 1
        for index in range(menurow,menurow+self.numerorighe):
            self.increaserow(linerow, menurow+linerow)
            linerow = linerow + 1

        menurow = menurow + linerow+1
        self.stopbtn = QPushButton('Stop timer', self)
        self.stopbtn.clicked.connect(self.stopcounter)
        self.layout.addWidget(self.stopbtn, menurow, self.column1)

        self.RunningProject = QLineEdit()
        self.RunningProject.setPlaceholderText("---")
        self.RunningProject.setStyleSheet("border: 1px solid grey;background-color: green ")
        self.RunningProject.setEnabled(False)
        self.layout.addWidget(self.RunningProject, menurow, self.column2)

        self.LasRunProject = QLineEdit()
        self.LasRunProject.setPlaceholderText("---")
        self.LasRunProject.setStyleSheet("border: 1px solid grey;background-color: green ")
        self.LasRunProject.setEnabled(False)
        self.layout.addWidget(self.LasRunProject, menurow, self.column3)

        self.ElapsedTimeProject = QLineEdit()
        self.ElapsedTimeProject.setPlaceholderText("---")
        self.ElapsedTimeProject.setStyleSheet("border: 1px solid grey;background-color: green ")
        self.ElapsedTimeProject.setEnabled(False)
        self.layout.addWidget(self.ElapsedTimeProject, menurow, self.column4)

        self.FilterBtn = QCheckBox()
        self.FilterBtn.setText("Enable Filter")
        self.FilterBtn.setEnabled(True)
        self.FilterBtn.setChecked(True)
        self.FilterBtn.stateChanged.connect(self.updatefilter)
        self.layout.addWidget(self.FilterBtn, menurow, self.column5)


        self.PrevPagebtn = QPushButton('<-- Page', self)
        self.PrevPagebtn.clicked.connect(self.ChangePage)
        self.layout.addWidget(self.PrevPagebtn, menurow, self.column6)

        self.NextPagebtn = QPushButton('Page -->', self)
        self.NextPagebtn.clicked.connect(self.ChangePage)
        self.layout.addWidget(self.NextPagebtn, menurow, self.column7)

        self.PageText  = QLabel(" / ", self)
        self.PageText.setStyleSheet("border: 0px solid black")
        self.layout.addWidget(self.PageText, menurow, self.column8)

        self.updateGUI()

    def openprj_btn(self):
        self.loadtheprj("", self.feedback)
        self.updateprj()

    def loadtheprj(self, project, feedback):
        ret=True
        if project != "":
            self.projecthandler=work_project(project,feedback, self.sheet,self.columnnamelist,self.skiprow)
        else:
            self.openprj()
            self.projecthandler = work_project(self.project, feedback, self.sheet, self.columnnamelist, self.skiprow)
            load = self.loadini()
            if not load:
                self.feedback("Error in config file, aborting", "ok")
                sys.exit()
        if self.projecthandler.ret is None:
            while ret==True:
                ret=self.feedback("Error loading. Reconfigure?")
                if ret:
                    self.settings(True)
                    self.projecthandler = work_project(project, feedback, self.sheet, self.columnnamelist, self.skiprow)
                    if self.projecthandler.ret is not None:
                        ret=False
                    load = self.loadini()
                    if not load:
                        self.feedback("Error in config file, aborting","ok")
                        sys.exit()
                else:
                    sys.exit()
        else:
            self.HeadersColumn= self.projecthandler.ret
            self.keys=self.projecthandler.keys

    def updatefilter(self):
        if self.FilterBtn.isChecked():
            self.updatelist(True)
        else:
            self.updatelist(False)
        self.actualpage=0       # always restart drawing from first page
        self.updateGUI()

    def updatelist(self, usefilter=True):
        if usefilter:
            filter=self.filterprj
        else:
            filter="all"
        self.projecthandler.loadlistwork()
        self.listofworks, self.dictofworks=self.projecthandler.updatelist(self.actualprojectname, filter)

    def stopcounter(self, modifier=''):
        modifiers = QtWidgets.QApplication.keyboardModifiers()
        self.configini = self.confighand.configreadout()
        runningfromini=self.configini.getboolean('RUN','isprojectrunning')
        alwaysnotes=self.configini.getboolean('PRJ','alwaysnotes')
        if self.isprojectrunning!=runningfromini:
            self.feedback=("Mismatch execution of timer, aborting..")
            self.isprojectrunning = False
            self.actualprojectruntime = 0
            self.configini['RUN']['isprojectrunning']=self.isprojectrunning
            self.configini['RUN']['actualprojectruntime']=self.actualprojectruntime
            return 0

        if self.isprojectrunning :
            notes=""
            end = datetime.now()
            ok=False
            if modifiers == QtCore.Qt.ShiftModifier or modifier=='shift':
                text, ok = QInputDialog().getText(self, "QInputDialog().getText()",
                                                  "Set manual time", QLineEdit.Normal,
                                                  datetime.strftime(end, self.format))
                if ok and text:
                    end = datetime.strptime(text, self.format)
            if modifiers == QtCore.Qt.ControlModifier or modifier=='ctrl' or alwaysnotes==True:
                text1, ok1 = QInputDialog().getText(self, "QInputDialog().getText()",
                                                  "Insert Notes", QLineEdit.Normal,
                                                  "")
                if ok1 and text1:
                    notes = text1
            #Always load from config.ini, because we can arrive from a closed app until running..
            self.isprojectrunning=self.configini.getboolean('RUN','isprojectrunning')
            self.actualprojectruntime=datetime.strptime(self.configini['RUN']['actualprojectruntime'],'%y/%m/%d, %H:%M:%S')

            seconds = (end - self.actualprojectruntime).total_seconds()
            timeelapsed=round(seconds/60/self.step,0)*0.25   #round up
            #timeelapsed=(math.ceil(seconds/60/self.step))*0.25     #math.ceil seems to be aggressive with times!
            if ok:
                self.feedback("Added time: "+str(timeelapsed),"ok")
            try:
                actualvalue=float(self.dictofworks[self.actualprojectname][self.projecthandler.columnEffectiveHour])
            except:
                actualvalue=0
            if actualvalue>0:
                updatevalue=actualvalue+timeelapsed
            else:
                updatevalue=timeelapsed
            row=1+self.skiprow
            for elem in self.listofworks:
                try:
                    row = row + 1
                    temp=elem.index(self.actualprojectname)
                    break
                except:
                    print("not found")

            self.projecthandler.writevalue(updatevalue, self.projecthandler.columnEffectiveHour,row, self.sheet) #TBD
            self.csv_handler(self.actualprojectruntime,end,timeelapsed,notes)
            self.isprojectrunning = False
            self.actualprojectruntime = 0
            self.actualprojectname=''
            self.configini['RUN']['isprojectrunning']=str(self.isprojectrunning)
            self.configini['RUN']['actualprojectruntime']=str(self.actualprojectruntime)
            self.configini['RUN']['actualprojectname']=str(self.actualprojectname)
            self.confighand.writevalue(self.configini)
            self.projecthandler.read_excel(self.sheet)
            self.updatefilter()

    def startcounter(self):
        sender = self.sender()
        name=self.displayed[int(sender.text()[-1])][self.projecthandler.columnboard]
        print(name + ' was pressed')
        invalid=False
        if not (name in self.dictofworks):
            invalid=True
        if invalid:
            print("Invalid project")
        else:
            self.actualprojectname=name

            if not self.isprojectrunning:
                self.isprojectrunning=True
                now = datetime.now()
                self.actualprojectruntime=now.strftime('%y/%m/%d, %H:%M:%S')


            else:
                self.feedback("Project already running, please stop first","ok")
                return 0
        self.configini['RUN']['actualprojectruntime']=str(self.actualprojectruntime)
        self.configini['RUN']['isprojectrunning']=str(self.isprojectrunning)
        self.configini['RUN']['actualprojectname']=str(self.actualprojectname)
        self.confighand.writevalue(self.configini)
        self.updateGUI()

    def openprj(self):
        filename = QFileDialog.getOpenFileName()

        if filename != "":
            self.project=filename[0]
            self.configini['PRJ']['acutalproject']=filename[0]
            self.confighand.writevalue(self.configini)
            self.projecthandler = work_project(self.project, self.feedback, self.sheet, self.columnnamelist, self.skiprow)
        else:
            return None

    def updateprj(self):
        self.projecthandler.read_excel(self.sheet)
        self.updatelist()

        self.updateGUI()


    def saveprj(self):
        self.feedback("Not implemented","ok")
        pass

    def exittool(self):
        sys.exit()

    def updateGUI(self):
        self.setWindowTitle("Work Timer "+ str(self.version)+"  - PROJECT "+self.nameprj)

        # now create the displayed item
        attr_index_active=1
        attr_index_filtered=0

        self.displayed=[]
        if self.FilterBtn.isChecked():
            for ix in range(0, len(self.listofworks)):
                name=self.listofworks[ix][self.projecthandler.columnboard]
                if self.dictofworks[name+"attr"][attr_index_filtered]==True or self.dictofworks[name+"attr"][attr_index_active]==True:
                    self.displayed.append(self.listofworks[ix])
        else:
            self.displayed=self.listofworks


        if len(self.displayed)>self.numerorighe:
            self.PrevPagebtn.setEnabled(True)
            self.NextPagebtn.setEnabled(True)
        else:
            self.PrevPagebtn.setEnabled(False)
            self.NextPagebtn.setEnabled(False)
        if self.actualpage*self.numerorighe > len(self.displayed)-self.numerorighe:
            self.NextPagebtn.setEnabled(False)
        else:
            self.NextPagebtn.setEnabled(True)
        self.PageText.setText("Page: "+str(self.actualpage+1) + " / " + str(math.ceil(len(self.displayed) / self.numerorighe)))



        for index in range(0, self.numerorighe):
            idxl=index + (self.numerorighe * self.actualpage)
            if int(idxl)<len(self.displayed):
                name = self.displayed[idxl][self.projecthandler.columnboard]
                self.LineAzienda[index].setText(str(self.dictofworks[name][self.projecthandler.columncustomerprj]))
                self.LineCliente[index].setText(str(self.dictofworks[name][self.projecthandler.columncustomer]))
                self.LineNomeScheda[index].setText(str(self.dictofworks[name][self.projecthandler.columnboard]))
                self.LineOREp[index].setText(str(self.dictofworks[name][self.projecthandler.columnprevhour]))
                self.LinePreventivo[index].setText(str(self.dictofworks[name][self.projecthandler.columnCost]))

                # --- TBD percentage
                try:
                    percentage = (self.dictofworks[name][self.projecthandler.columnEffectiveHour] / \
                                 self.dictofworks[name][self.projecthandler.columnprevhour])
                    text=int(percentage*100)
                    self.LineOreLavorate_PB[index].setStyleSheet('''
                        QProgressBar {  background-color: white;
                                        color: black;               /* Text color (not highlighted)
                                        border: 2px solid white;      /* Border color */
                                        border-radius: 5px;           /* Rounded border edges */
                                        margin-left: 2px;
                                        margin-right: 2px;           
                                        text-align: center            /* Center the X% indicator */
                                    }
                        QProgressBar::chunk{background-color: #86ED26;width: 6px; margin: 0.5px}"
                        ''')
                except:
                    percentage = 0
                try:
                    a=float(percentage)
                    if a>=1:
                        text=round(percentage*100)
                        percentage=(( a % 1 ))   #Show only the exceeded time
                        self.LineOreLavorate_PB[index].setStyleSheet('''
                            QProgressBar {  background-color: white;
                                            color: black;               /* Text color (not highlighted)
                                            border: 2px solid white;      /* Border color */
                                            border-radius: 5px;           /* Rounded border edges */
                                            margin-left: 2px;
                                            margin-right: 2px;           
                                            text-align: center            /* Center the X% indicator */
                                        }
                            QProgressBar::chunk{background-color: yellow;width: 6px; margin: 0.5px}"
                            ''')
                except:
                    percentage=0



                self.LineOreLavorate_PB[index].setValue(int(percentage*100))
                self.LineOreLavorate_PB[index].setFormat(str(text)+"%")
                self.LineOreLavorate[index].setText(
                    str(self.dictofworks[name][self.projecthandler.columnEffectiveHour]))
                # --- TBD
                self.LineStatus[index].setText(str(self.dictofworks[name][self.projecthandler.columnStatus]))

                self.btnstart[index].setText("Start "+str(index))
                self.btnstart[index].setEnabled(True)
            else:
                self.LineAzienda[index].setText("")
                self.LineCliente[index].setText("")
                self.LineNomeScheda[index].setText("")
                self.LineOREp[index].setText("")
                self.LinePreventivo[index].setText("")
                self.LineOreLavorate[index].setText("")
                self.LineStatus[index].setText("")
                self.btnstart[index].setText("")
                self.LineOreLavorate_PB[index].setValue(0)
                self.btnstart[index].setEnabled(False)

        if self.isprojectrunning:
            self.LasRunProject.setText(self.actualprojectruntime)
            self.LasRunProject.setStyleSheet("border: 1px solid grey;background-color: yellow ")

            self.RunningProject.setText((self.dictofworks[self.actualprojectname][self.projecthandler.columnboard]))
            self.RunningProject.setStyleSheet("border: 1px solid grey;background-color: yellow ")
            time = datetime.now()
            elapsed=(time - datetime.strptime(self.actualprojectruntime,'%y/%m/%d, %H:%M:%S')).total_seconds()
            elapsed = round(elapsed / 60 / self.step, 0) * 0.25
            self.ElapsedTimeProject.setText(str(elapsed))
            self.ElapsedTimeProject.setStyleSheet("border: 1px solid grey;background-color: yellow ")
        else:
            self.LasRunProject.setText("---")
            self.LasRunProject.setStyleSheet("border: 1px solid grey;background-color: green ")
            self.RunningProject.setText("---")
            self.RunningProject.setStyleSheet("border: 1px solid grey;background-color: green ")
            self.ElapsedTimeProject.setText("--")
            self.ElapsedTimeProject.setStyleSheet("border: 1px solid grey;background-color: green ")


    def increaserow(self,linerow, menurow):
        self.listofelements.append(linerow)
        self.LineAzienda.append(QLineEdit())
        self.LineAzienda[linerow].setPlaceholderText("")
        self.LineAzienda[linerow].setStyleSheet("color: black;border: 1px solid gray;")
        self.LineAzienda[linerow].setEnabled(False)

        self.LineCliente.append(QLineEdit())
        self.LineCliente[linerow].setPlaceholderText("")
        self.LineCliente[linerow].setStyleSheet("color: black;border: 1px solid gray;")
        self.LineCliente[linerow].setEnabled(False)

        self.LineNomeScheda.append(QLineEdit())
        self.LineNomeScheda[linerow].setPlaceholderText("")
        self.LineNomeScheda[linerow].setStyleSheet("color: black;border: 1px solid gray;")
        self.LineNomeScheda[linerow].setEnabled(False)

        self.LineOREp.append(QLineEdit())
        self.LineOREp[linerow].setPlaceholderText("")
        self.LineOREp[linerow].setStyleSheet("color: black;border: 1px solid gray;")
        self.LineOREp[linerow].setEnabled(False)

        self.LinePreventivo.append(QLineEdit())
        self.LinePreventivo[linerow].setPlaceholderText("")
        self.LinePreventivo[linerow].setStyleSheet("color: black;border: 1px solid gray;")
        self.LinePreventivo[linerow].setEnabled(False)

        self.LineOreLavorate.append(QLineEdit())
        self.LineOreLavorate[linerow].setPlaceholderText("")
        self.LineOreLavorate[linerow].setStyleSheet("color: black;border: 1px solid gray;")
        self.LineOreLavorate[linerow].setEnabled(False)
        self.LineOreLavorate_PB.append(QProgressBar())
        self.LineOreLavorate_PB[linerow].setMaximumWidth(120)
        self.LineOreLavorate_PB[linerow].setStyleSheet('''
            QProgressBar {  background-color: white;
                            color: black;               /* Text color (not highlighted)
                            border: 2px solid white;      /* Border color */
                            border-radius: 5px;           /* Rounded border edges */
                            margin-left: 2px;
                            margin-right: 2px;           
                            text-align: center            /* Center the X% indicator */
                        }
            QProgressBar::chunk{background-color: #86ED26;width: 6px; margin: 0.5px}"
            ''')
        self.LineOreLavorate_PB[linerow].setValue(0)

        self.LineStatus.append(QLineEdit())
        self.LineStatus[linerow].setPlaceholderText("")
        self.LineStatus[linerow].setStyleSheet("color: black;border: 1px solid gray;")
        self.LineStatus[linerow].setEnabled(False)

        self.btnstart.append(QPushButton('Start timer ' + str(linerow), self))
        self.btnstart[linerow].clicked.connect(self.startcounter)
        self.btnstart[linerow].setMaximumWidth(60)
        self.btnstart[linerow].setStyleSheet('text-align: center;')
        self.layout.addWidget(self.btnstart[-1], menurow, self.column9)

        self.layout.addWidget(self.LineAzienda[-1], menurow, self.column1)
        self.layout.addWidget(self.LineCliente[-1], menurow, self.column2)
        self.layout.addWidget(self.LineNomeScheda[-1], menurow, self.column3)
        self.layout.addWidget(self.LineOREp[-1], menurow, self.column4)
        self.layout.addWidget(self.LinePreventivo[-1], menurow, self.column5)
        self.layout.addWidget(self.LineOreLavorate[-1], menurow, self.column6)
        self.layout.addWidget(self.LineOreLavorate_PB[-1], menurow, self.column8)
        self.layout.addWidget(self.LineStatus[-1], menurow, self.column7)

    def editstarttime(self):
        from dialogs import fields
        if self.isprojectrunning:
            time=self.actualprojectruntime
            listtime=[]
            listtime.append(time)
            listtime.append("")
            modtime=fields(listtime)
            modtime.setAttribute(QtCore.Qt.WA_DeleteOnClose)
            if modtime.exec_() == QtWidgets.QDialog.Accepted:
                print(modtime.val)
                modtime = modtime.val
            try:
                tchange=(datetime.strptime(modtime[1], self.format)-(datetime.strptime(self.actualprojectruntime, self.format)))
                change=tchange.total_seconds()
                print(change)
                if change!=0:
                    print(change)
                    self.feedback("Modified time: "+str(-1*round(change/60,0))+" minutes", 'ok')
                    self.actualprojectruntime=modtime[1]
                    self.configini['RUN']['actualprojectruntime'] = str(self.actualprojectruntime)
                    self.confighand.writevalue(self.configini)
                    self.updateGUI()
                else:
                    self.feedback("No changes")
            except Exception as re:
                self.feedback("Invalid data")
                print(re)
        else:
            self.feedback("Start a timer first!","ok")

    def csv_handler(self, timestart, timestop, addedtime, notes=""):
        row = 1 + self.skiprow
        for elem in self.listofworks:
            try:
                row = row + 1
                temp = elem.index(self.actualprojectname)
                break
            except:
                print("not found")
        name=self.dictofworks[self.actualprojectname][self.projecthandler.columnboard]
        csvfile=str(self.pathprj)+"/dbfile/"+str(name)+".csv"
        statinfo = os.access(csvfile, mode=os.W_OK)
        if not statinfo:
            os.makedirs(os.path.dirname(csvfile), exist_ok=True)
        with open(csvfile,"a") as f:
            syear = str(timestart.year).zfill(2)
            sday=str(timestart.day).zfill(2)
            smonth=str(timestart.month).zfill(2)
            shour=str(timestart.hour).zfill(2)
            sminute=str(timestart.minute).zfill(2)
            #sminute=round(sminute/self.step)
            styear = str(timestop.year).zfill(2)
            stday=str(timestop.day).zfill(2)
            stmonth=str(timestop.month).zfill(2)
            sthour=str(timestop.hour).zfill(2)
            stminute=str(timestop.minute).zfill(2)
            f.write(str(name)+";"+sday+"/"+smonth+"/"+syear+";"+shour+":"+sminute+";"+stday+"/"+stmonth+"/"+styear+";"+sthour+":"+stminute+";"+str(addedtime)+";"+str(notes)+"\n")
        f.close()

    def settings(self, reconfigure=False):
        from dialogs import settings

        self.configini = self.confighand.configreadout()
        setup=settings(self.configini)
        setup.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        if setup.exec_() == QtWidgets.QDialog.Accepted:
            self.configini = setup.configini
            if self.configini is None:
                ret=self.feedback("Error on .ini file occurred, please re-install ini file and reload. Exit?")
                if ret:
                    sys.exit()

        try:
            self.sheet = self.configini['PRJ']['sheet']
            self.skiprow = int(self.configini['PRJ']['skiprow'])
            self.step = int(self.configini['PRJ']['stepminutes'])
            self.filterprj=str(self.configini['PRJ']['filterprojectby'])
            self.columnnamelist=[]
            self.columnnamelist.append(str(self.configini['PRJ']['columncustomerprj']))
            self.columnnamelist.append(str(self.configini['PRJ']['columncustomer']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnboard']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnprevhour']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnUnitCost']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnCost']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnCostNoTax']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnEffectiveHour']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnStatus']))
            self.confighand.writevalue(self.configini)
        except Exception as re:
            self.feedback("Error loading ini, aborting. ", re)
            sys.exit()
        if not reconfigure:
            self.updateprj()

    def ChangePage(self):
        sender = self.sender()
        print(sender.text() + ' was pressed')
        if sender.text()=='Page -->' and ((self.actualpage+1)*self.numerorighe<len(self.listofworks)):
            self.actualpage=self.actualpage+1
        if sender.text() == '<-- Page':
            if self.actualpage>0:
                self.actualpage=self.actualpage-1

        print(self.actualpage)
        self.updateGUI()

    def loadini(self):
        try:
            self.sheet = self.configini['PRJ']['sheet']
            self.skiprow = int(self.configini['PRJ']['skiprow'])
            self.step = int(self.configini['PRJ']['stepminutes'])
            self.filterprj = str(self.configini['PRJ']['filterprojectby'])
            self.alwaysnotes = self.configini.getboolean('PRJ', 'alwaysnotes')
            self.columnnamelist = []
            self.columnnamelist.append(str(self.configini['PRJ']['columncustomerprj']))
            self.columnnamelist.append(str(self.configini['PRJ']['columncustomer']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnboard']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnprevhour']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnUnitCost']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnCost']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnCostNoTax']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnEffectiveHour']))
            self.columnnamelist.append(str(self.configini['PRJ']['columnStatus']))
            self.isprojectrunning = self.configini.getboolean('RUN', 'isprojectrunning')
            return True
        except Exception as re:
            print(re)
            return False

    def showstatistics(self, daily):
        modifiers = QtWidgets.QApplication.keyboardModifiers()
        if daily:
            dailystat = statistics(self.pathprj,self.nameprj,True)  #TBD (placeholder : how many hours of work day by day)
            if dailystat.exec_() == QtWidgets.QDialog.Accepted:
                pass
        else:
            stats = statistics(self.pathprj,self.nameprj)
            stats.setAttribute(QtCore.Qt.WA_DeleteOnClose)
            if stats.exec_() == QtWidgets.QDialog.Accepted:
                self.pandashow(stats.data)

    def container(self, action):
        if action == "open":
            self.openprj()
        if action == "stats":
            self.showstatistics(False)
        if action == "daily":
            self.showstatistics(True)
        if action == "modtime":
            self.editstarttime()
        if action == "settings":
            self.settings()
        if action == "update":
            self.updateprj()
        if action == "stop":
            self.stopcounter('shift')
        if action == "stopnotes":
            self.stopcounter('ctrl')

from PyQt5.QtWidgets import QMainWindow, QAction, qApp, QApplication
from PyQt5.QtGui import QIcon

class WKMenu(QMainWindow):

    def __init__(self, container, panda=False):
        super().__init__()

        self.container=container
        self.panda=panda
        self.initUI()

    def initUI(self):

        exitAct = QAction(QIcon('exit.png'), '&Exit', self)
        exitAct.setShortcut('Ctrl+Q')
        exitAct.setStatusTip('Exit application')
        exitAct.triggered.connect(qApp.quit)

        self.statusBar()

        menubar = self.menuBar()
        menubar.setStyleSheet('''
                QMenuBar {
                    background-color:light grey
                }''')
        fileMenu = menubar.addMenu('&File')

        exitAct.setShortcut('Ctrl+Q')
        exitAct.setStatusTip('Exit application')

        OpenAction = QAction("&Open..", self)
        OpenAction.setShortcut("Ctrl+O")
        OpenAction.setStatusTip('Open xls file')
        OpenAction.triggered.connect(self.open)

        UpdateAction = QAction("&Update from xls", self)
        UpdateAction.setShortcut("Ctrl+U")
        UpdateAction.setStatusTip('Update view from xls file')
        UpdateAction.triggered.connect(self.update)

        #Create File menu entry
        fileMenu.addAction(OpenAction)
        fileMenu.addAction(UpdateAction)
        fileMenu.addSeparator()
        fileMenu.addAction(exitAct)

        #Create Tool menu entry
        toolMenu = menubar.addMenu('&Tool')
        if self.panda:
            StatisticAction = QAction("&Statistic", self)
            StatisticAction.setStatusTip('Open statistic interface')
            StatisticAction.triggered.connect(self.statistic)

        DailyAction = QAction("&Daily work", self)
        DailyAction.setStatusTip('Open daily stats interface')
        DailyAction.triggered.connect(self.daily)

        modtimeAction = QAction("&Modify start time", self)
        modtimeAction.setStatusTip('Modify last start time')
        modtimeAction.triggered.connect(self.modtime)

        stoptimeAction = QAction("&Stop with selection", self)
        stoptimeAction.setStatusTip('Manual Stop Time')
        stoptimeAction.triggered.connect(self.stoptime)

        stoptimenoteAction = QAction("&Stop with notes", self)
        stoptimenoteAction.setStatusTip('Stop and add note')
        stoptimenoteAction.triggered.connect(self.stopnotestime)

        settingsAction = QAction("&Settings", self)
        settingsAction.setStatusTip('Open settings')
        settingsAction.triggered.connect(self.settings)

        #Create Tool menu entry
        if self.panda:
            toolMenu.addAction(StatisticAction)
        toolMenu.addAction(DailyAction)
        toolMenu.addSeparator()
        toolMenu.addAction(modtimeAction)
        toolMenu.addAction(stoptimeAction)
        toolMenu.addAction(stoptimenoteAction)
        toolMenu.addSeparator()
        toolMenu.addAction(settingsAction)


        self.setGeometry(300, 300, 300, 200)
    def open(self):
        self.container("open")

    def statistic(self):
        self.container("stats")

    def daily(self):
        self.container("daily")

    def modtime(self):
        self.container("modtime")

    def settings(self):
        self.container("settings")

    def update(self):
        self.container("update")

    def stoptime(self):
        self.container("stop")

    def stopnotestime(self):
        self.container("stopnotes")