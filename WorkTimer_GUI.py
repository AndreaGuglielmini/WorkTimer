# TBD Minutes are displayed without 0
# TBD review prj update
# TBD filter project based on status
# TBD config ini configurable with interface

import datetime
from datetime import timedelta, datetime

from dialogs import *


class Window(QWidget):


    def __init__(self, version, project, feedback,confighand):
        QWidget.__init__(self)
        layout = QGridLayout()
        self.setLayout(layout)
        self.version = version
        self.feedback=feedback

        #Read data from config ini
        self.confighand=confighand
        self.configini = confighand.configreadout()
        self.sheet=self.configini['PRJ']['sheet']
        self.skiprow = int(self.configini['PRJ']['skiprow'])
        self.step = int(self.configini['PRJ']['stepminutes'])
        self.filterprj=str(self.configini['PRJ']['filterprojectby'])


        self.isprojectrunning = self.configini.getboolean('RUN','isprojectrunning')
        if self.isprojectrunning:
            self.actualprojectruntime = self.configini['RUN']['actualprojectruntime']
            self.actualproject = int(self.configini['RUN']['actualproject'])
            self.actualprojectname = str(self.configini['RUN']['actualprojectname'])
        else:
            self.actualproject=0
            self.actualprojectname = ""
            self.actualprojectruntime=0
            self.isprojectrunning=False


        from DB_lib import work_project
        if project != "":
            self.projecthandler=work_project(project,feedback, self.sheet,self.skiprow)
        else:
            self.openprj()
            self.projecthandler = work_project(self.project, feedback, self.sheet,self.skiprow)
        self.listofworks, prjrun=self.projecthandler.updatelist(self.actualprojectname, self.filterprj)
        if prjrun>0:
            self.actualproject=prjrun

        import os
        self.nameprj = os.path.basename(project)
        self.pathprj=os.path.split(project)[0]

        self.setWindowTitle("Work Timer "+ str(version)+"  - PROJECT "+self.nameprj)

        menurow=0
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

        self.OpenPRJbtn = QPushButton('Open Prj', self)
        self.OpenPRJbtn.clicked.connect(self.openprj)
        layout.addWidget(self.OpenPRJbtn, menurow, self.column1)

        self.UpdatePRJbtn = QPushButton('Update Prj', self)
        self.UpdatePRJbtn.clicked.connect(self.updateprj)
        layout.addWidget(self.UpdatePRJbtn, menurow, self.column2)

        self.SavePRJbtn = QPushButton('Save Prj', self)
        self.SavePRJbtn.clicked.connect(self.saveprj)
        layout.addWidget(self.SavePRJbtn, menurow, self.column3)

        self.exittoolbtn = QPushButton('Exit tool', self)
        self.exittoolbtn.clicked.connect(self.exittool)
        layout.addWidget(self.exittoolbtn, menurow, self.column3)

        self.numerorighe = 10
        ### ROW 2 to maxrow ###
        menurow = menurow+1
        self.maxrow = menurow+self.numerorighe
        self.LineAzienda = []
        self.LineCliente = []
        self.LineNomeScheda = []
        self.LineOREp = []
        self.LinePreventivo = []
        self.LineOreLavorate = []
        self.btnstart=[]
        self.listofelements=[]

        linerow=0
        if len(self.listofworks) > self.numerorighe:
            items=self.maxrow       #TBD pulsanti di pagina
        else:
            items=len(self.listofworks)+1

        for index in range(menurow,items):
            self.increaserow(layout, linerow)
            linerow = linerow + 1


        menurow=menurow+linerow
        self.stopbtn = QPushButton('Stop timer', self)
        self.stopbtn.clicked.connect(self.stopcounter)
        layout.addWidget(self.stopbtn, menurow, self.column1)

        self.RunningProject=QLineEdit()
        self.RunningProject.setPlaceholderText("---")
        self.RunningProject.setStyleSheet("border: 1px solid grey;background-color: green ")
        self.RunningProject.setEnabled(False)
        layout.addWidget(self.RunningProject,menurow,self.column2)

        self.LasRunProject=QLineEdit()
        self.LasRunProject.setPlaceholderText("---")
        self.LasRunProject.setStyleSheet("border: 1px solid grey;background-color: green ")
        self.LasRunProject.setEnabled(False)
        layout.addWidget(self.LasRunProject,menurow,self.column3)

        self.ElapsedTimeProject=QLineEdit()
        self.ElapsedTimeProject.setPlaceholderText("---")
        self.ElapsedTimeProject.setStyleSheet("border: 1px solid grey;background-color: green ")
        self.ElapsedTimeProject.setEnabled(False)
        layout.addWidget(self.ElapsedTimeProject,menurow,self.column4)


        self.updateGUI()


    def increaserow(self, layout, linerow):
        self.listofelements.append(linerow)
        self.LineAzienda.append(QLineEdit())
        self.LineAzienda[linerow].setPlaceholderText("")
        self.LineAzienda[linerow].setStyleSheet("border: 1px solid grey;")
        self.LineAzienda[linerow].setEnabled(False)

        self.LineCliente.append(QLineEdit())
        self.LineCliente[linerow].setPlaceholderText("")
        self.LineCliente[linerow].setStyleSheet("border: 1px solid grey;")
        self.LineCliente[linerow].setEnabled(False)

        self.LineNomeScheda.append(QLineEdit())
        self.LineNomeScheda[linerow].setPlaceholderText("")
        self.LineNomeScheda[linerow].setStyleSheet("border: 1px solid grey;")
        self.LineNomeScheda[linerow].setEnabled(False)

        self.LineOREp.append(QLineEdit())
        self.LineOREp[linerow].setPlaceholderText("")
        self.LineOREp[linerow].setStyleSheet("border: 1px solid grey;")
        self.LineOREp[linerow].setEnabled(False)

        self.LinePreventivo.append(QLineEdit())
        self.LinePreventivo[linerow].setPlaceholderText("")
        self.LinePreventivo[linerow].setStyleSheet("border: 1px solid grey;")
        self.LinePreventivo[linerow].setEnabled(False)

        self.LineOreLavorate.append(QLineEdit())
        self.LineOreLavorate[linerow].setPlaceholderText("")
        self.LineOreLavorate[linerow].setStyleSheet("border: 1px solid grey;")
        self.LineOreLavorate[linerow].setEnabled(False)

        self.btnstart.append(QPushButton('Start timer ' + str(linerow), self))
        self.btnstart[linerow].clicked.connect(self.startcounter)
        layout.addWidget(self.btnstart[-1], linerow + 1, self.column7)

        layout.addWidget(self.LineAzienda[-1], linerow + 1, self.column1)
        layout.addWidget(self.LineCliente[-1], linerow + 1, self.column2)
        layout.addWidget(self.LineNomeScheda[-1], linerow + 1, self.column3)
        layout.addWidget(self.LineOREp[-1], linerow + 1, self.column4)
        layout.addWidget(self.LinePreventivo[-1], linerow + 1, self.column5)
        layout.addWidget(self.LineOreLavorate[-1], linerow + 1, self.column6)



    def stopcounter(self):
        self.configini = self.confighand.configreadout()
        runningfromini=self.configini.getboolean('RUN','isprojectrunning')
        if self.isprojectrunning!=runningfromini:
            self.feedback=("Mismatch execution of timer, aborting..")
            self.isprojectrunning = False
            self.actualprojectruntime = 0
            self.actualproject=''
            self.configini['RUN']['isprojectrunning']=self.isprojectrunning
            self.configini['RUN']['actualprojectruntime']=self.actualprojectruntime
            self.configini['RUN']['actualproject']=self.actualproject
            return 0

        if self.isprojectrunning :
            end = datetime.now()

            #Always load from config.ini, because we can arrive from a closed app until running..
            self.isprojectrunning=self.configini.getboolean('RUN','isprojectrunning')
            self.actualprojectruntime=datetime.strptime(self.configini['RUN']['actualprojectruntime'],'%y/%m/%d, %H:%M:%S')
            self.actualproject=int(self.configini['RUN']['actualproject'])
            seconds = (end - self.actualprojectruntime).total_seconds()
            timeelapsed=round(seconds/60/self.step,0)*0.25
            try:
                actualvalue=int(self.listofworks[self.actualproject][self.projecthandler.columnOreLavorate])
            except:
                actualvalue=0
            if actualvalue>0:
                updatevalue=actualvalue+timeelapsed
            else:
                updatevalue=timeelapsed

            self.projecthandler.writevalue(updatevalue, self.actualproject+self.skiprow+2, self.projecthandler.columnOreLavorate, self.sheet)  # TBD skiprow + 2 is not best practice...
            self.csv_handler(self.actualprojectruntime,end)
            self.isprojectrunning = False
            self.actualprojectruntime = 0
            self.actualproject=''
            self.configini['RUN']['isprojectrunning']=str(self.isprojectrunning)
            self.configini['RUN']['actualprojectruntime']=str(self.actualprojectruntime)
            self.configini['RUN']['actualproject']=str(self.actualproject)
            self.confighand.writevalue(self.configini)
            self.updateprj()
            self.updateGUI()

    def startcounter(self):
        sender = self.sender()
        print(sender.text() + ' was pressed')
        self.actualproject=int(sender.text()[-1])
        if self.actualproject>=len(self.listofworks):
            print("Invalid project")
        else:
            self.RunningProject.setText((self.listofworks[self.actualproject][self.projecthandler.columnscheda]))
            self.RunningProject.setStyleSheet("border: 1px solid grey;background-color: yellow ")

            self.actualprojectname=((self.listofworks[self.actualproject][self.projecthandler.columnscheda]))

            if not self.isprojectrunning:
                self.isprojectrunning=True
                now = datetime.now()
                self.actualprojectruntime=now.strftime('%y/%m/%d, %H:%M:%S')


            else:
                self.feedback("Project already running, please stop first","ok")
                return 0
        self.configini['RUN']['actualproject']=str(self.actualproject)
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

        else:
            return None

    def updateprj(self):
        self.projecthandler.read_excel(self.sheet)
        self.projecthandler.loadlistwork()
        self.listofworks = self.projecthandler.listofwork
        print(self.listofworks)

        self.updateGUI()


    def saveprj(self):
        self.feedback("Not implemented","ok")
        pass

    def exittool(self):
        exit()

    def updateGUI(self):
        if self.isprojectrunning:
            self.LasRunProject.setText(self.actualprojectruntime)
            self.LasRunProject.setStyleSheet("border: 1px solid grey;background-color: yellow ")
            self.RunningProject.setText((self.listofworks[self.actualproject][self.projecthandler.columnscheda]))
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
        if len(self.listofworks)<self.numerorighe:
            max=len(self.listofworks)
        else:
            max=self.numerorighe
        for index in range(0, max):
            self.LineAzienda[index].setText(self.listofworks[index][self.projecthandler.columnazienda])
            self.LineCliente[index].setText(self.listofworks[index][self.projecthandler.columncliente])
            self.LineNomeScheda[index].setText(self.listofworks[index][self.projecthandler.columnscheda])
            self.LineOREp[index].setText(str(self.listofworks[index][self.projecthandler.columnOrePrev]))
            self.LinePreventivo[index].setText(str(self.listofworks[index][self.projecthandler.columnPreventivo]))
            self.LineOreLavorate[index].setText(str(self.listofworks[index][self.projecthandler.columnOreLavorate]))

    def csv_handler(self, timestart, timestop):
        csvfile=str(self.pathprj)+"/dbfile/"+(self.listofworks[self.actualproject][self.projecthandler.columnscheda])+".csv"
        with open(csvfile,"a") as f:
            sday=timestart.day
            smonth=timestart.month
            shour=timestart.hour
            sminute=timestart.minute
            sminute=round(sminute/self.step)
            sthour=timestop.hour
            stminute=timestop.minute
            f.write(str(sday)+"/"+str(smonth)+";"+str(shour)+":"+str(sminute)+";"+str(sthour)+":"+str(stminute)+"\n")
        f.close()



