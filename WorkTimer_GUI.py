
import pyqtgraph as pg

import time, datetime
from datetime import timedelta, datetime

from dialogs import *



pg.setConfigOptions(antialias=True, useOpenGL=True)

class Window(QWidget):


    def __init__(self, version, project, dbfile, feedback, config,confighand):
        QWidget.__init__(self)
        layout = QGridLayout()
        self.setLayout(layout)
        self.version = version
        self.feedback=feedback
        self.config=config
        self.confighand=confighand
        self.dbfile=dbfile
        self.sheet='Lavori'  #TBD parametrizzabile in config.ini

        self.actualproject=""
        self.actualprojectruntime=0
        self.isprojectrunning=False
        self.skiprow=5

        from DB_lib import work_project
        if project != "":
            self.projecthandler=work_project(project,feedback, self.sheet)
        else:
            self.openprj()
            self.projecthandler = work_project(project, feedback)
        self.listofworks=self.projecthandler.listofwork

        self.setWindowTitle("Work Timer "+ str(version)+"  - PROJECT "+project)

        menurow=0
        column1=0
        column2=1
        column3=2
        column4 = 3
        column5 = 4
        column6 = 5
        column7 = 6
        column8 = 7
        column9 = 8
        column10 = 9
        column11 = 10

        ### ROW 1 ###

        self.OpenPRJbtn = QPushButton('Open Prj', self)
        self.OpenPRJbtn.clicked.connect(self.openprj)
        layout.addWidget(self.OpenPRJbtn, menurow, column1)

        self.UpdatePRJbtn = QPushButton('Update Prj', self)
        self.UpdatePRJbtn.clicked.connect(self.updateprj)
        layout.addWidget(self.UpdatePRJbtn, menurow, column2)

        self.SavePRJbtn = QPushButton('Save Prj', self)
        self.SavePRJbtn.clicked.connect(self.saveprj)
        layout.addWidget(self.SavePRJbtn, menurow, column3)

        self.exittoolbtn = QPushButton('Exit tool', self)
        self.exittoolbtn.clicked.connect(self.exittool)
        layout.addWidget(self.exittoolbtn, menurow, column3)

        self.numerorighe=10
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
        for index in range(menurow,self.maxrow):
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

            self.btnstart.append(QPushButton('Start timer '+str(linerow), self))
            self.btnstart[linerow].clicked.connect(self.startcounter)
            layout.addWidget(self.btnstart[-1],linerow+1,column7)



            layout.addWidget(self.LineAzienda[-1],linerow+1,column1)
            layout.addWidget(self.LineCliente[-1],linerow+1,column2)
            layout.addWidget(self.LineNomeScheda[-1],linerow+1,column3)
            layout.addWidget(self.LineOREp[-1],linerow+1,column4)
            layout.addWidget(self.LinePreventivo[-1],linerow+1,column5)
            layout.addWidget(self.LineOreLavorate[-1],linerow+1,column6)
            linerow=linerow+1

        self.updateGUI()
        menurow=menurow+linerow
        self.stopbtn = QPushButton('Stop timer', self)
        self.stopbtn.clicked.connect(self.stopcounter)
        layout.addWidget(self.stopbtn, menurow, column1)

        self.RunningProject=QLineEdit()
        self.RunningProject.setPlaceholderText("---")
        self.RunningProject.setStyleSheet("border: 1px solid grey;background-color: green ")
        self.RunningProject.setEnabled(False)
        layout.addWidget(self.RunningProject,menurow,column2)

        self.LasRunProject=QLineEdit()
        self.LasRunProject.setPlaceholderText("---")
        self.LasRunProject.setStyleSheet("border: 1px solid grey;background-color: green ")
        self.LasRunProject.setEnabled(False)
        layout.addWidget(self.LasRunProject,menurow,column3)
    def stopcounter(self):
        if self.isprojectrunning :
            end = datetime.now()
            seconds = (end - self.actualprojectruntime).total_seconds()
            step = 15
            timeelapsed=round(seconds/60/step,0)*0.25
            try:
                actualvalue=int(self.listofworks[self.actualproject][self.projecthandler.columnOreLavorate])
            except:
                actualvalue=0
            if actualvalue>0:
                updatevalue=actualvalue+timeelapsed
            else:
                updatevalue=timeelapsed

            self.projecthandler.writevalue(updatevalue, self.actualproject+self.skiprow, self.projecthandler.columnOreLavorate, self.sheet)  # TBD
            self.isprojectrunning = False
            self.actualprojectruntime=0
            self.updateprj()
            self.updateGUI()

    def startcounter(self):
        #TBD save to inifile the start and the status
        # TBD save to csv file all fields
        sender = self.sender()
        print(sender.text() + ' was pressed')
        self.actualproject=int(sender.text()[-1])
        if self.actualproject>=len(self.listofworks):
            print("Invalid project")
        else:
            self.RunningProject.setText((self.listofworks[self.actualproject][self.projecthandler.columnscheda]))
            self.RunningProject.setStyleSheet("border: 1px solid grey;background-color: yellow ")
            if not self.isprojectrunning:
                self.isprojectrunning=True
                self.actualprojectruntime = datetime.now()
                hour=self.actualprojectruntime.hour
                minute=self.actualprojectruntime.minute
                self.LasRunProject.setText(str(hour)+":"+str(minute))
                self.LasRunProject.setStyleSheet("border: 1px solid grey;background-color: yellow ")
            else:
                self.feedback("Project already running, please stop first")




    def openprj(self):
        filename = QFileDialog.getOpenFileName()

        if filename != "":
            self.project=filename[0]
            self.config['PRJ']['acutalproject']=filename[0]
            self.confighand.writevalue(self.config)
            filename = QFileDialog.getOpenFileName()

            if filename != "":
                self.dbfile = filename[0]
                self.config['PRJ']['actualdatabase'] = filename[0]
                self.confighand.writevalue(self.config)
            else:
                return None

        else:
            return None




    def updateprj(self):
        self.projecthandler.read_excel(self.sheet)
        self.projecthandler.loadlistwork()
        self.listofworks = self.projecthandler.listofwork
        print(self.listofworks)


    def saveprj(self):
        pass

    def exittool(self):
        exit()

    def updateGUI(self):
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
