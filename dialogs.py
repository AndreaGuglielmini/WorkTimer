from PyQt5.QtWidgets import *
import pandas as pd
import matplotlib as mpl
from datetime import datetime
import os

def messageshow(text, style="std", title="Information"):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText(text)
    msg.setWindowTitle(title)
    if style=="std":
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    if style=="ok":
        msg.setStandardButtons(QMessageBox.Ok)
    msg.show()
    retval = msg.exec_()
    if retval == QMessageBox.Ok:
        return True
    else:
        return False
class EditWindowQWidget(QWidget):
    def __init__(self, file, readonly=True, namefile="",windowname="Settings"):
        QWidget.__init__(self)
        Optionlayout = QGridLayout()

        self.setLayout(Optionlayout)
        self.windowname=windowname
        self.file=file
        self.readonly=readonly
        self.namefile=namefile
        self.initUI()

        Optionlayout.addWidget(self.toolbar, 0, 0)
        Optionlayout.addWidget(self.text_edit, 1, 0)

    def initUI(self):

        self.setGeometry(400, 200, 500, 500)
        self.setWindowTitle(self.windowname+" "+self.namefile)

        self.text_edit = QPlainTextEdit(self)
        #self.setCentralWidget(self.text_edit)
        self.text_edit.setPlainText(self.file)
        if self.readonly:
            self.text_edit.setReadOnly(self.readonly)
        self.text_edit.setMinimumWidth(800)
        self.text_edit.setMinimumHeight(800)
        self.init_toolbar()

        self.show()

    def init_toolbar(self):
        # create actions
        if not self.readonly:
            self.save_action = QAction('Save', self)
            self.save_action.setShortcut('Ctrl+S')
            self.save_action.triggered.connect(self.save_file)

        self.quit_action = QAction('Quit', self)
        self.quit_action.setShortcut('Ctrl+x')
        self.quit_action.triggered.connect(self.quit_editor)
        # create toolbar
        self.toolbar = QToolBar("My main toolbar")
        if not self.readonly:
            self.toolbar.addAction(self.save_action)
        self.toolbar.addAction(self.quit_action)

    def save_file(self):
        #filename = QFileDialog.getSaveFileName(self, 'Save file', '', 'Text Files (*.txt)')
        filename = self.namefile
        try:
            with open(filename, 'w') as f:
                f.write(self.text_edit.toPlainText())
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error writing file")
            msg.setWindowTitle("Error")
            msg.exec_()

    def quit_editor(self):
        self.close()

import sys
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QDialog
class fields(QDialog):
    def __init__(self, list,title="", parent=None):
        super(fields, self).__init__(parent)
        self.val = list
        self.title=title
        self.initUI()

    def initUI(self):
        endButton = QtWidgets.QPushButton('OK')
        endButton.clicked.connect(self.on_clicked)
        nokButton = QtWidgets.QPushButton('Cancel')
        nokButton.clicked.connect(self.returnNOK)
        self.oldPNref = QLineEdit()
        self.oldPNref.setPlaceholderText(self.val[0])

        self.newPNref = QLineEdit()
        self.newPNref.setText(self.val[0])



        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.oldPNref)
        layout.addWidget(self.newPNref)
        layout.addWidget(endButton)
        layout.addWidget(nokButton)
        self.setWindowTitle(self.title)

    @QtCore.pyqtSlot()
    def on_clicked(self):
        old=self.oldPNref.text()
        new=self.newPNref.text()
        list=[]
        list.append(self.val[0])
        list.append(new)
        self.val=list
        self.accept()
        return self.val


    def returnNOK(self):
        self.close()


class settings(QDialog):
    def __init__(self, configini,title="", parent=None):
        super(settings, self).__init__(parent)
        self.title=title
        self.configini=configini
        self.initUI()


    def initUI(self):

        layout = QtWidgets.QGridLayout(self)
        self.setWindowTitle(self.title)

        self.loadini()

        row = 0
        self.NameSheet = QLabel('Name of sheet with data', self)
        self.NameSheet.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameSheet, row, 0)
        self.LineSheetName = QLineEdit()
        self.LineSheetName.setText(self.sheetname)
        self.LineSheetName.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineSheetName, row, 1)

        row = row + 1
        self.NameSkipRow  = QLabel('Number of row to skip (First row must contain the column headers)', self)
        self.NameSkipRow.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameSkipRow, row, 0)
        self.LineSkipRow = QLineEdit()
        self.LineSkipRow.setText(self.skiprow)
        self.LineSkipRow.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineSkipRow, row, 1)

        row = row + 1
        self.NameStepMinutes  = QLabel('Minimum slot time of work (minutes)', self)
        self.NameStepMinutes.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameStepMinutes, row, 0)
        self.LineStepMinutes = QLineEdit()
        self.LineStepMinutes.setText(self.minutes)
        self.LineStepMinutes.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineStepMinutes, row, 1)

        row = row + 1
        self.NameFilter  = QLabel('Filter Active project by this field', self)
        self.NameFilter.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameFilter, row, 0)
        self.LineFilter = QLineEdit()
        self.LineFilter.setText(self.filter)
        self.LineFilter.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineFilter, row, 1)

        row = row + 1
        self.listcolumn=';'
        self.NameColumn  = QLabel('Column name (separate by ;)', self)
        self.NameColumn.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameColumn, row, 0)
        self.LineColumn = QLineEdit()
        self.LineColumn.setText(self.listcolumn.join(self.columnname))
        self.LineColumn.returnPressed.connect(self.changevalue)
        self.LineColumn.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn, row, 1)

        row = row + 1
        self.NameColumnbtn  = QtWidgets.QPushButton('Show')
        self.NameColumnbtn.clicked.connect(self.showm)
        self.NameColumnbtn.setMaximumWidth(100)
        layout.addWidget(self.NameColumnbtn, row, 1)
        self.NameColumn  = QLabel('Required:', self)
        self.NameColumn.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameColumn, row, 0)

        row = row + 1
        self.AlwaysNotesbtn  = QtWidgets.QCheckBox()
        self.AlwaysNotesbtn.clicked.connect(self.changevalue)
        self.AlwaysNotesbtn.setMaximumWidth(100)
        if self.alwaysnotes:
            self.AlwaysNotesbtn.setChecked(True)
        else:
            self.AlwaysNotesbtn.setChecked(False)
        layout.addWidget(self.AlwaysNotesbtn, row, 1)
        self.NameColumn  = QLabel('Always ask for notes:', self)
        self.NameColumn.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameColumn, row, 0)

        row=row+1
        endButton = QtWidgets.QPushButton('OK')
        endButton.clicked.connect(self.returnOK)
        endButton.setMaximumWidth(40)
        layout.addWidget(endButton, row, 0)

    def loadini(self):
        try:
            self.sheetname=self.configini['PRJ']['sheet']
            self.skiprow=self.configini['PRJ']['skiprow']
            self.minutes=self.configini['PRJ']['stepminutes']
            self.filter=self.configini['PRJ']['filterprojectby']
            self.alwaysnotes=self.configini.getboolean('PRJ','alwaysnotes')
            self.columnname=[]
            self.columnname.append(str(self.configini['PRJ']['columncustomerprj']))
            self.columnname.append(str(self.configini['PRJ']['columncustomer']))
            self.columnname.append(str(self.configini['PRJ']['columnboard']))
            self.columnname.append(str(self.configini['PRJ']['columnprevhour']))
            self.columnname.append(str(self.configini['PRJ']['columnunitcost']))
            self.columnname.append(str(self.configini['PRJ']['columncost']))
            self.columnname.append(str(self.configini['PRJ']['columncostnotax']))
            self.columnname.append(str(self.configini['PRJ']['columneffectivehour']))
            self.columnname.append(str(self.configini['PRJ']['columnstatus']))
        except Exception as re:
            messageshow("Error loading .ini file. Re-install application or check\n"+str(re),"ok")
            sys.exit()


    def changevalue(self):
        try:
            self.configini['PRJ']['sheet']=self.LineSheetName.text()
            self.configini['PRJ']['skiprow']=self.LineSkipRow.text()
            self.configini['PRJ']['stepminutes']=self.LineStepMinutes.text()
            self.configini['PRJ']['filterprojectby']=self.LineFilter.text()
            if self.AlwaysNotesbtn.isChecked():
                self.configini['PRJ']['alwaysnotes'] = 'True'
            else:
                self.configini['PRJ']['alwaysnotes'] = 'False'
            self.columnname=self.LineColumn.text().split(self.listcolumn)
            self.configini['PRJ']['columncustomerprj']=self.columnname[0]
            self.configini['PRJ']['columncustomer']=self.columnname[1]
            self.configini['PRJ']['columnboard']=self.columnname[2]
            self.configini['PRJ']['columnprevhour']=self.columnname[3]
            self.configini['PRJ']['columnunitcost']=self.columnname[4]
            self.configini['PRJ']['columncost']=self.columnname[5]
            self.configini['PRJ']['columncostnotax']=self.columnname[6]
            self.configini['PRJ']['columneffectivehour']=self.columnname[7]
            self.configini['PRJ']['columnstatus']=self.columnname[8]

        except Exception as re:
            print('Error loading ini file --> ', re)
            self.accept()
            return None
    def showm(self):
        messageshow('columncustomerprj;columncustomer;columnboard;columnprevhour;columnunitcost;columncost;columncostnotax;columneffectivehour;columnstatus',"ok")

    def returnOK(self):
        self.accept()
        return self.configini


class statistics(QDialog):
    def __init__(self, projectfolder,projectname="", daily=False, parent=None):
        super(statistics, self).__init__(parent)
        self.projectname=projectname
        self.projectfolder=projectfolder
        self.daily=daily
        self.spenttime=[]
        self.printable=[]
        self.listofcsv = []
        self.data = []
        self.loadcsv()
        self.initUI()

    def initUI(self):
        layout = QtWidgets.QGridLayout(self)
        self.layout=layout
        self.setWindowTitle(self.projectname)
        self.projects_cb=[]
        self.projects_btn=[]

        row = 0
        for i in range(0, len(self.listofcsv)):
            self.projects_cb.append(QCheckBox(self.listofcsv[i][:-4], self))
            layout.addWidget(self.projects_cb[i], row, 1)
            self.projects_cb[i].setChecked(True)
            row=row+1

        endButton = QtWidgets.QPushButton('OK')
        endButton.clicked.connect(self.returnOK)
        endButton.setMaximumWidth(40)
        layout.addWidget(endButton, row, 0)
        if self.daily:
            self.dofilter()
            for cb in self.projects_cb:
                cb.stateChanged.connect(self.changefilter)

            self.dateedit = QtWidgets.QDateEdit(calendarPopup=True)
            self.dateedit.setDateTime(QtCore.QDateTime.currentDateTime())
            self.dateedit.dateChanged.connect(self.newdate)
            layout.addWidget(self.dateedit, row, 1)

            self.timesum_PB=QProgressBar()

            self.timesum_PB.setStyleSheet('''
                QProgressBar {  background-color: grey;
                                color: white;               /* Text color (not highlighted)
                                border: 2px solid white;      /* Border color */
                                border-radius: 5px;           /* Rounded border edges */
                                margin-left: 2px;
                                margin-right: 2px;
                                text-align: center            /* Center the X% indicator */
                            }
                QProgressBar::chunk{background-color: green;vertical-width: 6px}"
                ''')
            self.timesum_PB.setOrientation(QtCore.Qt.Vertical)
            layout.addWidget(self.timesum_PB, 0, 0,row-2,0)
            self.label = QLabel("..loading..")
            layout.addWidget(self.label, row - 1, 0)
            row=row+1
            next_b=QtWidgets.QPushButton('>')
            next_b.clicked.connect(self.nextday)
            next_b.setMaximumWidth(40)
            prev_b=QtWidgets.QPushButton('<')
            prev_b.clicked.connect(self.prevday)
            prev_b.setMaximumWidth(40)
            layout.addWidget(next_b, row, 1)
            layout.addWidget(prev_b, row, 0)

            self.changedate()
    def nextday(self):
        val=self.dateedit.date().addDays(1)
        self.dateedit.setDateTime(QtCore.QDateTime(val))
        self.changedate()

    def prevday(self):
        val=self.dateedit.date().addDays(-1)
        self.dateedit.setDateTime(QtCore.QDateTime(val))
        self.changedate()

    def changefilter(self):
        isanyone=False
        for cb in self.projects_cb:
            if cb.isChecked():
                isanyone=True

        if isanyone:
            self.dofilter()
            self.changedate(True)

    def newdate(self):
        for cb in self.projects_cb:
            cb.setChecked(True)
        self.changedate()

    def changedate(self,filter=False):
        req_date = self.dateedit.date()
        req_date = req_date.toString('dd/MM/yyyy')
        #testfilter = self.data.query("date_start==@req_date")
        testfilter = self.data[self.data['date_start'] == req_date]
        testfilterlist = testfilter['projectname'].tolist()
        print(testfilter)
        print(testfilterlist)
        if not filter:
            for cb in self.projects_cb:
                if cb.text() in testfilterlist:
                    cb.setEnabled(True)
                else:
                    cb.setEnabled(False)
        #prj = testfilter.tolist('projectname')
        #print(prj)

        timesum = self.data.loc[self.data['date_start'] == req_date, 'timeelapsed'].sum()
        print(timesum)
        if timesum > 0:
            timesum24 = (timesum/24) * 100
        else:
            timesum24 = 0
        if timesum > 24:
            timesum24=0
            timesum="ERR"
        self.timesum_PB.setValue(timesum24)
        print(self.timesum_PB.value())
        self.timesum_PB.setFormat("")
        self.label.setText((str(timesum) + " hour"))


    def loadcsv(self):
        csvfolder=self.projectfolder+"/dbfile/"

        for file in os.listdir(csvfolder):
            if file.lower().endswith('.csv'):
                self.listofcsv.append(file)
        print(self.listofcsv)

    def dofilter(self):
        format = '%d/%m/%Y'
        dateparse = lambda dates: [datetime.strptime(d, format) for d in dates]
        frame = []
        for i in range(0,len(self.projects_cb)):
            if self.projects_cb[i].isChecked():
                filepath = self.projectfolder + "/dbfile/" + self.listofcsv[i]
                try:
                    df = pd.read_csv(filepath, delimiter=";",
                                 names=['projectname', 'date_start', 'time_start', 'date_stop', 'time_stop',
                                        'timeelapsed', 'notes'], parse_dates=True, date_parser=dateparse)
                except:
                    messageshow("Error loading csv file for "+str(filepath))
                frame.append(df)
        self.data = pd.concat(frame, axis=0, ignore_index=True)


    def returnOK(self):
        if not self.daily:
            self.dofilter()
            self.loadcsv()
        self.accept()
        return True

