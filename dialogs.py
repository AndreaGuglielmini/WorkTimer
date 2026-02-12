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
        self.NameColumn  = QLabel('Company;Customer;Prj;Preventived hour;unit cost;cost;cost no tax;hour;status', self)
        self.NameColumn.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameColumn, row, 0)

        row = row + 1
        self.LineColumn0 = QLineEdit()
        self.LineColumn0.setText(self.columncustomerprj)
        self.LineColumn0.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn0, row, 1)

        row = row + 1
        self.LineColumn1 = QLineEdit()
        self.LineColumn1.setText(self.columncustomer)
        self.LineColumn1.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn1, row, 1)

        row = row + 1
        self.LineColumn2 = QLineEdit()
        self.LineColumn2.setText(self.columnboard)
        self.LineColumn2.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn2, row, 1)

        row = row + 1
        self.LineColumn3 = QLineEdit()
        self.LineColumn3.setText(self.columnprevhour)
        self.LineColumn3.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn3, row, 1)

        row = row + 1
        self.LineColumn4 = QLineEdit()
        self.LineColumn4.setText(self.columnunitcost)
        self.LineColumn4.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn4, row, 1)

        row = row + 1
        self.LineColumn5 = QLineEdit()
        self.LineColumn5.setText(self.columncost)
        self.LineColumn5.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn5, row, 1)

        row = row + 1
        self.LineColumn6 = QLineEdit()
        self.LineColumn6.setText(self.columncostnotax)
        self.LineColumn6.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn6, row, 1)

        row = row + 1
        self.LineColumn7 = QLineEdit()
        self.LineColumn7.setText(self.columneffectivehour)
        self.LineColumn7.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn7, row, 1)

        row = row + 1
        self.LineColumn8 = QLineEdit()
        self.LineColumn8.setText(self.columnstatus)
        self.LineColumn8.textChanged.connect(self.changevalue)
        layout.addWidget(self.LineColumn8, row, 1)

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
            self.columncustomerprj=(str(self.configini['PRJ']['columncustomerprj']))
            self.columncustomer=(str(self.configini['PRJ']['columncustomer']))
            self.columnboard=(str(self.configini['PRJ']['columnboard']))
            self.columnprevhour=(str(self.configini['PRJ']['columnprevhour']))
            self.columnunitcost=(str(self.configini['PRJ']['columnunitcost']))
            self.columncost=(str(self.configini['PRJ']['columncost']))
            self.columncostnotax=(str(self.configini['PRJ']['columncostnotax']))
            self.columneffectivehour=(str(self.configini['PRJ']['columneffectivehour']))
            self.columnstatus=(str(self.configini['PRJ']['columnstatus']))
        except Exception as re:
            messageshow("Error loading .ini file. Re-install application or check\n"+str(re),"ok")
            self.sheetname=""
            self.skiprow=""
            self.minutes=""
            self.filter=""
            self.alwaysnotes=False
            self.columncustomerprj=""
            self.columncustomer=""
            self.columnboard=""
            self.columnprevhour=""
            self.columnunitcost=""
            self.columncost=""
            self.columncostnotax=""
            self.columneffectivehour=""
            self.columnstatus=""



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
            self.configini['PRJ']['columncustomerprj']=self.LineColumn0.text()
            self.configini['PRJ']['columncustomer']=self.LineColumn1.text()
            self.configini['PRJ']['columnboard']=self.LineColumn2.text()
            self.configini['PRJ']['columnprevhour']=self.LineColumn3.text()
            self.configini['PRJ']['columnunitcost']=self.LineColumn4.text()
            self.configini['PRJ']['columncost']=self.LineColumn5.text()
            self.configini['PRJ']['columncostnotax']=self.LineColumn6.text()
            self.configini['PRJ']['columneffectivehour']=self.LineColumn7.text()
            self.configini['PRJ']['columnstatus']=self.LineColumn8.text()

        except Exception as re:
            print('Error loading ini file --> ', re)
            self.accept()
            return None
    def showm(self):
        messageshow('columncustomerprj;columncustomer;columnboard;columnprevhour;columnunitcost;columncost;columncostnotax;columneffectivehour;columnstatus',"ok")

    def returnOK(self):
        self.accept()
        return self.configini

import os
from PyQt5.QtWidgets import QDialog, QCheckBox, QLabel, QProgressBar
import PyQt5.QtWidgets as QtWidgets
from PyQt5 import QtCore
import pandas as pd
from datetime import datetime


class statistics(QDialog):
    def __init__(self, project_folder, project_name="", daily=False, parent=None):
        super().__init__(parent)
        self.project_name = project_name
        self.project_folder = project_folder
        self.daily = daily
        self.spent_time = []
        self.printable = []
        self.csv_file_list = []
        self.data = []

        self.load_csv_files()
        self.init_ui()

    def init_ui(self):
        self.layout = QtWidgets.QGridLayout(self)
        self.setWindowTitle(self.project_name)

        self.project_checkboxes = []
        self.setup_project_checkboxes()

        self.setup_buttons()
        if self.daily:
            self.dofilter()
            self.setup_date_controls()
        self.update_progress_bar()

    def setup_project_checkboxes(self):
        max_rows = 8
        for i, csv_file in enumerate(self.csv_file_list):
            column = int(((i / max_rows) % 2)) + 1
            row = i - (column - 1) * max_rows
            checkbox = QCheckBox(csv_file[:-4], self)
            checkbox.setChecked(True)
            self.layout.addWidget(checkbox, row, column)
            self.project_checkboxes.append(checkbox)

    def setup_buttons(self):
        end_button = QtWidgets.QPushButton("OK")
        end_button.clicked.connect(self.return_ok)
        end_button.setMaximumWidth(40)
        self.layout.addWidget(end_button, len(self.project_checkboxes) + 1, 0)

    def setup_date_controls(self):
        self.date_edit = QtWidgets.QDateEdit(calendarPopup=True)
        self.date_edit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.date_edit.dateChanged.connect(self.new_date)
        self.layout.addWidget(self.date_edit, len(self.project_checkboxes) + 2, 1)

        self.time_sum_progress_bar = QProgressBar()
        self.time_sum_progress_bar.setOrientation(QtCore.Qt.Vertical)
        self.layout.addWidget(self.time_sum_progress_bar, 0, 0)

        self.label = QLabel("..loading..")
        self.layout.addWidget(self.label, len(self.project_checkboxes) + 3, 0)

    def new_date(self):
        for checkbox in self.project_checkboxes:
            checkbox.setChecked(True)
        self.changed_date()

    def changed_date(self):
        selected_date = self.date_edit.date().toString("dd/MM/yyyy")
        filtered_data = self.data[self.data["date_start"] == selected_date]

        self.update_checkbox_states(filtered_data)
        self.update_progress_bar(filtered_data)

    def update_checkbox_states(self, filtered_data):
        valid_projects = filtered_data["projectname"].tolist()
        for checkbox in self.project_checkboxes:
            checkbox.setEnabled(checkbox.text() in valid_projects)

    def update_progress_bar(self, filtered_data=None):
        if filtered_data is None:
            self.time_sum_progress_bar.setValue(0)
            self.label.setText("0 hour")
            return

        total_time = filtered_data["timeelapsed"].sum()
        percentage = min((total_time / 24) * 100, 100)
        self.time_sum_progress_bar.setValue(int(percentage))
        self.label.setText(f"{total_time} hour")

    def load_csv_files(self):
        csv_folder = os.path.join(self.project_folder, "dbfile")
        self.csv_file_list = [f for f in os.listdir(csv_folder) if f.lower().endswith(".csv")]

    def dofilter(self):
        frame = []
        for checkbox, csv_file in zip(self.project_checkboxes, self.csv_file_list):
            if checkbox.isChecked():
                file_path = os.path.join(self.project_folder, "dbfile", csv_file)
                try:
                    df = pd.read_csv(
                        file_path,
                        delimiter=";",
                        names=[
                            "projectname",
                            "date_start",
                            "time_start",
                            "date_stop",
                            "time_stop",
                            "timeelapsed",
                            "notes",
                        ],
                        parse_dates=["date_start"],
                    )
                    frame.append(df)
                except Exception as e:
                    print(f"Error loading CSV file {file_path}: {e}")
        self.data = pd.concat(frame, axis=0, ignore_index=True)

    def return_ok(self):
        if not self.daily:
            self.dofilter()
        self.accept()


from PyQt5 import QtCore, QtWidgets

class MagicWizard(QtWidgets.QWizard):
    def __init__(self, configini,filepath,listerr=[], parent=None):
        super(MagicWizard, self).__init__(parent)
        self.addPage(Page1(self))
        self.addPage(Page2(configini,filepath))
        self.addPage(Page3(configini,listerr))
        self.setWindowTitle("WorkTimer Setting Wizard")
        self.resize(640,480)
        self.configini=configini

    def getData(self):
        return self.configini

    def exec_(self, **data):
        return super().exec_()

class Page1(QtWidgets.QWizardPage):
    def __init__(self, parent=None):
        super(Page1, self).__init__(parent)
        layout = QtWidgets.QVBoxLayout()
        self.box=QTextEdit()
        layout.addWidget(self.box)
        self.setLayout(layout)
        self.box.setStyleSheet("color: grey; background-color: light grey; border-width: 4px; font: 20px");
        self.box.setEnabled(False)
        self.box.setText(("Welcome to WorkTimer settings wizard. This will setup the application referring to your "
                          "excel file.\n\n\n"
                          "You need a preformatted Excel, see Exampleproject.xlsx into installation folder.\n\n"
                          "Click next to start.."))


class Page2(QtWidgets.QWizardPage):
    def __init__(self, configini,filepath,parent=None):
        super(Page2, self).__init__(parent)
        layout = QtWidgets.QGridLayout(self)
        self.configini=configini

        self.loadini()

        self.box=QTextEdit()
        layout.addWidget(self.box)
        self.setLayout(layout)
        self.box.setStyleSheet("color: grey; background-color: light grey; border-width: 4px; font: 12px");
        self.box.setEnabled(False)
        self.box.setText(("Insert sheet name containing data table and number of row to skip (first row with column name.\n"
                          "Insert minimum working slot time.\n"
                          "Insert filter value"))
        self.box.setMaximumHeight(80)

        row = 1

        self.NameINI = QLabel('Path of ini file', self)
        self.NameINI.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameINI, row, 0)
        self.NameINIEdit = QLineEdit()
        self.NameINIEdit.setText(filepath)
        self.NameINIEdit.setEnabled(False)
        layout.addWidget(self.NameINIEdit, row, 1)

        row = row + 1

        self.NamePrj = QLabel('Path of Excel table file', self)
        self.NamePrj.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NamePrj, row, 0)
        self.PRJSheetName = QLineEdit()
        self.PRJSheetName.setText(self.projectname)
        self.PRJSheetName.textChanged.connect(self.changevalue)
        layout.addWidget(self.PRJSheetName, row, 1)

        row = row + 1

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

    def loadini(self):
        try:
            self.projectname = self.configini['PRJ']['acutalproject']
            self.sheetname=self.configini['PRJ']['sheet']
            self.skiprow=self.configini['PRJ']['skiprow']
            self.minutes=self.configini['PRJ']['stepminutes']
            self.filter=self.configini['PRJ']['filterprojectby']
        except Exception as re:
            messageshow("Error loading .ini file. Re-install application or check\n"+str(re),"ok")
            self.sheetname=""
            self.skiprow=""
            self.minutes=""
            self.filter=""

    def changevalue(self):
        try:
            self.projectname = self.configini['PRJ']['acutalproject']
            self.configini['PRJ']['acutalproject'] = self.PRJSheetName.text()
            self.configini['PRJ']['sheet']=self.LineSheetName.text()
            self.configini['PRJ']['skiprow']=self.LineSkipRow.text()
            self.configini['PRJ']['stepminutes']=self.LineStepMinutes.text()
            self.configini['PRJ']['filterprojectby']=self.LineFilter.text()


        except Exception as re:
            print('Error loading ini file --> ', re)
            return None

class Page3(QtWidgets.QWizardPage):
    def __init__(self, configini,columnerr=[],parent=None):
        super(Page3, self).__init__(parent)
        layout = QtWidgets.QGridLayout(self)
        self.configini=configini
        self.loadini()
        self.box=QTextEdit()
        #layout.addWidget(self.box)
        self.setLayout(layout)
        self.qlabeltitle = QLabel()
        self.qlabeltitle.setStyleSheet("border-width: 4px; font: 12px");
        self.qlabeltitle.setText(("Insert name of columns accordlying with your excel"))

        self.qlabeltitle1 = QLabel()
        self.qlabeltitle1.setStyleSheet("border-width: 4px; font: 12px");
        self.qlabeltitle1.setText(("Required field are:\n"))

        layout.addWidget(self.qlabeltitle1, 1, 0, 1, 1)
        layout.addWidget(self.qlabeltitle, 0, 0)
        startrow=2
        row = startrow
        self.LineColumn0 = QLineEdit()
        self.LineColumn0.setText(self.columncustomerprj)
        self.LineColumn0.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn0.setStyleSheet("background-color:red")
        layout.addWidget(self.LineColumn0, row, 1)
        self.qlabel0=QLabel()
        self.qlabel0.setText("-> Company name : Use it for identify the main contractor")
        layout.addWidget(self.qlabel0, row, 0)

        row = row + 1
        self.LineColumn1 = QLineEdit()
        self.LineColumn1.setText(self.columncustomer)
        self.LineColumn1.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn1.setStyleSheet("background-color:red")
        layout.addWidget(self.LineColumn1, row, 1)
        self.qlabel1=QLabel()
        self.qlabel1.setText("-> Customer name : Use it for identify final customer")
        layout.addWidget(self.qlabel1, row, 0)

        row = row + 1
        self.LineColumn2 = QLineEdit()
        self.LineColumn2.setText(self.columnboard)
        self.LineColumn2.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn2.setStyleSheet("background-color:red")
        layout.addWidget(self.LineColumn2, row, 1)
        self.qlabel2=QLabel()
        self.qlabel2.setText("-> Project name: unique Identify the project. Will be use for data collection")
        layout.addWidget(self.qlabel2, row, 0)

        row = row + 1
        self.LineColumn3 = QLineEdit()
        self.LineColumn3.setText(self.columnprevhour)
        self.LineColumn3.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn3.setStyleSheet("background-color:red")
        layout.addWidget(self.LineColumn3, row, 1)
        self.qlabel3=QLabel()
        self.qlabel3.setText("-> Estimated hour of work: provisional")
        layout.addWidget(self.qlabel3, row, 0)

        row = row + 1
        self.LineColumn4 = QLineEdit()
        self.LineColumn4.setText(self.columnunitcost)
        self.LineColumn4.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn4.setStyleSheet("background-color:red")
        layout.addWidget(self.LineColumn4, row, 1)
        self.qlabel4=QLabel()
        self.qlabel4.setText("-> Unit cost: cost for hour")
        layout.addWidget(self.qlabel4, row, 0)

        row = row + 1
        self.LineColumn5 = QLineEdit()
        self.LineColumn5.setText(self.columncost)
        self.LineColumn5.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn5.setStyleSheet("background-color:red")
        layout.addWidget(self.LineColumn5, row, 1)
        self.qlabel5=QLabel()
        self.qlabel5.setText("-> Cost: Cost * Estimated")
        layout.addWidget(self.qlabel5, row, 0)

        row = row + 1
        self.LineColumn6 = QLineEdit()
        self.LineColumn6.setText(self.columncostnotax)
        self.LineColumn6.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn6.setStyleSheet("background-color:red")
        layout.addWidget(self.LineColumn6, row, 1)
        self.qlabel6=QLabel()
        self.qlabel6.setText("-> Cost no tax: used for calculate effective cost")
        layout.addWidget(self.qlabel6, row, 0)

        row = row + 1
        self.LineColumn7 = QLineEdit()
        self.LineColumn7.setText(self.columneffectivehour)
        self.LineColumn7.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn7.setStyleSheet("background-color:red;color:white")
        layout.addWidget(self.LineColumn7, row, 1)
        self.qlabel7=QLabel()
        self.qlabel7.setText("-> Hour: Effective hour of work, this field will be updated by application")
        layout.addWidget(self.qlabel7, row, 0)

        row = row + 1
        self.LineColumn8 = QLineEdit()
        self.LineColumn8.setText(self.columnstatus)
        self.LineColumn8.textChanged.connect(self.changevalue)
        if row-startrow in columnerr:
            self.LineColumn8.setStyleSheet("background-color:red")
        layout.addWidget(self.LineColumn8, row, 1)
        self.qlabel8=QLabel()
        self.qlabel8.setText("-> Status: status of project. Used colum for filter")
        layout.addWidget(self.qlabel8, row, 0)



    def loadini(self):
        try:
            self.projectname = self.configini['PRJ']['acutalproject']
            self.sheetname = self.configini['PRJ']['sheet']
            self.skiprow = self.configini['PRJ']['skiprow']
            self.minutes = self.configini['PRJ']['stepminutes']
            self.filter = self.configini['PRJ']['filterprojectby']
            self.alwaysnotes = self.configini.getboolean('PRJ', 'alwaysnotes')
            self.columncustomerprj = (str(self.configini['PRJ']['columncustomerprj']))
            self.columncustomer = (str(self.configini['PRJ']['columncustomer']))
            self.columnboard = (str(self.configini['PRJ']['columnboard']))
            self.columnprevhour = (str(self.configini['PRJ']['columnprevhour']))
            self.columnunitcost = (str(self.configini['PRJ']['columnunitcost']))
            self.columncost = (str(self.configini['PRJ']['columncost']))
            self.columncostnotax = (str(self.configini['PRJ']['columncostnotax']))
            self.columneffectivehour = (str(self.configini['PRJ']['columneffectivehour']))
            self.columnstatus = (str(self.configini['PRJ']['columnstatus']))
        except Exception as re:
            messageshow("Error loading .ini file. Re-install application or check\n" + str(re), "ok")
            self.sheetname = ""
            self.skiprow = ""
            self.minutes = ""
            self.filter = ""
            self.alwaysnotes = False
            self.columncustomerprj = ""
            self.columncustomer = ""
            self.columnboard = ""
            self.columnprevhour = ""
            self.columnunitcost = ""
            self.columncost = ""
            self.columncostnotax = ""
            self.columneffectivehour = ""
            self.columnstatus = ""

    def changevalue(self):
        try:
            self.configini['PRJ']['columncustomerprj'] = self.LineColumn0.text()
            self.configini['PRJ']['columncustomer'] = self.LineColumn1.text()
            self.configini['PRJ']['columnboard'] = self.LineColumn2.text()
            self.configini['PRJ']['columnprevhour'] = self.LineColumn3.text()
            self.configini['PRJ']['columnunitcost'] = self.LineColumn4.text()
            self.configini['PRJ']['columncost'] = self.LineColumn5.text()
            self.configini['PRJ']['columncostnotax'] = self.LineColumn6.text()
            self.configini['PRJ']['columneffectivehour'] = self.LineColumn7.text()
            self.configini['PRJ']['columnstatus'] = self.LineColumn8.text()

        except Exception as re:
            print('Error loading ini file --> ', re)
            return None
