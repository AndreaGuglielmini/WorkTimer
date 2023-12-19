from PyQt5.QtWidgets import *

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
        layout.addWidget(self.LineColumn, row, 1)

        row = row + 1
        self.NameColumnbtn  = QtWidgets.QPushButton('Show')
        self.NameColumnbtn.clicked.connect(self.show)
        self.NameColumnbtn.setMaximumWidth(100)
        layout.addWidget(self.NameColumnbtn, row, 1)
        self.NameColumn  = QLabel('Required:', self)
        self.NameColumn.setStyleSheet("border: 0px solid black")
        layout.addWidget(self.NameColumn, row, 0)

        row=row+1
        endButton = QtWidgets.QPushButton('OK')
        endButton.clicked.connect(self.returnOK)
        endButton.setMaximumWidth(40)
        layout.addWidget(endButton, row, 0)

    def loadini(self):
        self.sheetname=self.configini['PRJ']['sheet']
        self.skiprow=self.configini['PRJ']['skiprow']
        self.minutes=self.configini['PRJ']['stepminutes']
        self.filter=self.configini['PRJ']['filterprojectby']
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


    def changevalue(self):
        try:
            self.configini['PRJ']['sheet']=self.LineSheetName.text()
            self.configini['PRJ']['skiprow']=self.LineSkipRow.text()
            self.configini['PRJ']['stepminutes']=self.LineStepMinutes.text()
            self.configini['PRJ']['filterprojectby']=self.LineFilter.text()

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
    def show(self):
        messageshow('columncustomerprj;columncustomer;columnboard;columnprevhour;columnunitcost;columncost;columncostnotax;columneffectivehour;columnstatus',"ok")

    def returnOK(self):
        self.accept()
        return self.configini


def main():
    app = QtWidgets.QApplication(sys.argv)
    for i in range(10):
        ex = fields(i)
        ex.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        if ex.exec_() == QtWidgets.QDialog.Accepted:
            print(ex.val)





