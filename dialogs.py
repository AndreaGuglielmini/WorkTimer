from PyQt5.QtWidgets import *


class EditWindowQWidget(QWidget):
    def __init__(self, file, readonly=True, namefile="",windowname="PL303 Editor"):
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
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QDialog
class fields(QDialog):
    def __init__(self, list,title="", parent=None):
        super(fields, self).__init__(parent)
        self.val = list
        self.title=title
        self.val="null"
        self.initUI()

    def initUI(self):
        endButton = QtWidgets.QPushButton('OK')
        endButton.clicked.connect(self.on_clicked)
        nokButton = QtWidgets.QPushButton('Cancel')
        nokButton.clicked.connect(self.returnNOK)
        self.oldPNref = QLineEdit()
        self.oldPNref.setPlaceholderText("Old Partnumber")

        self.newPNref = QLineEdit()
        self.newPNref.setPlaceholderText("new PARTNUMBER")



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
        list.append(old)
        list.append(new)
        self.val=list
        self.accept()
        return self.val


    def returnNOK(self):
        self.close()

def main():
    app = QtWidgets.QApplication(sys.argv)
    for i in range(10):
        ex = fields(i)
        ex.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        if ex.exec_() == QtWidgets.QDialog.Accepted:
            print(ex.val)





