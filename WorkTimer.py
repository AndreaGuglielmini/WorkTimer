import os
import sys
version="0.9.0"
myappid = 'AndreaGuglielmini.WorkTimer.GUI'+version # arbitrary string

import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QMessageBox
from confighandler import *

# GENERIC CONFIG
configfile="config.ini"

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


if __name__ == "__main__":

    app = QApplication(sys.argv)
    print("###### WORK TIMER ########")
    print("Ver ", version)

    try:
        confighand=confighandler(configfile)
        config=confighand.configreadout()
        print(config)
    except Exception as re:
        print(re)
        messageshow("Error loading config file "+str(re), "ok")
        config=None

    if config is not None:
        project=config['PRJ']['acutalproject']
        db_csv=config['PRJ']['actualdatabase']
    else:
        pass #TBD select project file

    from os.path import exists

    if project == "":
        print("No project active")
    else:
        if exists(project):
            print("Working on: ", project)
        else:
            project=""
            messageshow("Project Not Found","ok")




    from WorkTimer_GUI import Window
    screen = Window(version, project,db_csv, messageshow, config,confighand)
    screen.show()

    sys.exit(app.exec_())