#===============================================================================================================
# Author:       Guglielmini Andrea
# Date:         13/02/2024
# Application:  WorKTimer Application
# Version:      See version
#===============================================================================================================

import sys, os
version="0.9.6"
myappid = 'AndreaGuglielmini.WorkTimer.GUI'+version # arbitrary string

import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QMessageBox,QFileDialog
from confighandler import *
from os.path import exists

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

# GENERIC CONFIG
configfile="config.ini"
configfilewrp=configfile+".wrp"


if __name__ == "__main__":
    try:
        from pandasgui import show as pandashow

    except Exception as re:
        print(re)
        pandashow=False
    app = QApplication(sys.argv)
    print("###### WORK TIMER ########")
    print("Ver ", version)
    inivalid = False
    # sanity check
    if not os.path.isfile(configfile):
        if not os.path.isfile(configfilewrp):
            filename = QFileDialog.getOpenFileName()

            if filename != "":
                f=open(configfilewrp,"w")
                f.write(filename[0])
                print("Writing: ", filename[0])
                f.close
                configfile=(filename[0])
                inivalid = True
            else:
                messageshow("No valid config.ini file", "ok")
                inivalid = False
        else:
            f=open(configfilewrp,"r")
            configfile=f.read()
            print(configfile)
            if exists(configfile):
                inivalid=True
            else:
                inivalid=False
    else:
        inivalid=True

    if not inivalid:
        print("Loading def .ini")
        import shutil

        shutil.copy2('config_def.ini', 'config.ini')  # complete target filename given
        try:
            os.remove(configfilewrp)
        except:
            pass

    try:
        confighand=confighandler(configfile)
        config=confighand.configreadout()
        print(config)
    except Exception as re:
        print(re)
        messageshow("Error loading config file "+str(re), "ok")
        config=None

    if config is not None:
        try:
            project=config['PRJ']['acutalproject']
        except:
            project=""
    else:
        pass



    if project == "":
        print("No project active")
    else:
        if exists(project):
            print("Working on: ", project)
        else:
            project=""
            messageshow("Project Not Found","ok")




    from WorkTimer_GUI import Window
    screen = Window(version, project, messageshow,confighand,pandashow,configfile)
    screen.show()
    sys.exit(app.exec_())
