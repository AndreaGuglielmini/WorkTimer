#===============================================================================================================
# Author:       Guglielmini Andrea
# Date:         22/11/2023
# Application:  DB lib
# Version:      1_0 - created
#===============================================================================================================

import os               # cwd request
import pandas as pd
from PyQt5.QtWidgets import QProgressBar, QSplashScreen
from PyQt5.QtGui import QPixmap
from PyQt5 import QtCore
from PyQt5.QtCore import  Qt
import openpyxl

class work_project:
    def __init__(self, excelfile,feedback, sheet, skip=0):


        self.feedback=feedback
        self.excelfile=excelfile
        self.skip=skip
        ## EXCEL FILE STRUCTURE, import from row 5##
        ## AZIENDA | CLIENTE | Nome Scheda | Ore Preventivo | Prezzo Unitario | Preventivo | Preventivo no rivalsa | Ore Lavorate | Stato

        self.setcolumn()

        self.listofwork=[]

        # read Excel file
        self.read_excel(sheet)
        self.loadlistwork()

    def read_excel(self, sheet):

        print("load files DB (can take some minutes)..")
        splash_pix = QPixmap("splash.jpeg", )
        splash = QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
        progressBar = QProgressBar(splash)
        progressBar.setMinimumWidth(200)
        progressBar.setValue(0)
        splash.setMask(splash_pix.mask())
        splash.showMessage("Loading project..", alignment=Qt.AlignBottom, color=Qt.black)
        splash.show()
        self.cat_list_raw=[]
        self.cat_header=[]


        print("loading -> "+self.excelfile)
        splash.showMessage(self.excelfile, alignment=Qt.AlignBottom, color=Qt.black)
        progressBar.setValue(0)
        try:
            cat_dataframe = pd.read_excel(self.excelfile, sheet, skiprows=self.skip)
        except Exception as re:
            print(re)
            print("Requires openpyxl version 3.0.10 to work properly")
            self.feedback("Requires openpyxl version 3.0.10 to work properly\r\n"+"Close all excel file before executing\r\n"+str(re),"ok")
            return 0
        self.cat_header = cat_dataframe.columns.tolist()
        self.cat_list_raw = cat_dataframe.values.tolist()


        try:
            splash.close()
        except Exception as re:
            print(re)
            pass

    def loadlistwork(self):
        self.listofwork=[]
        for x in self.cat_list_raw:
            self.listofwork.append(x)

        print("cat_list_raw: ",self.cat_list_raw)

    def setcolumn(self):
        self.columnazienda=0
        self.columncliente=1
        self.columnscheda=2
        self.columnOrePrev=3
        self.columnPrezzoUnit=4
        self.columnPreventivo=5
        self.columnPreventivoNoRiv=6
        self.columnOreLavorate=7
        self.columnStato = 8

    def writevalue(self, value, row, column,sheet):
        #book=load_workbook(self.excelfile)
        #writer = pd.ExcelWriter(self.excelfile, mode='a', engine='openpyxl')
        writer = openpyxl.load_workbook(self.excelfile)
        worksheet = writer.get_sheet_by_name(sheet)
        #writer.book=book
        columnlist=['A','B','C','D','E','F','G','H','I','L','M']
        writeto = str(columnlist[column])+str(row)
        print(worksheet)
        worksheet[writeto] = value
        writer.save(self.excelfile)
        writer.close()

    def updatelist(self, runningprj, filter):
        found=0
        listofworks=[]
        for element in self.listofwork:
            if element[self.columnStato] == filter:
                listofworks.append(element)
            if element[self.columnscheda]==runningprj:
                listofworks.append(element)
                found=len(listofworks)
        return listofworks, found