#===============================================================================================================
# Author:       Guglielmini Andrea
# Date:         22/11/2023
# Application:  DB lib
# Version:      1_0 - created
# Use template excel file. If some data validation is added, follow this guide https://www.reddit.com/r/learnpython/comments/mfy9qa/openpyxl_deleting_data_validation/
#===============================================================================================================

import pandas as pd
from PyQt5.QtWidgets import QProgressBar, QSplashScreen
from PyQt5.QtGui import QPixmap
from PyQt5 import QtCore
from PyQt5.QtCore import  Qt
import openpyxl
import xlwings as xl

class work_project:
    def __init__(self, excelfile,feedback, sheet,columnnamelist, skip=0):


        self.feedback=feedback
        self.excelfile=excelfile
        self.skip=skip
        self.keys = ["name", "filtered", "active", "index","status","customer","workhour","cost"]
        ## EXCEL FILE STRUCTURE configurable with config.ini



        self.listofwork=[]

        # read Excel file
        self.read_excel(sheet)
        self.ret=self.setcolumn(columnnamelist)
        if self.ret is None:
            return None
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
            app = xl.App(visible=False)         #workaround found at https://stackoverflow.com/questions/71789086/pandas-read-excel-with-formulas-and-get-values
            book = app.books.open(self.excelfile)
            book.save()
            app.kill()
            cat_dataframe = pd.read_excel(self.excelfile, sheet, skiprows=self.skip)
            progressBar.setValue(100)
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

    def setcolumn(self,columnnamelist):
        usedcolumn=[]
        for item in columnnamelist:
            for idx in range(0,len(self.cat_header)):
                if item == self.cat_header[idx]:
                    if item==columnnamelist[0]:
                        self.columncustomerprj=idx
                        usedcolumn.append(item)
                    if item==columnnamelist[1]:
                        self.columncustomer=idx
                        usedcolumn.append(item)
                    if item==columnnamelist[2]:
                        self.columnboard=idx
                        usedcolumn.append(item)
                    if item==columnnamelist[3]:
                        self.columnprevhour=idx
                        usedcolumn.append(item)
                    #if item==columnnamelist[4]:
                        #self.columnUnitCost=idx
                        #usedcolumn.append(item)
                    if item==columnnamelist[5]:
                        self.columnCost=idx
                        usedcolumn.append(item)
                    if item==columnnamelist[6]:
                        self.columnCostNoTax=idx
                    if item==columnnamelist[7]:
                        self.columnEffectiveHour=idx
                        usedcolumn.append(item)
                    if item==columnnamelist[8]:
                        self.columnStatus=idx
                        usedcolumn.append(item)
        # The usedcolumn.append(item) is a workaround for handle the list of columns headers (TBD)
        print(columnnamelist)
        try:
            print(self.columncustomerprj,self.columncustomer,self.columnboard,self.columnStatus, self.columnCost)
            return usedcolumn
        except:
            return None

    def writevalue(self, value, column,row,sheet):
        #book=load_workbook(self.excelfile)
        #writer = pd.ExcelWriter(self.excelfile, mode='a', engine='openpyxl')
        writer = openpyxl.load_workbook(self.excelfile)
        worksheet = writer.get_sheet_by_name(sheet)
        #writer.book=book
        columnlist=['A','B','C','D','E','F','G','H','I','L','M']
        writeto = str(columnlist[column])+str(row)
        print(worksheet)
        worksheet[writeto] = value
        try:
            writer.save(self.excelfile)
        except:
            ret=self.feedback("Error, file open or not found, retry?")
            if ret:
                writer.save(self.excelfile)
        writer.close()

    def updatelist(self, runningprj, filter):
        dictofworks={}
        cycle=0

        for element in self.listofwork:
            if element[self.columnStatus] == filter:
                isfiltered=True
            else:
                isfiltered=False
            if element[self.columnboard]==runningprj:
                isactive=True
            else:
                isactive=False

            #values = [element[self.columnboard],isfiltered,isactive,cycle,element[self.columnStatus],element[self.columncustomerprj],element[self.columnEffectiveHour],element[self.columnCost]]
            values=self.listofwork[cycle]
            attribute=[isfiltered,isactive]
            #dictofworks[cycle] = {k: v for k, v in zip(self.keys, values)}
            dictofworks[element[self.columnboard]] = values
            dictofworks[element[self.columnboard]+"attr"] = attribute
            cycle=cycle+1
        return self.listofwork,dictofworks