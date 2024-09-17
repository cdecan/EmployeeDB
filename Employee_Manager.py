from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5 import QtWidgets
from PyQt5.QtCore import *
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QApplication
import tkinter as tk
from tkinter.filedialog import askopenfilename
import shutil
from datetime import datetime
from datetime import timedelta
from collections import Counter
from dateutil.easter import *

import sys
from os import path, makedirs, startfile, system
from PyQt5.uic import loadUiType


FORM_CLASS,_=loadUiType(path.join(path.dirname('__file__'),"main.ui"))
FORM_CLASS2,_=loadUiType(path.join(path.dirname('__file__'),"AddEmployee.ui"))

import sqlite3

class Main(QMainWindow, FORM_CLASS):
    
    
    def __init__(self, parent=None):
        super(Main,self).__init__(parent)
        QMainWindow.__init__(self)
        
        self.currentEmployeeNumber = 0
        self.currentEmployeeName = "None"
        
        self.setupUi(self)
        self.Handle_Buttons()
        self.GET_EMPLOYEE_DATA()
        
        # Visuals
        self.setFixedSize(1045, 772)
        self.EmployeeSelector.setColumnWidth(6, 200)
        self.EmployeeSelector.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.EmployeeSelector.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.EmployeeSelector.setSortingEnabled(True)
        ###
        self.Month1.setCurrentPage(self.Month1.yearShown(), 1)
        self.Month2.setCurrentPage(self.Month1.yearShown(), 2)
        self.Month3.setCurrentPage(self.Month1.yearShown(), 3)
        self.Month4.setCurrentPage(self.Month1.yearShown(), 4)
        self.CALENDAR_DATE_CHANGED()
        ###
        self.EmployeeSelector.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.EmployeeSelector.customContextMenuRequested.connect(self.GenerateMenu)
        self.EmployeeSelector.viewport().installEventFilter(self)

        
        ###
        self.yes = QPixmap("resources/yes.png")
        self.no = QPixmap("resources/mp.png")
        self.monthButtons3.setIcon(QIcon('resources/right.png'))
        self.monthButtons1.setIcon(QIcon('resources/left.png'))
        self.current_month = datetime.now().month - 1
        self.current_year = datetime.now().year
        self.SetMonthsFullView(1)
                
        
        
    def Handle_Buttons(self):
        self.refreshButton.clicked.connect(self.GET_EMPLOYEE_DATA)
        self.VacationCalendar.selectionChanged.connect(self.CALENDAR_DATE_CHANGED)
        self.EmployeeSelector.doubleClicked.connect(self.DOUBLE_CLICKED_TABLE)
        self.mainTab.currentChanged.connect(self.ON_TAB_CHANGE)
        self.vacationTab.currentChanged.connect(self.CALENDAR_DATE_CHANGED)
        self.vacationSaveButton.clicked.connect(self.SAVE_VACATION)
        self.newEmployeeButton.clicked.connect(self.NEW_EMPLOYEE)
        self.monthButtons1.clicked.connect(lambda: self.SetMonthsFullView(0))
        self.monthButtons2.clicked.connect(lambda: self.SetMonthsFullView(1))
        self.monthButtons3.clicked.connect(lambda: self.SetMonthsFullView(2))
        self.VacationViewAllTab.currentChanged.connect(self.MONTH_TAB_CHANGED)
        self.ndupbutton.clicked.connect(lambda: self.UploadFile("new"))
        self.odupbutton.clicked.connect(lambda: self.UploadFile("old"))
        self.trainupbutton.clicked.connect(lambda: self.UploadFile("training"))
        self.ndviewbutton.clicked.connect(lambda: self.ViewFile("new"))
        self.odviewbutton.clicked.connect(lambda: self.ViewFile("old"))
        self.trainviewbutton.clicked.connect(lambda: self.ViewFile("training"))
        self.Month1.activated.connect(self.DOUBLE_CLICKED_DATE)
        self.Month2.activated.connect(self.DOUBLE_CLICKED_DATE)
        self.Month3.activated.connect(self.DOUBLE_CLICKED_DATE)
        self.Month4.activated.connect(self.DOUBLE_CLICKED_DATE)
        
        
        
    def DOUBLE_CLICKED_DATE(self, date):
        self.VacationCalendar.setSelectedDate(date)
        self.VacationViewAllTab.setCurrentIndex(0)
        
        
        
    def eventFilter(self, source, event):
        if(event.type() == QtCore.QEvent.MouseButtonPress and
           event.buttons() == QtCore.Qt.RightButton and
           source is self.EmployeeSelector.viewport()):
            item = self.EmployeeSelector.itemAt(event.pos())
            #print('Global Pos:', event.globalPos())
            if item is not None:
                #print('Table Item:', item.row(), item.column())
                self.menu = QMenu(self)
                cempNum = int(self.EmployeeSelector.item(item.row(), 0).text())
                dst1 = path.join(path.dirname(path.abspath(__file__)), "data\\" + str(cempNum) + "\\discipline\\new\\file.pdf")
                dst2 = path.join(path.dirname(path.abspath(__file__)), "data\\" + str(cempNum) + "\\discipline\\old\\file.pdf")
                dst3 = path.join(path.dirname(path.abspath(__file__)), "data\\" + str(cempNum) + "\\training\\file.pdf")
                if path.exists(dst1):
                    action1 = self.menu.addAction("Discipline - New")
                    action1.triggered.connect(lambda: self.ViewFile("new"))
                if path.exists(dst2):
                    action2 = self.menu.addAction("Discipline - Old")
                    action2.triggered.connect(lambda: self.ViewFile("old"))
                if path.exists(dst3):
                    action3 = self.menu.addAction("Training")
                    action3.triggered.connect(lambda: self.ViewFile("training"))
        return super().eventFilter(source, event)
    
    def GenerateMenu(self, pos):
        #print("pos======",pos)
        self.menu.exec_(self.EmployeeSelector.mapToGlobal(pos))   # +++


        
    def UploadFile(self, i):
        self.errorMessage.setText("")
        root = tk.Tk()
        root.withdraw()
        filename = askopenfilename(filetypes=[("PDF files", "*.pdf")])
        destination = "data/"
        destination += str(self.currentEmployeeNumber)
        
        
        if i == "new":
            destination += "/discipline/new/file.pdf"
        elif i == "old":
            destination += "/discipline/old/file.pdf"
        elif i =="training":
            destination += "/training/file.pdf"
        makedirs(path.dirname(destination), exist_ok=True)
        shutil.copy2(filename, destination)
        
        dst = path.join(path.dirname(path.abspath(__file__)),destination)
        if not path.exists(dst):
            self.errorMessage.setText("Error: upload failed")
            return
        
    def ViewFile(self, i):
        self.errorMessage.setText("")
        destination = "data/"
        destination += str(self.currentEmployeeNumber)
        
        if i == "new":
            destination += "/discipline/new"
        elif i == "old":
            destination += "/discipline/old"
        elif i =="training":
            destination += "/training"
            
        destination += "/file.pdf"
        
        dst = path.join(path.dirname(path.abspath(__file__)),destination)
        
        if not path.exists(dst):
            self.errorMessage.setText("Error: file not found")
            return
        
        startfile(dst)
        
    def NEW_EMPLOYEE(self):
        cc = CharacterCreator()
        widget.removeWidget(widget.currentWidget())
        widget.addWidget(cc)
        
        
    def SetupImages(self):
        
        
        dst1 = path.join(path.dirname(path.abspath(__file__)), "data\\" + str(self.currentEmployeeNumber) + "\\discipline\\new\\file.pdf")
        dst2 = path.join(path.dirname(path.abspath(__file__)), "data\\" + str(self.currentEmployeeNumber) + "\\discipline\\old\\file.pdf")
        dst3 = path.join(path.dirname(path.abspath(__file__)), "data\\" + str(self.currentEmployeeNumber) + "\\training\\file.pdf")
        i1 = self.ndimage
        i2 = self.odimage
        i3 = self.trainingimage
        
        if path.exists(dst1):
            i1.setPixmap(self.yes.scaled(i1.size(), QtCore.Qt.IgnoreAspectRatio))
        else:
            i1.setPixmap(self.no.scaled(i1.size(), QtCore.Qt.IgnoreAspectRatio))
        if path.exists(dst2):
            i2.setPixmap(self.yes.scaled(i2.size(), QtCore.Qt.IgnoreAspectRatio))
        else:
            i2.setPixmap(self.no.scaled(i2.size(), QtCore.Qt.IgnoreAspectRatio))
        if path.exists(dst3):
            i3.setPixmap(self.yes.scaled(i3.size(), QtCore.Qt.IgnoreAspectRatio))
        else:
            i3.setPixmap(self.no.scaled(i3.size(), QtCore.Qt.IgnoreAspectRatio))
        
        
    def ON_TAB_CHANGE(self, i):
        if i == 1:
            self.SetupImages()
        if i == 2:
            self.CALENDAR_DATE_CHANGED()
    
    def MONTH_TAB_CHANGED(self, i):
        if i == 1:
            self.CALENDAR_DATE_CHANGED()
        if i == 2:
            self.setCalendarFullView(1)
        
    def DOUBLE_CLICKED_TABLE(self, item):
        cRow = item.row()
        self.currentEmployeeNumber = int(self.EmployeeSelector.item(cRow, 0).text())
        self.currentEmployeeName = self.EmployeeSelector.item(cRow, 1).text()
        self.SET_EMPLOYEE()
    
    def GET_EMPLOYEE_DATA(self):
        db=sqlite3.connect('employeedb.db')
        cursor=db.cursor()
        
        command= '''SELECT * from EmployeeDB '''
        
        result = cursor.execute(command)
        
        self.EmployeeSelector.setRowCount(0)
        
        for row_number, row_data in enumerate(result):
            self.EmployeeSelector.insertRow(row_number)
            for column_number, data in enumerate(row_data):
                self.EmployeeSelector.setItem(row_number, column_number, QTableWidgetItem(str(data)))
                
        self.currentEmployeeNumber = int(self.EmployeeSelector.item(0, 0).text())
        self.currentEmployeeName = self.EmployeeSelector.item(0, 1).text()
        self.SET_EMPLOYEE()
    
    def SET_EMPLOYEE(self):
        self.currentEmployeeLabel.setText(self.currentEmployeeName)
        self.vacationEmployeeName.setText(self.currentEmployeeName)
        
    def CALENDAR_DATE_CHANGED(self):
        dateSelected = self.VacationCalendar.selectedDate().toPyDate()
        self.updateVacationChecklist(dateSelected)
        self.updateVacationDayList(dateSelected)
        self.updateCalendarColors(self.VacationCalendar)
        
        
    def getNameFromNumber(self, num):
        db = sqlite3.connect("employeedb.db")
        cursor = db.cursor()
        query = "SELECT Name FROM EmployeeDB WHERE EmployeeNumber = ?"
        result = cursor.execute(query, (num,)).fetchall()
        
        return result[0][0]
        
    def NumOfEmployees(self):
        db = sqlite3.connect("employeedb.db")
        cursor = db.cursor()
        query = "SELECT * FROM EmployeeDB"
        wowza = cursor.execute(query).fetchall()
        
        return len(wowza)
        
    def updateCalendarColors(self, vacationCalendar):
        
        employeeAmount = self.NumOfEmployees()
        
        tasks = ['Selected', 'Requested']
        
        db = sqlite3.connect("employeedb.db")
        cursor = db.cursor()
        query = "SELECT Date FROM VacationDates WHERE Task = ?"
        vacationResults = cursor.execute(query, (tasks[0],)).fetchall()
        
        
        red = QColor(255,0,0)
        redBrush = QBrush(red)
        green = QColor(0,255,0)
        greenBrush = QBrush(green)
        yellow = QColor(255,255,0)
        yellowBrush = QBrush(yellow)
        
        
        blue = QColor(0,0,255)
        blueBrush = QBrush(blue)
        
        green_format = QTextCharFormat()
        green_format.setBackground(greenBrush)
        green_format.setForeground(self.palette().color(QPalette.HighlightedText))
        
        red_format = QTextCharFormat()
        red_format.setBackground(redBrush)
        red_format.setForeground(self.palette().color(QPalette.HighlightedText))
        
        yellow_format = QTextCharFormat()
        yellow_format.setBackground(yellowBrush)
        yellow_format.setForeground(self.palette().color(QPalette.HighlightedText))
        
        blue_format = QTextCharFormat()
        blue_format.setBackground(blueBrush)
        blue_format.setForeground(self.palette().color(QPalette.HighlightedText))
                
        c = Counter(vacationResults)
        
        twentypercentdays = []
        
        for year in range(datetime.now().year, datetime.now().year + 6):
        
            xmas = datetime(year, 12, 24).date()
            nyears = datetime(year, 12, 31).date()
            eas = easter(year)
            bigDays = [xmas, nyears, eas]
            
            for bigday in bigDays:
            
                weekday = bigday.weekday()
                start_delta = timedelta(days=weekday)
                start_of_week = bigday - start_delta
                
                for day in range(7):
                    twentypercentdays.append(start_of_week + timedelta(days=day-1))
                    
                    
        for bigday in twentypercentdays:
            qtDate = QDate.fromString(str(bigday), "yyyy-MM-dd")
            vacationCalendar.setDateTextFormat(qtDate, blue_format)
        
        
        
        for date in vacationResults:
            dateBreakdown = date[0].split("-")
            qDate = QDate(int(dateBreakdown[0]), int(dateBreakdown[1]), int(dateBreakdown[2]))
            if date in twentypercentdays and c[date] / employeeAmount >= 0.2:
                vacationCalendar.setDateTextFormat(qDate, red_format)
            elif c[date] / employeeAmount >= 0.1:
                vacationCalendar.setDateTextFormat(qDate, red_format)
            else:
                vacationCalendar.setDateTextFormat(qDate, green_format)
        
        vacationResults = cursor.execute(query, (tasks[1],)).fetchall()
        
        for date in vacationResults:
            dateBreakdown = date[0].split("-")
            qDate = QDate(int(dateBreakdown[0]), int(dateBreakdown[1]), int(dateBreakdown[2]))
            vacationCalendar.setDateTextFormat(qDate, yellow_format)
        
    
    def updateVacationDayList(self, date):
        self.vacationDayList.clear()
        db = sqlite3.connect("employeedb.db")
        cursor = db.cursor()
        query = "SELECT Task, EmployeeNumber FROM VacationDates WHERE Date = ?"
        vacationResults = cursor.execute(query, (date,)).fetchall()
        
        finalString = ""
        
        for x in range(len(vacationResults)):
            finalString += self.getNameFromNumber(vacationResults[x][1])
            finalString += ": "
            finalString += vacationResults[x][0]
            finalString += "\n"
        
            
        
        self.vacationDayList.setText(str(finalString))
        
        
    def updateVacationChecklist(self, date):
        self.vacationListWidget.clear()
        
        tasks = ['Selected', 'Requested']
        
        db = sqlite3.connect("employeedb.db")
        cursor = db.cursor()
        
        query = "SELECT Task FROM VacationDates WHERE Date = ? AND EmployeeNumber = ?"
        results = cursor.execute(query, (date, self.currentEmployeeNumber)).fetchall()
        
        for task in tasks:
            item = QListWidgetItem(task)
            checked = False
            for result in results:
                if task in result[0]:
                    checked = True
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            if(checked):
                item.setCheckState(QtCore.Qt.Checked)
            else:
                item.setCheckState(QtCore.Qt.Unchecked)
            self.vacationListWidget.addItem(item)
            
    def SAVE_VACATION(self):
        db = sqlite3.connect("employeedb.db")
        cursor = db.cursor()
        date = self.VacationCalendar.selectedDate().toPyDate()
        for i in range(self.vacationListWidget.count()):
            item = self.vacationListWidget.item(i)
            task = item.text()
            if item.checkState() == QtCore.Qt.Checked:
                query = "INSERT INTO VacationDates (EmployeeNumber, Task, Date) VALUES (?, ?, ?)"
            else:
                query = "DELETE FROM VacationDates WHERE EmployeeNumber=? AND Task=? AND Date=?"
            
            cursor.execute(query, (self.currentEmployeeNumber, task, date))
        db.commit()
    
    def SetMonthsFullView(self, i):
        
        iterList = [1,2,3,4,5,6,7,8,9,10,11,12]
        
        if i == 0:
            if self.current_month - 4 < 0:
                self.current_year -= 1
            self.current_month = (self.current_month - 4)%12
        elif i == 1:
            self.current_month = datetime.now().month - 1
            self.current_year = datetime.now().year
        elif i == 2:
            if self.current_month + 4 > 11:
                self.current_year += 1
            self.current_month = (self.current_month + 4)%12
            
        months = [iterList[self.current_month], iterList[(self.current_month+1)%12], iterList[(self.current_month+2)%12], iterList[(self.current_month+3)%12]]
        years = [self.current_year, self.current_year, self.current_year, self.current_year]
        
        for j in range(0, len(months)):
            if months[j] < months[0]:
                years[j] += 1
        
        self.Month1.setCurrentPage(years[0], months[0])
        self.Month2.setCurrentPage(years[1], months[1])
        self.Month3.setCurrentPage(years[2], months[2])
        self.Month4.setCurrentPage(years[3], months[3])
        self.updateCalendarColors(self.Month1)
        self.updateCalendarColors(self.Month2)
        self.updateCalendarColors(self.Month3)
        self.updateCalendarColors(self.Month4)
    
class CharacterCreator(QMainWindow, FORM_CLASS2):
    def __init__(self, parent=None):
        super(CharacterCreator,self).__init__(parent)
        QMainWindow.__init__(self)
        
        self.currentEmployeeNumber = 0
        self.currentEmployeeName = "None"
        
        self.setupUi(self)
        self.Handle_Buttons()
        
        self.Make_Events()
        
        # Visuals
        self.setFixedSize(1045, 772)
        ###
        
    def Make_Events(self):
        self.EmployeeNumberField.installEventFilter(self)
        self.EmployeeNameField.installEventFilter(self)
        self.SeniorityField.installEventFilter(self)
        self.PhoneField.installEventFilter(self)
        self.CellField.installEventFilter(self)
        self.ResidenceField.installEventFilter(self)
        self.ContactField.installEventFilter(self)
        
    def Handle_Buttons(self):
        self.addEmployeeButton.clicked.connect(self.ADD_EMPLOYEE)
        self.pushButton.clicked.connect(self.RETURN)
        
    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress:
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                return True
            if event.key() == (Qt.Key_Tab):
                self.focusNextPrevChild(True)
                return True
        return super().eventFilter(obj, event)
        
    def ADD_EMPLOYEE(self):
        
        self.errorMessage.setStyleSheet('font: 12pt "MS Shell Dlg 2";color: rgb(255, 0, 0);')
        self.errorMessage.setText("")
        
        legalChars = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890-(), "
        employeeNumber = self.EmployeeNumberField.toPlainText()
        employeeName = self.EmployeeNameField.toPlainText()
        employeeSeniority = self.SeniorityField.toPlainText()
        employeePhone = self.PhoneField.toPlainText()
        employeeCell = self.CellField.toPlainText()
        employeeRes = self.ResidenceField.toPlainText()
        employeeEcontact = self.ContactField.toPlainText()
        # CHECK LEGALITY
        
        if not employeeNumber or not employeeName or not employeeSeniority or not employeeRes or not employeeEcontact:
            self.errorMessage.setText("Error: Missing Required Form")
            return
        
        for letter in employeeNumber:
            if letter not in legalChars:
                self.errorMessage.setText("Error: Employee Number Has Invalid Characters")
                return
        for letter in employeeName:
            if letter not in legalChars:
                self.errorMessage.setText("Error: Employee Name Has Invalid Characters")
                return
        for letter in employeeEcontact:
            if letter not in legalChars:
                self.errorMessage.setText("Error: Emergency Contact Has Invalid Characters")
                return
        for letter in employeeSeniority:
            if letter not in legalChars:
                self.errorMessage.setText("Error: Seniority Has Invalid Characters")
                return
        for letter in employeePhone:
            if letter not in legalChars:
                self.errorMessage.setText("Error: Phone Number Has Invalid Characters")
                return
        for letter in employeeCell:
            if letter not in legalChars:
                self.errorMessage.setText("Error: Cell Number Has Invalid Characters")
                return
        for letter in employeeRes:
            if letter not in legalChars:
                self.errorMessage.setText("Error: Residence Has Invalid Characters")
                return
        db = sqlite3.connect("employeedb.db")
        cursor = db.cursor()
        query = "INSERT INTO VacationDates (EmployeeNumber, Name, Seniority, PhoneNumber, CellNumber, Residence, EmergencyContact) VALUES (?, ?, ?, ?, ?, ?, ?)"
        cursor.execute(query, (employeeNumber, employeeName, employeeSeniority, employeePhone, employeeCell, employeeRes, employeeEcontact))
        db.commit()
        self.errorMessage.setStyleSheet('font: 12pt "MS Shell Dlg 2";color: rgb(0, 255, 0);')
        self.errorMessage.setText("Employee Added")
    
    def RETURN(self):
        home = Main()
        widget.removeWidget(widget.currentWidget())
        widget.addWidget(home)
    
    
def main():
    pass
    
    
if __name__ == '__main__':
    app=QApplication(sys.argv)
    widget = QtWidgets.QStackedWidget()
    widget.setFixedSize(1045, 772)
    widget.setWindowTitle("System")
    
    mainWindow = Main()
    #ccWindow = CharacterCreator()
    widget.addWidget(mainWindow)
    #widget.addWidget(ccWindow)
    widget.show()
    app.exec_()
    
    main()
    