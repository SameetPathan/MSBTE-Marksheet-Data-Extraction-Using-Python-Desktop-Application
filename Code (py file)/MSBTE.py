
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox
import requests
import re
from bs4 import BeautifulSoup
import xlsxwriter


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(781, 523)
        MainWindow.setToolTip("")
        MainWindow.setToolTipDuration(-2)
        MainWindow.setAutoFillBackground(False)
        MainWindow.setStyleSheet("background-color: rgb(0, 0, 0);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(90, 360, 111, 41))
        self.pushButton.setStyleSheet("background-color: rgb(124, 124, 124);")
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.Semester_1)

        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(560, 390, 101, 41))
        self.pushButton_2.setToolTipDuration(2)
        self.pushButton_2.setStyleSheet("background-color: rgb(124, 124, 124);")
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.endall)

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(0, 10, 781, 101))
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setStyleSheet("background-color: rgb(136, 136, 136);\n"
"background-color: rgb(88, 88, 88);")
        self.label.setObjectName("label")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(360, 130, 291, 41))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.textEdit.setFont(font)
        self.textEdit.setToolTip("")
        self.textEdit.setStyleSheet("background-color: rgb(0, 86, 0);")
        self.textEdit.setObjectName("textEdit")
        self.textEdit_2 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_2.setGeometry(QtCore.QRect(360, 200, 291, 41))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.textEdit_2.setFont(font)
        self.textEdit_2.setStyleSheet("background-color: rgb(0, 86, 0);")
        self.textEdit_2.setObjectName("textEdit_2")
        self.textEdit_3 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_3.setGeometry(QtCore.QRect(360, 270, 291, 41))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.textEdit_3.setFont(font)
        self.textEdit_3.setStyleSheet("background-color: rgb(0, 86, 0);")
        self.textEdit_3.setObjectName("textEdit_3")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(520, 480, 251, 31))
        self.label_2.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(90, 150, 231, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(90, 210, 231, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_4.setFont(font)
        self.label_4.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(90, 280, 231, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_5.setFont(font)
        self.label_5.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_5.setObjectName("label_5")

        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(250, 360, 111, 41))
        self.pushButton_3.setStyleSheet("background-color: rgb(124, 124, 124);")
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.Semester_2)

        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(90, 430, 111, 41))
        self.pushButton_4.setStyleSheet("background-color: rgb(124, 124, 124);")
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.clicked.connect(self.Semester_4)

        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QtCore.QRect(250, 430, 111, 41))
        self.pushButton_5.setStyleSheet("background-color: rgb(124, 124, 124);")
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_5.clicked.connect(self.Semester_5)


        self.pushButton_6 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_6.setGeometry(QtCore.QRect(400, 360, 111, 41))
        self.pushButton_6.setStyleSheet("background-color: rgb(124, 124, 124);")
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_6.clicked.connect(self.Semester_3)


        self.pushButton_7 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_7.setGeometry(QtCore.QRect(400, 430, 111, 41))
        self.pushButton_7.setStyleSheet("background-color: rgb(124, 124, 124);")
        self.pushButton_7.setObjectName("pushButton_7")
        self.pushButton_7.clicked.connect(self.showEr)

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "SEMESTER 1"))
        self.pushButton_2.setText(_translate("MainWindow", "EXIT/CANCEL"))
        self.label.setText(_translate("MainWindow", "<html><head/><body><p align=\"center\"><span style=\" font-weight:400;\">MSBTE STUDENT MARKSHEET DATA EXTRACTION</span></p><p align=\"center\"><span style=\" font-weight:400;\">(COMPUTER ENGINEERING DEPARTMENT)</span></p></body></html>"))
        self.textEdit.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:14pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:8.25pt;\"><br /></p></body></html>"))
        self.textEdit_2.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:14pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:8.25pt;\"><br /></p></body></html>"))
        self.textEdit_3.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:14pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:8.25pt;\"><br /></p></body></html>"))
        self.label_2.setText(_translate("MainWindow", "PROJECT BY SAMEET PATHAN AND RUTUJA PATIL"))
        self.label_3.setText(_translate("MainWindow", "STARTING SEAT NUMBER"))
        self.label_4.setText(_translate("MainWindow", "LAST SEAT NUMBER"))
        self.label_5.setText(_translate("MainWindow", "FILE NAME TO BE SAVED"))
        self.pushButton_3.setText(_translate("MainWindow", "SEMESTER 2"))
        self.pushButton_4.setText(_translate("MainWindow", "SEMESTER 4"))
        self.pushButton_5.setText(_translate("MainWindow", "SEMESTER 5"))
        self.pushButton_6.setText(_translate("MainWindow", "SEMESTER 3"))
        self.pushButton_7.setText(_translate("MainWindow", "SEMESTER 6"))

        self.msg = QMessageBox()
        self.msg.setWindowTitle("Process")
        self.msg.setText("Extraction Completed")

        self.msg1 = QMessageBox()
        self.msg1.setWindowTitle("Error")
        self.msg1.setText("Please check your Internet connection" + "\n" + "And if you have enter data properly ")

        self.msg2 = QMessageBox()
        self.msg2.setWindowTitle("Starting Extraction")
        self.msg2.setText("Process Running" + "\n" + "Please wait....")

    def endall(self):
        exit()

    def showEr(self):
        self.msg12 = QMessageBox()
        self.msg12.setWindowTitle("Error")
        self.msg12.setText("The code hasnt been updated for this Semester")
        self.msg12.show()


    def Semester_4(self):
        try:
            start1 = self.textEdit.toPlainText()
            end1 = self.textEdit_2.toPlainText()
            file = self.textEdit_3.toPlainText()
            start = int(start1)
            end = int(end1)

            # start = int(input("Enter Starting seat Number:"))
            # end = int(input("Enter Ending seat Number:"))
            # file = input("Enter Excel File(If already exist enter same name or create with new name):")
            s = 2;
            workbook = xlsxwriter.Workbook(file + '.xlsx')
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold': True})
            worksheet.write('A1', 'NAME', bold)
            worksheet.write('B1', 'Enrolment', bold)
            worksheet.write('C1', 'Examination', bold)
            worksheet.write('D1', 'Seat Number', bold)
            worksheet.write('E1', 'Semester', bold)
            worksheet.write('F1', 'Course', bold)
            worksheet.write('G1', 'Subject 1 ', bold)
            worksheet.write('H1', 'Marks 1', bold)
            worksheet.write('I1', 'Subject 2', bold)
            worksheet.write('J1', 'Marks 2', bold)
            worksheet.write('K1', 'Subject 3', bold)
            worksheet.write('L1', 'Marks 3', bold)
            worksheet.write('M1', 'Subject 4', bold)
            worksheet.write('N1', 'Marks 4', bold)
            worksheet.write('O1', 'Subject 5', bold)
            worksheet.write('P1', 'Marks 5', bold)
            worksheet.write('Q1', 'Total Marks', bold)
            worksheet.write('R1', 'Obtained Marks', bold)
            worksheet.write('S1', 'Percentage', bold)
            # start progress

            self.msg2.show()

            for n in range(start, end + 1):

                u = str(n)

                url_final = "https://msbte.org.in/SOLCOVD2020RESBTELIVE/SNONFNL20246RESLIVE/SeatNumber/18/" + u + "Marksheet.html"
                res = ""
                try:
                    res = requests.get(url_final)
                except requests.ConnectionError as exception:
                    continue
                    #print("error")

                html_page = res.content
                soup = BeautifulSoup(html_page, 'html.parser')
                text = soup.find_all(text=True)
                output = ''
                blacklist = ['[document]', 'noscript', 'header', 'html', 'meta', 'head', 'title', 'input', 'script',
                             'style']
                for t in text:
                    if t.parent.name not in blacklist:
                        output += '{}'.format(t) + "\n"
                output1 = re.sub(
                    "|Maharashtra State Board of Technical Education|MR. / MS.|Statement of Marks|TITLE OF S|ENROLLMENT NO.|EXAMNINATION|SEAT NO.|COURSE"
                    "|E & OE|Url:-http://www.msbte.org.in|Result Declared On 07/01/2020|Ref:Formerly known as The board of Technical Examinations Maharashtra State of Technical Education Act 1997|Mah XXXVIII of 1997"
                    "|and Maharashtra Government Gazette Notification Section IV-B issued on march 31,1999|-172/16/200/20-04:07:2007 12:00:28"
                    "|Progressive Assessment|End Semester Exam|Distinction|Result Withheld Due To Pending Lower Semester|WFLS|Failure Marks|Additional Practical|"
                    "|Failure But Allowed To Keep Term|Pending Lower Year|Condoned Marks|Aggregate|Condoned|Lower Semester Pending|Industrial Training|Practical Test Marks"
                    "|Optional|Sessional|Practical|Result Withheld Due to Pending Lower Year|Exemption|Project Work|Theory Test Marks|Percentage of Marks|Absent|Team Work|Theory|"
                    "|ABBREVATION DETAILS|Class awarded for Diploma is based on aggregate marks obtained in pre-final & final semester|Allowed to Keep Term"
                    "|Carry Forward Marks|Candidate is eligible for admission to V/VII Semester only if he/she is fully passed in I & II /III & IV semesters & availed benefit of A.T.K.T/PASS at III & IV /V & VI semesters taken together respectively"
                    "|  number of failure subjects in I & II semesters taken  together|"
                    "|Eligibility for III semester is based on total|This certificate of marks is issued as per prevaling rules and regulations of MSBTE at the time of this exam|"
                    "|Report Discrepancy in this certificate to Head of the institution|INSTRUCTIONS|MAHARASHTRA STATE BOARD OF TECHNICAL EDUCATION|SECRETARY|"
                    "|This Marksheet is Downloaded from Internet|RESULT WITH|MARKS|TOTAL MAX|DATE :|OBTAINED|TOTAL CREDIT|CREDITS|MAX|OBT|MAX|  HEAD|DIST|C#|Oral|OR|A.T.K.T|"
                    "|PLY|ESE|@|"
                    "|TOTAL MARKS|", "", output)

                outputfinal = output1.replace("\n", " ").lstrip().replace("               ", "\n#").replace(
                    "      ",
                    "\n#").replace(
                    "   ", "\n#")
                # print(outputfinal)

                # name etracting
                i = 0
                name = ''
                ans = ''
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    name = name + ans
                    i = i + 1
                #print("Name : ", name)

                # enrolment extracting
                occurrence = 1
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                enrolment = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    enrolment = enrolment + ans
                    i = i + 1
                #print("Enrolment : ", enrolment)
                # examination
                occurrence = 2
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                examination = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    examination = examination + ans
                    i = i + 1
                #print("Examination : ", examination)
                # seat
                occurrence = 3
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                seatno = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    seatno = seatno + ans
                    i = i + 1
                #print("Seat Number : ", seatno)
                # semester
                occurrence = 4
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                semester = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    semester = semester + ans
                    i = i + 1
                #print("Semester : ", semester)
                # course
                occurrence = 6
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                course = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    course = course + ans
                    i = i + 1
                #print("Course : ", course)

                # SUbject 1
                occurrence = 10
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub1 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub1 = sub1 + ans
                    i = i + 1
                sub1.replace("TH", "")

                #print("Subject 1 : ", sub1.replace("TH", ""))

                # SUbject 1 marks
                occurrence = 11
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m1 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m1 = m1 + ans
                    i = i + 1
                m1t = m1[1] + m1[2] + m1[3]
                ma1 = m1[5] + m1[6] + m1[7]
                #print("Total Marks : ", m1t)
                #print("Obtained Marks : ", ma1)
                # SUbject 2
                occurrence = 16
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub2 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub2 = sub2 + ans
                    i = i + 1
                sub2.replace("TH", "")
                #print("Subject 2 : ", sub2.replace("TH", ""))
                # SUbject 2 marks
                occurrence = 17
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m2 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m2 = m2 + ans
                    i = i + 1
                m2t = m2[1] + m2[2] + m2[3]
                ma2 = m2[5] + m2[6] + m2[7]
                #print("Total Marks : ", m2t)
                #print("Obtained Marks : ", ma2)
                # SUbject 3
                occurrence = 22
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub3 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub3 = sub3 + ans
                    i = i + 1
                sub3.replace("TH", "")
                #print("Subject 3 : ", sub3.replace("TH", ""))
                # SUbject 3 marks
                occurrence = 23
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m3 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m3 = m3 + ans
                    i = i + 1
                m3t = m3[1] + m3[2] + m3[3]
                ma3 = m3[5] + m3[6] + m3[7]
                #print("Total Marks : ", m3t)
                #print("Obtained Marks : ", ma3)
                # SUbject 4
                occurrence = 28
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub4 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub4 = sub4 + ans
                    i = i + 1
                sub4.replace("TH", "")
                #print("Subject 4 : ", sub4.replace("TH", ""))
                # SUbject 4 marks
                occurrence = 29
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m4 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m4 = m4 + ans
                    i = i + 1
                m4t = m4[1] + m4[2] + m4[3]
                ma4 = m4[5] + m4[6] + m4[7]
                #print("Total Marks : ", m4t)
                #print("Obtained Marks : ", ma4)
                # SUbject 5
                occurrence = 34
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub5 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub5 = sub5 + ans
                    i = i + 1
                sub5.replace("TH", "")
                #print("Subject 5 : ", sub5.replace("TH", ""))
                # SUbject 5 marks
                occurrence = 35
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m5 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m5 = m5 + ans
                    i = i + 1
                m5t = m5[1] + m5[2] + m5[3]
                ma5 = m5[5] + m5[6] + m5[7]
                #print("Total Marks : ", m5t)
                #print("Obtained Marks : ", ma5)

                # Total marks / marks obtained/ percentage
                occurrence = 41
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                total = ""
                percent = ""
                obtain = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    total = total + ans
                    i = i + 1
                total.lstrip()

                percent = total[6:11]
                obtain = total[11:16]
                totalm = total[0:6]
                #print("total : ", total[0:6])
                #print("obtain : ", obtain)
                #print("percent : ", percent)

                # save data

                # The workbook object is then used to add new
                # worksheet via the add_worksheet() method.

                # Use the worksheet object to write
                # data via the write() method.
                index = str(s)
                # style =('font: bold 1, color red;')
                worksheet.write('A' + index, name)
                worksheet.write('B' + index, enrolment)
                worksheet.write('C' + index, examination)
                worksheet.write('D' + index, seatno)
                worksheet.write('E' + index, semester)
                worksheet.write('F' + index, course)
                worksheet.write('G' + index, sub1)
                worksheet.write('H' + index, ma1)
                worksheet.write('I' + index, sub2)
                worksheet.write('J' + index, ma2)
                worksheet.write('K' + index, sub3)
                worksheet.write('L' + index, ma3)
                worksheet.write('M' + index, sub4)
                worksheet.write('N' + index, ma4)
                worksheet.write('O' + index, sub5)
                worksheet.write('P' + index, ma5)
                worksheet.write('Q' + index, totalm)
                worksheet.write('R' + index, obtain)
                worksheet.write('S' + index, percent)
                s = s + 1
                # close the Excel file
            workbook.close()
            self.msg2.close()
            self.msg.show()
        except:

            self.msg1.show()
            self.msg2.close()

    def Semester_5(self):
        try:
            start1 = self.textEdit.toPlainText()
            end1 = self.textEdit_2.toPlainText()
            file = self.textEdit_3.toPlainText()
            start = int(start1)
            end = int(end1)

            # start = int(input("Enter Starting seat Number:"))
            # end = int(input("Enter Ending seat Number:"))
            # file = input("Enter Excel File(If already exist enter same name or create with new name):")
            s = 2;
            workbook = xlsxwriter.Workbook(file + '.xlsx')
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold': True})
            worksheet.write('A1', 'NAME', bold)
            worksheet.write('B1', 'Enrolment', bold)
            worksheet.write('C1', 'Examination', bold)
            worksheet.write('D1', 'Seat Number', bold)
            worksheet.write('E1', 'Semester', bold)
            worksheet.write('F1', 'Course', bold)
            worksheet.write('G1', 'Subject 1 ', bold)
            worksheet.write('H1', 'Marks 1', bold)
            worksheet.write('I1', 'Subject 2', bold)
            worksheet.write('J1', 'Marks 2', bold)
            worksheet.write('K1', 'Subject 3', bold)
            worksheet.write('L1', 'Marks 3', bold)
            worksheet.write('M1', 'Subject 4', bold)
            worksheet.write('N1', 'Marks 4', bold)
            worksheet.write('O1', 'Subject 5', bold)
            worksheet.write('P1', 'Marks 5', bold)
            worksheet.write('Q1', 'Subject 6', bold)
            worksheet.write('R1', 'Marks 6', bold)
            worksheet.write('S1', 'Subject 7', bold)
            worksheet.write('T1', 'Marks 7', bold)
            worksheet.write('U1', 'Total Marks', bold)
            worksheet.write('V1', 'Obtained Marks', bold)
            worksheet.write('W1', 'Percentage', bold)
            # start progress

            self.msg2.show()
            for n in range(start, end + 1):

                u = str(n)
                # you need to change this
                url_final = "https://msbte.org.in/SHTFNL20BTERESLIVE/SHTFNL20BTERESLIVE/SeatNumber/18/" + u + "Marksheet.html"

                try:
                    res = requests.get(url_final)
                except requests.ConnectionError as exception:
                    continue
                    # print("error")

                html_page = res.content
                soup = BeautifulSoup(html_page, 'html.parser')
                text = soup.find_all(text=True)
                output = ''
                blacklist = ['[document]', 'noscript', 'header', 'html', 'meta', 'head', 'title', 'input', 'script',
                             'style']
                for t in text:
                    if t.parent.name not in blacklist:
                        output += '{}'.format(t) + "\n"
                output1 = re.sub(
                    "|Maharashtra State Board of Technical Education|MR. / MS.|Statement of Marks|TITLE OF S|ENROLLMENT NO.|EXAMNINATION|SEAT NO.|COURSE"
                    "|E & OE|Url:-http://www.msbte.org.in|Result Declared On 07/01/2020|Ref:Formerly known as The board of Technical Examinations Maharashtra State of Technical Education Act 1997|Mah XXXVIII of 1997"
                    "|and Maharashtra Government Gazette Notification Section IV-B issued on march 31,1999|-172/16/200/20-04:07:2007 12:00:28"
                    "|Progressive Assessment|End Semester Exam|Distinction|Result Withheld Due To Pending Lower Semester|WFLS|Failure Marks|Additional Practical|"
                    "|Failure But Allowed To Keep Term|Pending Lower Year|Condoned Marks|Aggregate|Condoned|Lower Semester Pending|Industrial Training|Practical Test Marks"
                    "|Optional|Sessional|Practical|Result Withheld Due to Pending Lower Year|Exemption|Project Work|Theory Test Marks|Percentage of Marks|Absent|Team Work|Theory|"
                    "|ABBREVATION DETAILS|Class awarded for Diploma is based on aggregate marks obtained in pre-final & final semester|Allowed to Keep Term"
                    "|Carry Forward Marks|Candidate is eligible for admission to V/VII Semester only if he/she is fully passed in I & II /III & IV semesters & availed benefit of A.T.K.T/PASS at III & IV /V & VI semesters taken together respectively"
                    "|  number of failure subjects in I & II semesters taken  together|"
                    "|Eligibility for III semester is based on total|This certificate of marks is issued as per prevaling rules and regulations of MSBTE at the time of this exam|"
                    "|Report Discrepancy in this certificate to Head of the institution|INSTRUCTIONS|MAHARASHTRA STATE BOARD OF TECHNICAL EDUCATION|SECRETARY|"
                    "|This Marksheet is Downloaded from Internet|RESULT WITH|MARKS|TOTAL MAX|DATE :|OBTAINED|TOTAL CREDIT|CREDITS|MAX|MIN|OBT|MAX|  HEAD|DIST|C#|Oral|OR|A.T.K.T|"
                    "|PLY|ESE|@|"
                    "|TOTAL MARKS|", "", output)

                outputfinal = output1.replace("\n", " ").lstrip().replace("               ", "\n#").replace("      ",
                                                                                                            "\n#").replace(
                    "   ", "\n#")
                # print(outputfinal)

                # name etracting
                i = 0
                name = ''
                ans = ''
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    name = name + ans
                    i = i + 1
                # print("Name : ", name)

                # enrolment extracting
                occurrence = 1
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                enrolment = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    enrolment = enrolment + ans
                    i = i + 1
                # print("Enrolment : ", enrolment)
                # examination
                occurrence = 2
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                examination = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    examination = examination + ans
                    i = i + 1
                # print("Examination : ", examination)
                # seat
                occurrence = 3
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                seatno = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    seatno = seatno + ans
                    i = i + 1
                # print("Seat Number : ", seatno)
                # semester
                occurrence = 4
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                semester = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    semester = semester + ans
                    i = i + 1
                # print("Semester : ", semester)
                # course
                occurrence = 6
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                course = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    course = course + ans
                    i = i + 1
                # print("Course : ", course)

                # SUbject 1
                occurrence = 10
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub1 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub1 = sub1 + ans
                    i = i + 1
                sub1.replace("TH", "")
                # print("Subject 1 : ", sub1.replace("TH", ""))

                # SUbject 1 marks
                occurrence = 11
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m1 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m1 = m1 + ans
                    i = i + 1
                m1t = m1[1] + m1[2] + m1[3]
                ma1 = m1[5] + m1[6] + m1[7]
                # print("Total Marks : ", m1t)
                # print("Obtained Marks : ", ma1)
                # SUbject 2
                occurrence = 13
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub2 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub2 = sub2 + ans
                    i = i + 1
                sub2.replace("TH", "")
                # print("Subject 2 : ", sub2.replace("TH", ""))
                # SUbject 2 marks
                occurrence = 14
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m2 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m2 = m2 + ans
                    i = i + 1
                m2t = m2[1] + m2[2] + m2[3]
                ma2 = m2[5] + m2[6] + m2[7]
                # print("Total Marks : ", m2t)
                # print("Obtained Marks : ", ma2)
                # SUbject 3
                occurrence = 19
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub3 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub3 = sub3 + ans
                    i = i + 1
                sub3.replace("TH", "")
                # print("Subject 3 : ", sub3.replace("TH", ""))
                # SUbject 3 marks
                occurrence = 20
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m3 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m3 = m3 + ans
                    i = i + 1
                m3t = m3[1] + m3[2] + m3[3]
                ma3 = m3[5] + m3[6] + m3[7]
                # print("Total Marks : ", m3t)
                # print("Obtained Marks : ", ma3)
                # SUbject 4
                occurrence = 25
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub4 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub4 = sub4 + ans
                    i = i + 1
                sub4.replace("TH", "")
                # print("Subject 4 : ", sub4.replace("TH", ""))
                # SUbject 4 marks
                occurrence = 26
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m4 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m4 = m4 + ans
                    i = i + 1
                m4t = m4[1] + m4[2] + m4[3]
                ma4 = m4[5] + m4[6] + m4[7]
                # print("Total Marks : ", mt)
                # print("Obtained Marks : ", m4)
                # SUbject 5
                occurrence = 31
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub5 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub5 = sub5 + ans
                    i = i + 1
                sub5.replace("TH", "")
                # print("Subject 5 : ", sub5.replace("TH", ""))
                # SUbject 5 marks
                occurrence = 32
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m5 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m5 = m5 + ans
                    i = i + 1
                m5t = m5[1] + m5[2] + m5[3]
                ma5 = m5[5] + m5[6] + m5[7]
                # print("Total Marks : ", mt)
                # print("Obtained Marks : ", m5)
                # SUbject 6
                occurrence = 37
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub6 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub6 = sub6 + ans
                    i = i + 1
                sub6.replace("PR", "")
                # print("Subject 6 : ", sub6.replace("PR", ""))
                # SUbject 6 marks
                occurrence = 38
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m6 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m6 = m6 + ans
                    i = i + 1
                m6t = m6[1] + m6[2] + m6[3]
                ma6 = m6[5] + m6[6] + m6[7]
                # print("Total Marks : ", mt)
                # print("Obtained Marks : ", m6)
                # SUbject 7
                occurrence = 40
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub7 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub7 = sub7 + ans
                    i = i + 1
                sub7.replace("PR", "")
                # print("Subject 7 : ", sub7.replace("PR", ""))
                # SUbject 7 marks
                occurrence = 41
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m7 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m7 = m7 + ans
                    i = i + 1
                m7t = m7[1] + m7[2] + m7[3]
                ma7 = m7[5] + m7[6] + m7[7]
                # print("Total Marks : ", mt)
                # print("Obtained Marks : ", m7)

                # Total marks / marks obtained/ percentage
                occurrence = 47
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                out = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    out = out + ans
                    i = i + 1
                sample = out.lstrip().replace(" ", "\n#")
                # print("sample:",sample)
                # total
                i = 0
                total = ''
                ans = ''
                while sample[i] != "#":
                    ans = sample[i]
                    total = total + ans
                    i = i + 1
                # print("Total Subject Marks : ", total)
                # Marks obtained
                occurrence = 2
                val = -1
                for i in range(0, occurrence):
                    val = sample.find("#", val + 1)
                i = val + 1
                obtain = ""
                ans = ""
                while sample[i] != "#":
                    ans = sample[i]
                    obtain = obtain + ans
                    i = i + 1
                # print("Total Subject Marks Obtained : ", obtain)
                # Percentage
                occurrence = 1
                val = -1
                for i in range(0, occurrence):
                    val = sample.find("#", val + 1)
                i = val + 1
                percent = ""
                ans = ""
                while sample[i] != "#":
                    ans = sample[i]
                    percent = percent + ans
                    i = i + 1
                # print("Percentage : ", percent)

                # save data

                # The workbook object is then used to add new
                # worksheet via the add_worksheet() method.

                # Use the worksheet object to write
                # data via the write() method.
                index = str(s)
                # style =('font: bold 1, color red;')
                worksheet.write('A' + index, name)
                worksheet.write('B' + index, enrolment)
                worksheet.write('C' + index, examination)
                worksheet.write('D' + index, seatno)
                worksheet.write('E' + index, semester)
                worksheet.write('F' + index, course)
                worksheet.write('G' + index, sub1)
                worksheet.write('H' + index, ma1)
                worksheet.write('I' + index, sub2)
                worksheet.write('J' + index, ma2)
                worksheet.write('K' + index, sub3)
                worksheet.write('L' + index, ma3)
                worksheet.write('M' + index, sub4)
                worksheet.write('N' + index, ma4)
                worksheet.write('O' + index, sub5)
                worksheet.write('P' + index, ma5)
                worksheet.write('Q' + index, sub6)
                worksheet.write('R' + index, ma6)
                worksheet.write('S' + index, sub7)
                worksheet.write('T' + index, ma7)
                worksheet.write('U' + index, total)
                worksheet.write('V' + index, obtain)
                worksheet.write('W' + index, percent)
                s = s + 1
                # close the Excel file
            workbook.close()
            self.msg2.close()
            self.msg.show()
        except:
            self.msg1.show()
            self.msg2.close()

    def Semester_1(self):
        try:
            start1 = self.textEdit.toPlainText()
            end1 = self.textEdit_2.toPlainText()
            file = self.textEdit_3.toPlainText()
            start = int(start1)
            end = int(end1)

            # start = int(input("Enter Starting seat Number:"))
            # end = int(input("Enter Ending seat Number:"))
            # file = input("Enter Excel File(If already exist enter same name or create with new name):")
            s = 2;
            workbook = xlsxwriter.Workbook(file + '.xlsx')
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold': True})
            worksheet.write('A1', 'NAME', bold)
            worksheet.write('B1', 'Enrolment', bold)
            worksheet.write('C1', 'Examination', bold)
            worksheet.write('D1', 'Seat Number', bold)
            worksheet.write('E1', 'Semester', bold)
            worksheet.write('F1', 'Course', bold)
            worksheet.write('G1', 'Subject 1 ', bold)
            worksheet.write('H1', 'Marks 1', bold)
            worksheet.write('I1', 'Subject 2', bold)
            worksheet.write('J1', 'Marks 2', bold)
            worksheet.write('K1', 'Subject 3', bold)
            worksheet.write('L1', 'Marks 3', bold)
            worksheet.write('M1', 'Subject 4', bold)
            worksheet.write('N1', 'Marks 4', bold)
            worksheet.write('O1', 'Subject 5', bold)
            worksheet.write('P1', 'Marks 5', bold)
            worksheet.write('Q1', 'Total Marks', bold)
            worksheet.write('R1', 'Obtained Marks', bold)
            worksheet.write('S1', 'Percentage', bold)
            # start progress

            self.msg2.show()

            for n in range(start, end + 1):

                u = str(n)

                url_final = "https://www.msbte.org.in/SHTFNL20BTERESLIVE/SHTFNL20BTERESLIVE/SeatNumber/18/" + u + "Marksheet.html"
                res = ""
                try:
                    res = requests.get(url_final)
                except requests.ConnectionError as exception:
                    continue
                    # print("error")

                html_page = res.content
                soup = BeautifulSoup(html_page, 'html.parser')
                text = soup.find_all(text=True)
                output = ''
                blacklist = ['[document]', 'noscript', 'header', 'html', 'meta', 'head', 'title', 'input', 'script',
                             'style']
                for t in text:
                    if t.parent.name not in blacklist:
                        output += '{}'.format(t) + "\n"
                output1 = re.sub(
                    "|Maharashtra State Board of Technical Education|MR. / MS.|Statement of Marks|TITLE OF S|ENROLLMENT NO.|EXAMNINATION|SEAT NO.|COURSE"
                    "|E & OE|Url:-http://www.msbte.org.in|Result Declared On 07/01/2020|Ref:Formerly known as The board of Technical Examinations Maharashtra State of Technical Education Act 1997|Mah XXXVIII of 1997"
                    "|and Maharashtra Government Gazette Notification Section IV-B issued on march 31,1999|-172/16/200/20-04:07:2007 12:00:28"
                    "|Progressive Assessment|End Semester Exam|Distinction|Result Withheld Due To Pending Lower Semester|WFLS|Failure Marks|Additional Practical|"
                    "|Failure But Allowed To Keep Term|Pending Lower Year|Condoned Marks|Aggregate|Condoned|Lower Semester Pending|Industrial Training|Practical Test Marks"
                    "|Optional|Sessional|Practical|Result Withheld Due to Pending Lower Year|Exemption|Project Work|Theory Test Marks|Percentage of Marks|Absent|Team Work|Theory|"
                    "|ABBREVATION DETAILS|Class awarded for Diploma is based on aggregate marks obtained in pre-final & final semester|Allowed to Keep Term"
                    "|Carry Forward Marks|Candidate is eligible for admission to V/VII Semester only if he/she is fully passed in I & II /III & IV semesters & availed benefit of A.T.K.T/PASS at III & IV /V & VI semesters taken together respectively"
                    "|  number of failure subjects in I & II semesters taken  together|"
                    "|Eligibility for III semester is based on total|This certificate of marks is issued as per prevaling rules and regulations of MSBTE at the time of this exam|"
                    "|Report Discrepancy in this certificate to Head of the institution|INSTRUCTIONS|MAHARASHTRA STATE BOARD OF TECHNICAL EDUCATION|SECRETARY|"
                    "|This Marksheet is Downloaded from Internet|RESULT WITH|MARKS|TOTAL MAX|DATE :|OBTAINED|TOTAL CREDIT|CREDITS|MAX|OBT|MAX|  HEAD|DIST|C#|Oral|OR|A.T.K.T|"
                    "|PLY|ESE|@|"
                    "|TOTAL MARKS|", "", output)

                outputfinal = output1.replace("\n", " ").lstrip().replace("               ", "\n#").replace(
                    "      ",
                    "\n#").replace(
                    "   ", "\n#")
                # print(outputfinal)

                # name etracting
                i = 0
                name = ''
                ans = ''
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    name = name + ans
                    i = i + 1
                # print("Name : ", name)

                # enrolment extracting
                occurrence = 1
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                enrolment = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    enrolment = enrolment + ans
                    i = i + 1
                # print("Enrolment : ", enrolment)
                # examination
                occurrence = 2
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                examination = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    examination = examination + ans
                    i = i + 1
                # print("Examination : ", examination)
                # seat
                occurrence = 3
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                seatno = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    seatno = seatno + ans
                    i = i + 1
                # print("Seat Number : ", seatno)
                # semester
                occurrence = 4
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                semester = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    semester = semester + ans
                    i = i + 1
                # print("Semester : ", semester)
                # course
                occurrence = 6
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                course = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    course = course + ans
                    i = i + 1
                # print("Course : ", course)

                # SUbject 1
                occurrence = 10
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub1 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub1 = sub1 + ans
                    i = i + 1
                sub1.replace("TH", "")

                # print("Subject 1 : ", sub1.replace("TH", ""))

                # SUbject 1 marks
                occurrence = 11
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m1 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m1 = m1 + ans
                    i = i + 1
                m1t = m1[1] + m1[2] + m1[3]
                ma1 = m1[5] + m1[6] + m1[7]
                # print("Total Marks : ", m1t)
                # print("Obtained Marks : ", ma1)
                # SUbject 2
                occurrence = 16
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub2 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub2 = sub2 + ans
                    i = i + 1
                sub2.replace("TH", "")
                # print("Subject 2 : ", sub2.replace("TH", ""))
                # SUbject 2 marks
                occurrence = 17
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m2 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m2 = m2 + ans
                    i = i + 1
                m2t = m2[1] + m2[2] + m2[3]
                ma2 = m2[5] + m2[6] + m2[7]
                # print("Total Marks : ", m2t)
                # print("Obtained Marks : ", ma2)
                # SUbject 3
                occurrence = 22
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub3 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub3 = sub3 + ans
                    i = i + 1
                sub3.replace("TH", "")
                # print("Subject 3 : ", sub3.replace("TH", ""))
                # SUbject 3 marks
                occurrence = 23
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m3 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m3 = m3 + ans
                    i = i + 1
                m3t = m3[1] + m3[2] + m3[3]
                ma3 = m3[5] + m3[6] + m3[7]
                # print("Total Marks : ", m3t)
                # print("Obtained Marks : ", ma3)
                # SUbject 4
                occurrence = 28
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub4 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub4 = sub4 + ans
                    i = i + 1
                sub4.replace("TH", "")
                # print("Subject 4 : ", sub4.replace("TH", ""))
                # SUbject 4 marks
                occurrence = 29
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m4 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m4 = m4 + ans
                    i = i + 1
                m4t = m4[1] + m4[2] + m4[3]
                ma4 = m4[5] + m4[6] + m4[7]
                # print("Total Marks : ", m4t)
                # print("Obtained Marks : ", ma4)

                # SUbject 5
                occurrence = 31
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub5 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub5 = sub5 + ans
                    i = i + 1
                sub5.replace("TH", "")
                # print("Subject 5 : ", sub5.replace("TH", ""))
                # SUbject 5 marks
                occurrence = 32
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m5 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m5 = m5 + ans
                    i = i + 1
                m5t = m5[1] + m5[2] + m5[3]
                ma5 = m5[5] + m5[6] + m5[7]
                # print("Total Marks : ", m5t)
                # print("Obtained Marks : ", ma5)

                # Total marks / marks obtained/ percentage
                occurrence = 38
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                total = ""
                percent = ""
                obtain = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    total = total + ans
                    i = i + 1
                total.lstrip()

                percent = total[6:11]
                obtain = total[11:16]
                totalm = total[0:6]
                # print("total : ", total[0:6])
                # print("obtain : ", obtain)
                # print("percent : ", percent)

                # save data

                # The workbook object is then used to add new
                # worksheet via the add_worksheet() method.

                # Use the worksheet object to write
                # data via the write() method.
                index = str(s)
                # style =('font: bold 1, color red;')
                worksheet.write('A' + index, name)
                worksheet.write('B' + index, enrolment)
                worksheet.write('C' + index, examination)
                worksheet.write('D' + index, seatno)
                worksheet.write('E' + index, semester)
                worksheet.write('F' + index, course)
                worksheet.write('G' + index, sub1)
                worksheet.write('H' + index, ma1)
                worksheet.write('I' + index, sub2)
                worksheet.write('J' + index, ma2)
                worksheet.write('K' + index, sub3)
                worksheet.write('L' + index, ma3)
                worksheet.write('M' + index, sub4)
                worksheet.write('N' + index, ma4)
                worksheet.write('O' + index, sub5)
                worksheet.write('P' + index, ma5)
                worksheet.write('Q' + index, totalm)
                worksheet.write('R' + index, obtain)
                worksheet.write('S' + index, percent)
                s = s + 1
                # close the Excel file
            workbook.close()
            self.msg2.close()
            self.msg.show()
        except:

            self.msg1.show()
            self.msg2.close()

    def Semester_2(self):
        try:
            start1 = self.textEdit.toPlainText()
            end1 = self.textEdit_2.toPlainText()
            file = self.textEdit_3.toPlainText()
            start = int(start1)
            end = int(end1)

            # start = int(input("Enter Starting seat Number:"))
            # end = int(input("Enter Ending seat Number:"))
            # file = input("Enter Excel File(If already exist enter same name or create with new name):")
            s = 2;
            workbook = xlsxwriter.Workbook(file + '.xlsx')
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold': True})
            worksheet.write('A1', 'NAME', bold)
            worksheet.write('B1', 'Enrolment', bold)
            worksheet.write('C1', 'Examination', bold)
            worksheet.write('D1', 'Seat Number', bold)
            worksheet.write('E1', 'Semester', bold)
            worksheet.write('F1', 'Course', bold)
            worksheet.write('G1', 'Subject 1 ', bold)
            worksheet.write('H1', 'Marks 1', bold)
            worksheet.write('I1', 'Subject 2', bold)
            worksheet.write('J1', 'Marks 2', bold)
            worksheet.write('K1', 'Subject 3', bold)
            worksheet.write('L1', 'Marks 3', bold)
            worksheet.write('M1', 'Subject 4', bold)
            worksheet.write('N1', 'Marks 4', bold)
            worksheet.write('O1', 'Subject 5', bold)
            worksheet.write('P1', 'Marks 5', bold)
            worksheet.write('Q1', 'Subject 6', bold)
            worksheet.write('R1', 'Marks 6', bold)
            worksheet.write('S1', 'Subject 7', bold)
            worksheet.write('T1', 'Marks 7', bold)
            worksheet.write('U1', 'Total Marks', bold)
            worksheet.write('V1', 'Obtained Marks', bold)
            worksheet.write('W1', 'Percentage', bold)
            #start progress


            self.msg2.show()
            for n in range(start, end + 1):

                u = str(n)
                # you need to change this
                url_final = "https://msbte.org.in/SOLCOVD2020RESBTELIVE/SNONFNL20246RESLIVE/SeatNumber/18/" + u + "Marksheet.html"

                try:
                    res = requests.get(url_final)
                except requests.ConnectionError as exception:
                    continue
                    #print("error")



                html_page = res.content
                soup = BeautifulSoup(html_page, 'html.parser')
                text = soup.find_all(text=True)
                output = ''
                blacklist = ['[document]', 'noscript', 'header', 'html', 'meta', 'head', 'title', 'input', 'script',
                             'style']
                for t in text:
                    if t.parent.name not in blacklist:
                        output += '{}'.format(t) + "\n"
                output1 = re.sub(
                    "|Maharashtra State Board of Technical Education|MR. / MS.|Statement of Marks|TITLE OF S|ENROLLMENT NO.|EXAMNINATION|SEAT NO.|COURSE"
                    "|E & OE|Url:-http://www.msbte.org.in|Result Declared On 07/01/2020|Ref:Formerly known as The board of Technical Examinations Maharashtra State of Technical Education Act 1997|Mah XXXVIII of 1997"
                    "|and Maharashtra Government Gazette Notification Section IV-B issued on march 31,1999|-172/16/200/20-04:07:2007 12:00:28"
                    "|Progressive Assessment|End Semester Exam|Distinction|Result Withheld Due To Pending Lower Semester|WFLS|Failure Marks|Additional Practical|"
                    "|Failure But Allowed To Keep Term|Pending Lower Year|Condoned Marks|Aggregate|Condoned|Lower Semester Pending|Industrial Training|Practical Test Marks"
                    "|Optional|Sessional|Practical|Result Withheld Due to Pending Lower Year|Exemption|Project Work|Theory Test Marks|Percentage of Marks|Absent|Team Work|Theory|"
                    "|ABBREVATION DETAILS|Class awarded for Diploma is based on aggregate marks obtained in pre-final & final semester|Allowed to Keep Term"
                    "|Carry Forward Marks|Candidate is eligible for admission to V/VII Semester only if he/she is fully passed in I & II /III & IV semesters & availed benefit of A.T.K.T/PASS at III & IV /V & VI semesters taken together respectively"
                    "|  number of failure subjects in I & II semesters taken  together|"
                    "|Eligibility for III semester is based on total|This certificate of marks is issued as per prevaling rules and regulations of MSBTE at the time of this exam|"
                    "|Report Discrepancy in this certificate to Head of the institution|INSTRUCTIONS|MAHARASHTRA STATE BOARD OF TECHNICAL EDUCATION|SECRETARY|"
                    "|This Marksheet is Downloaded from Internet|RESULT WITH|MARKS|TOTAL MAX|DATE :|OBTAINED|TOTAL CREDIT|CREDITS|MAX|MIN|OBT|MAX|  HEAD|DIST|C#|Oral|OR|A.T.K.T|"
                    "|PLY|ESE|@|"
                    "|TOTAL MARKS|", "", output)

                outputfinal = output1.replace("\n", " ").lstrip().replace("               ", "\n#").replace("      ",
                                                                                                            "\n#").replace(
                    "   ", "\n#")
               # print(outputfinal)

                # name etracting
                i = 0
                name = ''
                ans = ''
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    name = name + ans
                    i = i + 1
                #print("Name : ", name)

                # enrolment extracting
                occurrence = 1
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                enrolment = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    enrolment = enrolment + ans
                    i = i + 1
               # print("Enrolment : ", enrolment)
                # examination
                occurrence = 2
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                examination = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    examination = examination + ans
                    i = i + 1
                #print("Examination : ", examination)
                # seat
                occurrence = 3
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                seatno = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    seatno = seatno + ans
                    i = i + 1
                #print("Seat Number : ", seatno)
                # semester
                occurrence = 4
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                semester = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    semester = semester + ans
                    i = i + 1
                #print("Semester : ", semester)
                # course
                occurrence = 6
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                course = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    course = course + ans
                    i = i + 1
               # print("Course : ", course)

                # SUbject 1
                occurrence = 10
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub1 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub1 = sub1 + ans
                    i = i + 1
                sub1.replace("TH", "")
               # print("Subject 1 : ", sub1.replace("TH", ""))

                # SUbject 1 marks
                occurrence = 11
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m1 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m1 = m1 + ans
                    i = i + 1
                m1t = m1[1] + m1[2] + m1[3]
                ma1 = m1[5] + m1[6] + m1[7]
                #print("Total Marks : ", m1t)
                #print("Obtained Marks : ", ma1)
                # SUbject 2
                occurrence = 16
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub2 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub2 = sub2 + ans
                    i = i + 1
                sub2.replace("TH", "")
                #print("Subject 2 : ", sub2.replace("TH", ""))
                # SUbject 2 marks
                occurrence = 17
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m2 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m2 = m2 + ans
                    i = i + 1
                m2t = m2[1] + m2[2] + m2[3]
                ma2 = m2[5] + m2[6] + m2[7]
               # print("Total Marks : ", m2t)
               # print("Obtained Marks : ", ma2)
                # SUbject 3
                occurrence = 19
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub3 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub3 = sub3 + ans
                    i = i + 1
                sub3.replace("TH", "")
                #print("Subject 3 : ", sub3.replace("TH", ""))
                # SUbject 3 marks
                occurrence = 20
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m3 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m3 = m3 + ans
                    i = i + 1
                m3t = m3[1] + m3[2] + m3[3]
                ma3 = m3[5] + m3[6] + m3[7]
                #print("Total Marks : ", m3t)
               #print("Obtained Marks : ", ma3)
                # SUbject 4
                occurrence = 25
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub4 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub4 = sub4 + ans
                    i = i + 1
                sub4.replace("TH", "")
                #print("Subject 4 : ", sub4.replace("TH", ""))
                # SUbject 4 marks
                occurrence = 26
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m4 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m4 = m4 + ans
                    i = i + 1
                m4t = m4[1] + m4[2] + m4[3]
                ma4 = m4[5] + m4[6] + m4[7]
                #print("Total Marks : ", m4t)
               # print("Obtained Marks : ", ma4)
                # SUbject 5
                occurrence = 31
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub5 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub5 = sub5 + ans
                    i = i + 1
                sub5.replace("TH", "")
                #print("Subject 5 : ", sub5.replace("TH", ""))
                #SUbject 5 marks
                occurrence = 32
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m5 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m5 = m5 + ans
                    i = i + 1
                m5t = m5[1] + m5[2] + m5[3]
                ma5 = m5[5] + m5[6] + m5[7]
                #print("Total Marks : ", m5t)
                #print("Obtained Marks : ", ma5)
                # SUbject 6
                occurrence = 34
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub6 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub6 = sub6 + ans
                    i = i + 1
                sub6.replace("PR", "")
                #print("Subject 6 : ", sub6.replace("PR", ""))
                # SUbject 6 marks
                occurrence = 35
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m6 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m6 = m6 + ans
                    i = i + 1
                m6t = m6[1] + m6[2] + m6[3]
                ma6 = m6[5] + m6[6] + m6[7]
               # print("Total Marks : ", m6t)
                #print("Obtained Marks : ", ma6)
                # SUbject 7
                occurrence = 37
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub7 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub7 = sub7 + ans
                    i = i + 1
                sub7.replace("PR", "")
               # print("Subject 7 : ", sub7.replace("PR", ""))
                # SUbject 7 marks
                occurrence = 38
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m7 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m7 = m7 + ans
                    i = i + 1
                m7t = m7[1] + m7[2] + m7[3]
                ma7 = m7[5] + m7[6] + m7[7]
               # print("Total Marks : ", m7t)
                #print("Obtained Marks : ", ma7)

                # Total marks / marks obtained/ percentage
                occurrence = 44
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                total = ""
                percent = ""
                obtain = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    total = total + ans
                    i = i + 1
                total.lstrip()

                percent = total[6:11]
                obtain = total[11:16]
                totalm = total[0:6]
                #print("total : ", total[0:6])
               # print("obtain : ", obtain)
               # print("percent : ", percent)

                # save data

                # The workbook object is then used to add new
                # worksheet via the add_worksheet() method.

                # Use the worksheet object to write
                # data via the write() method.
                index = str(s)
                # style =('font: bold 1, color red;')
                worksheet.write('A' + index, name)
                worksheet.write('B' + index, enrolment)
                worksheet.write('C' + index, examination)
                worksheet.write('D' + index, seatno)
                worksheet.write('E' + index, semester)
                worksheet.write('F' + index, course)
                worksheet.write('G' + index, sub1)
                worksheet.write('H' + index, ma1)
                worksheet.write('I' + index, sub2)
                worksheet.write('J' + index, ma2)
                worksheet.write('K' + index, sub3)
                worksheet.write('L' + index, ma3)
                worksheet.write('M' + index, sub4)
                worksheet.write('N' + index, ma4)
                worksheet.write('O' + index, sub5)
                worksheet.write('P' + index, ma5)
                worksheet.write('Q' + index, sub6)
                worksheet.write('R' + index, ma6)
                worksheet.write('S' + index, sub7)
                worksheet.write('T' + index, ma7)
                worksheet.write('U' + index, totalm)
                worksheet.write('V' + index, obtain)
                worksheet.write('W' + index, percent)
                s = s + 1
                # close the Excel file
            workbook.close()
            self.msg2.close()
            self.msg.show()
        except:
            self.msg1.show()
            self.msg2.close()

    def Semester_3(self):
        try:
            start1 = self.textEdit.toPlainText()
            end1 = self.textEdit_2.toPlainText()
            file = self.textEdit_3.toPlainText()
            start = int(start1)
            end = int(end1)

            # start = int(input("Enter Starting seat Number:"))
            # end = int(input("Enter Ending seat Number:"))
            # file = input("Enter Excel File(If already exist enter same name or create with new name):")
            s = 2;
            workbook = xlsxwriter.Workbook(file + '.xlsx')
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold': True})
            worksheet.write('A1', 'NAME', bold)
            worksheet.write('B1', 'Enrolment', bold)
            worksheet.write('C1', 'Examination', bold)
            worksheet.write('D1', 'Seat Number', bold)
            worksheet.write('E1', 'Semester', bold)
            worksheet.write('F1', 'Course', bold)
            worksheet.write('G1', 'Subject 1 ', bold)
            worksheet.write('H1', 'Marks 1', bold)
            worksheet.write('I1', 'Subject 2', bold)
            worksheet.write('J1', 'Marks 2', bold)
            worksheet.write('K1', 'Subject 3', bold)
            worksheet.write('L1', 'Marks 3', bold)
            worksheet.write('M1', 'Subject 4', bold)
            worksheet.write('N1', 'Marks 4', bold)
            worksheet.write('O1', 'Subject 5', bold)
            worksheet.write('P1', 'Marks 5', bold)
            worksheet.write('Q1', 'Total Marks', bold)
            worksheet.write('R1', 'Obtained Marks', bold)
            worksheet.write('S1', 'Percentage', bold)
            # start progress

            self.msg2.show()


            for n in range(start, end + 1):

                u = str(n)

                url_final = "https://www.msbte.org.in/SHTFNL20BTERESLIVE/SHTFNL20BTERESLIVE/SeatNumber/18/" + u + "Marksheet.html"
                res = ""
                try:
                    res = requests.get(url_final)
                except requests.ConnectionError as exception:
                    continue
                    #print("error")

                html_page = res.content
                soup = BeautifulSoup(html_page, 'html.parser')
                text = soup.find_all(text=True)
                output = ''
                blacklist = ['[document]', 'noscript', 'header', 'html', 'meta', 'head', 'title', 'input', 'script',
                             'style']
                for t in text:
                    if t.parent.name not in blacklist:
                        output += '{}'.format(t) + "\n"
                output1 = re.sub(
                    "|Maharashtra State Board of Technical Education|MR. / MS.|Statement of Marks|TITLE OF S|ENROLLMENT NO.|EXAMNINATION|SEAT NO.|COURSE"
                    "|E & OE|Url:-http://www.msbte.org.in|Result Declared On 07/01/2020|Ref:Formerly known as The board of Technical Examinations Maharashtra State of Technical Education Act 1997|Mah XXXVIII of 1997"
                    "|and Maharashtra Government Gazette Notification Section IV-B issued on march 31,1999|-172/16/200/20-04:07:2007 12:00:28"
                    "|Progressive Assessment|End Semester Exam|Distinction|Result Withheld Due To Pending Lower Semester|WFLS|Failure Marks|Additional Practical|"
                    "|Failure But Allowed To Keep Term|Pending Lower Year|Condoned Marks|Aggregate|Condoned|Lower Semester Pending|Industrial Training|Practical Test Marks"
                    "|Optional|Sessional|Practical|Result Withheld Due to Pending Lower Year|Exemption|Project Work|Theory Test Marks|Percentage of Marks|Absent|Team Work|Theory|"
                    "|ABBREVATION DETAILS|Class awarded for Diploma is based on aggregate marks obtained in pre-final & final semester|Allowed to Keep Term"
                    "|Carry Forward Marks|Candidate is eligible for admission to V/VII Semester only if he/she is fully passed in I & II /III & IV semesters & availed benefit of A.T.K.T/PASS at III & IV /V & VI semesters taken together respectively"
                    "|  number of failure subjects in I & II semesters taken  together|"
                    "|Eligibility for III semester is based on total|This certificate of marks is issued as per prevaling rules and regulations of MSBTE at the time of this exam|"
                    "|Report Discrepancy in this certificate to Head of the institution|INSTRUCTIONS|MAHARASHTRA STATE BOARD OF TECHNICAL EDUCATION|SECRETARY|"
                    "|This Marksheet is Downloaded from Internet|RESULT WITH|MARKS|TOTAL MAX|DATE :|OBTAINED|TOTAL CREDIT|CREDITS|MAX|OBT|MAX|  HEAD|DIST|C#|Oral|OR|A.T.K.T|"
                    "|PLY|ESE|@|"
                    "|TOTAL MARKS|", "", output)

                outputfinal = output1.replace("\n", " ").lstrip().replace("               ", "\n#").replace("      ",
                                                                                                            "\n#").replace(
                    "   ", "\n#")
                #print(outputfinal)

                # name etracting
                i = 0
                name = ''
                ans = ''
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    name = name + ans
                    i = i + 1
                #print("Name : ", name)

                # enrolment extracting
                occurrence = 1
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                enrolment = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    enrolment = enrolment + ans
                    i = i + 1
                #print("Enrolment : ", enrolment)
                # examination
                occurrence = 2
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                examination = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    examination = examination + ans
                    i = i + 1
                #print("Examination : ", examination)
                # seat
                occurrence = 3
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                seatno = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    seatno = seatno + ans
                    i = i + 1
                #print("Seat Number : ", seatno)
                # semester
                occurrence = 4
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                semester = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    semester = semester + ans
                    i = i + 1
                #print("Semester : ", semester)
                # course
                occurrence = 6
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                course = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    course = course + ans
                    i = i + 1
                #print("Course : ", course)

                # SUbject 1
                occurrence = 10
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub1 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub1 = sub1 + ans
                    i = i + 1
                sub1.replace("TH", "")

                #print("Subject 1 : ", sub1.replace("TH", ""))

                # SUbject 1 marks
                occurrence = 11
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m1 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m1 = m1 + ans
                    i = i + 1
                m1t = m1[1] + m1[2] + m1[3]
                ma1 = m1[5] + m1[6] + m1[7]
                #print("Total Marks : ", m1t)
                #print("Obtained Marks : ", ma1)
                # SUbject 2
                occurrence = 16
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub2 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub2 = sub2 + ans
                    i = i + 1
                sub2.replace("TH", "")
                #print("Subject 2 : ", sub2.replace("TH", ""))
                # SUbject 2 marks
                occurrence = 17
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m2 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m2 = m2 + ans
                    i = i + 1
                m2t = m2[1] + m2[2] + m2[3]
                ma2 = m2[5] + m2[6] + m2[7]
                #print("Total Marks : ", m2t)
                #print("Obtained Marks : ", ma2)
                # SUbject 3
                occurrence = 22
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub3 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub3 = sub3 + ans
                    i = i + 1
                sub3.replace("TH", "")
                #print("Subject 3 : ", sub3.replace("TH", ""))
                # SUbject 3 marks
                occurrence = 23
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m3 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m3 = m3 + ans
                    i = i + 1
                m3t = m3[1] + m3[2] + m3[3]
                ma3 = m3[5] + m3[6] + m3[7]
                #print("Total Marks : ", m3t)
                #print("Obtained Marks : ", ma3)
                # SUbject 4
                occurrence = 28
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub4 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub4 = sub4 + ans
                    i = i + 1
                sub4.replace("TH", "")
                #print("Subject 4 : ", sub4.replace("TH", ""))
                # SUbject 4 marks
                occurrence = 29
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m4 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m4 = m4 + ans
                    i = i + 1
                m4t = m4[1] + m4[2] + m4[3]
                ma4 = m4[5] + m4[6] + m4[7]
                #print("Total Marks : ", m4t)
                #print("Obtained Marks : ", ma4)
                # SUbject 5
                occurrence = 34
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                sub5 = ""
                ans = ""
                while outputfinal[i] != "#" and outputfinal[i].isdigit() == False:
                    ans = outputfinal[i]
                    sub5 = sub5 + ans
                    i = i + 1
                sub5.replace("TH", "")
                #print("Subject 5 : ", sub5.replace("TH", ""))
                # SUbject 5 marks
                occurrence = 35
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                m5 = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    m5 = m5 + ans
                    i = i + 1
                m5t = m5[1] + m5[2] + m5[3]
                ma5 = m5[5] + m5[6] + m5[7]
                #print("Total Marks : ", m5t)
                #print("Obtained Marks : ", ma5)

                # Total marks / marks obtained/ percentage
                occurrence = 44
                val = -1
                for i in range(0, occurrence):
                    val = outputfinal.find("#", val + 1)
                i = val + 1
                total = ""
                percent = ""
                obtain = ""
                ans = ""
                while outputfinal[i] != "#":
                    ans = outputfinal[i]
                    total = total + ans
                    i = i + 1
                total.lstrip()

                percent = total[6:11]
                obtain = total[11:16]
                totalm = total[0:6]
                #print("total : ", total[0:6])
                #print("obtain : ", obtain)
                #print("percent : ", percent)

                # save data

                # The workbook object is then used to add new
                # worksheet via the add_worksheet() method.

                # Use the worksheet object to write
                # data via the write() method.
                index = str(s)
                # style =('font: bold 1, color red;')
                worksheet.write('A' + index, name)
                worksheet.write('B' + index, enrolment)
                worksheet.write('C' + index, examination)
                worksheet.write('D' + index, seatno)
                worksheet.write('E' + index, semester)
                worksheet.write('F' + index, course)
                worksheet.write('G' + index, sub1)
                worksheet.write('H' + index, ma1)
                worksheet.write('I' + index, sub2)
                worksheet.write('J' + index, ma2)
                worksheet.write('K' + index, sub3)
                worksheet.write('L' + index, ma3)
                worksheet.write('M' + index, sub4)
                worksheet.write('N' + index, ma4)
                worksheet.write('O' + index, sub5)
                worksheet.write('P' + index, ma5)
                worksheet.write('Q' + index, totalm)
                worksheet.write('R' + index, obtain)
                worksheet.write('S' + index, percent)
                s = s + 1
                # close the Excel file
            workbook.close()
            self.msg2.close()
            self.msg.show()
        except:

            self.msg1.show()
            self.msg2.close()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
