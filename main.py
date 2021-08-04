from PyQt5 import QtCore, QtGui, QtWidgets
from gen_template import *
from download_data import *
from os import listdir

class Ui_PhanMemThue(object):
    def setupUi(self, PhanMemThue):
        PhanMemThue.setObjectName("PhanMemThue")
        PhanMemThue.resize(490, 630)
        self.centralwidget = QtWidgets.QWidget(PhanMemThue)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 491, 621))
        self.tabWidget.setObjectName("tabWidget")
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.tabWidget.setFont(font)
        self.tabWidget.setFocusPolicy(QtCore.Qt.ClickFocus)
        self.GenTemplateTab = QtWidgets.QWidget()
        self.GenTemplateTab.setStyleSheet("background-color: #e6f2ff;")
        self.GenTemplateTab.setObjectName("GenTemplateTab")
        self.ExecuteBotton = QtWidgets.QPushButton(self.GenTemplateTab)
        self.ExecuteBotton.setGeometry(QtCore.QRect(30, 550, 421, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.ExecuteBotton.setFont(font)
        self.ExecuteBotton.setObjectName("ExecuteBotton")
        self.ExecuteBotton.setStyleSheet("background-color: #1b4f59; color: white; font-weight: bold;")
        self.OutputGroupBox = QtWidgets.QGroupBox(self.GenTemplateTab)
        self.OutputGroupBox.setGeometry(QtCore.QRect(30, 330, 421, 201))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(False)
        font.setWeight(50)
        self.OutputGroupBox.setFont(font)
        self.OutputGroupBox.setAutoFillBackground(True)
        self.OutputGroupBox.setFlat(False)
        self.OutputGroupBox.setCheckable(False)
        self.OutputGroupBox.setObjectName("OutputGroupBox")
        self.OutputExcelGroupBox = QtWidgets.QGroupBox(self.OutputGroupBox)
        self.OutputExcelGroupBox.setGeometry(QtCore.QRect(10, 120, 401, 71))
        self.OutputExcelGroupBox.setAutoFillBackground(True)
        self.OutputExcelGroupBox.setObjectName("OutputExcelGroupBox")
        self.SelectExcelFileButton = QtWidgets.QPushButton(self.OutputExcelGroupBox)
        self.SelectExcelFileButton.setGeometry(QtCore.QRect(310, 30, 81, 21))
        self.SelectExcelFileButton.setObjectName("SelectExcelFileButton")
        self.SelectExcelFileButton.setStyleSheet("background-color: #1b4f59; color: white; font-weight: bold;")
        self.OutputExcelPath = QtWidgets.QLineEdit(self.OutputExcelGroupBox)
        self.OutputExcelPath.setGeometry(QtCore.QRect(10, 30, 291, 21))
        self.OutputExcelPath.setObjectName("OutputExcelPath")
        self.OutputExcelPath.setText('')
        self.TemplateFolderGroupBox = QtWidgets.QGroupBox(self.OutputGroupBox)
        self.TemplateFolderGroupBox.setGeometry(QtCore.QRect(10, 30, 401, 71))
        self.TemplateFolderGroupBox.setAutoFillBackground(True)
        self.TemplateFolderGroupBox.setObjectName("TemplateFolderGroupBox")
        self.SelectTemplateFolderButtom = QtWidgets.QPushButton(self.TemplateFolderGroupBox)
        self.SelectTemplateFolderButtom.setGeometry(QtCore.QRect(310, 30, 81, 21))
        self.SelectTemplateFolderButtom.setObjectName("SelectTemplateFolderButtom")
        self.SelectTemplateFolderButtom.setStyleSheet("background-color: #1b4f59; color: white; font-weight: bold;")
        self.TemplateFolderPath = QtWidgets.QLineEdit(self.TemplateFolderGroupBox)
        self.TemplateFolderPath.setGeometry(QtCore.QRect(10, 30, 291, 21))
        self.TemplateFolderPath.setObjectName("TemplateFolderPath")
        self.TemplateFolderPath.setText('')
        self.InputGroupBox = QtWidgets.QGroupBox(self.GenTemplateTab)
        self.InputGroupBox.setGeometry(QtCore.QRect(30, 10, 421, 311))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(False)
        font.setWeight(50)
        self.InputGroupBox.setFont(font)
        self.InputGroupBox.setAutoFillBackground(True)
        self.InputGroupBox.setObjectName("InputGroupBox")
        # self.InputGroupBox.setStyleSheet("background-color: red; font-weight: bold;")
        self.InputScriptFolderGroupBox = QtWidgets.QGroupBox(self.InputGroupBox)
        self.InputScriptFolderGroupBox.setGeometry(QtCore.QRect(10, 30, 401, 71))
        self.InputScriptFolderGroupBox.setAutoFillBackground(True)
        self.InputScriptFolderGroupBox.setObjectName("InputScriptFolderGroupBox")
        self.SelectScriptFolderButtom = QtWidgets.QPushButton(self.InputScriptFolderGroupBox)
        self.SelectScriptFolderButtom.setGeometry(QtCore.QRect(310, 30, 81, 21))
        self.SelectScriptFolderButtom.setObjectName("SelectScriptFolderButtom")
        self.SelectScriptFolderButtom.setStyleSheet("background-color: #1b4f59; color: white; font-weight: bold;")

        self.ScriptsFolderPath = QtWidgets.QLineEdit(self.InputScriptFolderGroupBox)
        self.ScriptsFolderPath.setGeometry(QtCore.QRect(10, 30, 291, 21))
        self.ScriptsFolderPath.setObjectName("ScriptsFolderPath")
        self.ScriptsFolderPath.setText('')

        self.ListScriptFileScroll = QtWidgets.QWidget(self.InputGroupBox)
        self.ListScriptFileScroll.setGeometry(QtCore.QRect(10, 110, 401, 191))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.ListScriptFileScroll.setFont(font)
        self.ListScriptFileScroll.setObjectName("ListScriptFileScroll")
        area = QtWidgets.QScrollArea(self.ListScriptFileScroll)
        area.setWidgetResizable(True)
        scrollAreaWidgetContents = QtWidgets.QWidget()
        self.LayoutVScroll = QtWidgets.QVBoxLayout(scrollAreaWidgetContents)
        area.setWidget(scrollAreaWidgetContents)
        layoutV = QtWidgets.QVBoxLayout(self.ListScriptFileScroll)
        layoutV.addWidget(area)
        self.tabWidget.addTab(self.GenTemplateTab, "GenTemplateTab")
        self.GetDataTab = QtWidgets.QWidget()
        self.GetDataTab.setStyleSheet("background-color: #f9ecf2;")
        self.GetDataTab.setObjectName("GetDataTab")
        self.InputGroupBox_S2 = QtWidgets.QGroupBox(self.GetDataTab)
        self.InputGroupBox_S2.setGeometry(QtCore.QRect(30, 20, 421, 501))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(False)
        font.setWeight(50)
        self.InputGroupBox_S2.setFont(font)
        self.InputGroupBox_S2.setAutoFillBackground(True)
        self.InputGroupBox_S2.setObjectName("InputGroupBox_S2")
        self.InputExcelFileGroupBox_S2 = QtWidgets.QGroupBox(self.InputGroupBox_S2)
        self.InputExcelFileGroupBox_S2.setGeometry(QtCore.QRect(10, 30, 401, 71))
        self.InputExcelFileGroupBox_S2.setAutoFillBackground(True)
        self.InputExcelFileGroupBox_S2.setObjectName("InputExcelFileGroupBox_S2")
        self.SelectExcelFileButtom_S2 = QtWidgets.QPushButton(self.InputExcelFileGroupBox_S2)
        self.SelectExcelFileButtom_S2.setGeometry(QtCore.QRect(310, 30, 81, 21))
        self.SelectExcelFileButtom_S2.setObjectName("SelectExcelFileButtom_S2")
        self.SelectExcelFileButtom_S2.setStyleSheet("background-color: #1b4f59; color: white; font-weight: bold;")
        self.ExcelFilePath_S2 = QtWidgets.QLineEdit(self.InputExcelFileGroupBox_S2)
        self.ExcelFilePath_S2.setText('')
        self.ExcelFilePath_S2.setGeometry(QtCore.QRect(10, 30, 291, 21))
        self.ExcelFilePath_S2.setObjectName("ExcelFilePath_S2")
        self.ListScriptFileScroll_S2 = QtWidgets.QWidget(self.InputGroupBox_S2)
        self.ListScriptFileScroll_S2.setGeometry(QtCore.QRect(10, 200, 401, 291))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.ListScriptFileScroll_S2.setFont(font)
        self.ListScriptFileScroll_S2.setObjectName("ListScriptFileScroll_S2")

        area_s2 = QtWidgets.QScrollArea(self.ListScriptFileScroll_S2)
        area_s2.setWidgetResizable(True)
        scrollAreaWidgetContents_S2 = QtWidgets.QWidget()
        self.LayoutVScroll_S2 = QtWidgets.QVBoxLayout(scrollAreaWidgetContents_S2)
        area_s2.setWidget(scrollAreaWidgetContents_S2)
        layoutV_S2 = QtWidgets.QVBoxLayout(self.ListScriptFileScroll_S2)
        layoutV_S2.addWidget(area_s2)

        self.InputTemplateFolderGroupBox_S2 = QtWidgets.QGroupBox(self.InputGroupBox_S2)
        self.InputTemplateFolderGroupBox_S2.setGeometry(QtCore.QRect(10, 120, 401, 71))
        self.InputTemplateFolderGroupBox_S2.setAutoFillBackground(True)
        self.InputTemplateFolderGroupBox_S2.setObjectName("InputTemplateFolderGroupBox_S2")
        self.SelectTemplateFolderButtom_S2 = QtWidgets.QPushButton(self.InputTemplateFolderGroupBox_S2)
        self.SelectTemplateFolderButtom_S2.setGeometry(QtCore.QRect(310, 30, 81, 21))
        self.SelectTemplateFolderButtom_S2.setObjectName("SelectTemplateFolderButtom_S2")
        self.SelectTemplateFolderButtom_S2.setStyleSheet("background-color: #1b4f59; color: white; font-weight: bold;")
        self.TemplateFolderPath_S2 = QtWidgets.QLineEdit(self.InputTemplateFolderGroupBox_S2)
        self.TemplateFolderPath_S2.setText('')
        self.TemplateFolderPath_S2.setGeometry(QtCore.QRect(10, 30, 291, 21))
        self.TemplateFolderPath_S2.setObjectName("TemplateFolderPath_S2")
        self.ExecuteBotton_S2 = QtWidgets.QPushButton(self.GetDataTab)
        self.ExecuteBotton_S2.setGeometry(QtCore.QRect(30, 540, 421, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.ExecuteBotton_S2.setFont(font)
        self.ExecuteBotton_S2.setObjectName("ExecuteBotton_S2")
        self.ExecuteBotton_S2.setStyleSheet("background-color: #1b4f59; color: white; font-weight: bold;")
        self.tabWidget.addTab(self.GetDataTab, "GetDataTab")
        PhanMemThue.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(PhanMemThue)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 490, 21))
        self.menubar.setObjectName("menubar")
        self.menuPh_n_m_m_thu = QtWidgets.QMenu(self.menubar)
        self.menuPh_n_m_m_thu.setObjectName("menuPh_n_m_m_thu")
        PhanMemThue.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(PhanMemThue)
        self.statusbar.setObjectName("statusbar")
        PhanMemThue.setStatusBar(self.statusbar)
        self.menubar.addAction(self.menuPh_n_m_m_thu.menuAction())

        self.retranslateUi(PhanMemThue)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(PhanMemThue)

        # buttom event
        # Tab 1
        self.SelectScriptFolderButtom.clicked.connect(self.select_script_folder)
        self.SelectTemplateFolderButtom.clicked.connect(self.select_template_folder)
        self.SelectExcelFileButton.clicked.connect(self.select_excel_file)
        self.ExecuteBotton.clicked.connect(self.handle_script_files)
        # Tab 2
        self.SelectExcelFileButtom_S2.clicked.connect(self.select_excel_file_tab2)
        self.SelectTemplateFolderButtom_S2.clicked.connect(self.select_template_folder_tab2)
        self.ExecuteBotton_S2.clicked.connect(self.run_script_file)

    def retranslateUi(self, PhanMemThue):
        _translate = QtCore.QCoreApplication.translate
        PhanMemThue.setWindowTitle(_translate("PhanMemThue", "MainWindow"))
        self.ExecuteBotton.setText(_translate("PhanMemThue", "Execute"))
        self.OutputGroupBox.setTitle(_translate("PhanMemThue", "Output"))
        self.OutputExcelGroupBox.setTitle(_translate("PhanMemThue", "Save to Excel File"))
        self.SelectExcelFileButton.setText(_translate("PhanMemThue", "Select File"))
        self.TemplateFolderGroupBox.setTitle(_translate("PhanMemThue", "Output Template Folder"))
        self.SelectTemplateFolderButtom.setText(_translate("PhanMemThue", "Select Folder"))
        self.InputGroupBox.setTitle(_translate("PhanMemThue", "Input"))
        self.InputScriptFolderGroupBox.setTitle(_translate("PhanMemThue", "Input Script Folder"))
        self.SelectScriptFolderButtom.setText(_translate("PhanMemThue", "Select Folder"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.GenTemplateTab), _translate("PhanMemThue", "Gen Teamplate"))
        self.InputGroupBox_S2.setTitle(_translate("PhanMemThue", "Input"))
        self.InputExcelFileGroupBox_S2.setTitle(_translate("PhanMemThue", "Excel File"))
        self.SelectExcelFileButtom_S2.setText(_translate("PhanMemThue", "Select File"))
        self.InputTemplateFolderGroupBox_S2.setTitle(_translate("PhanMemThue", "Script Template Folder"))
        self.SelectTemplateFolderButtom_S2.setText(_translate("PhanMemThue", "Select Folder"))
        self.ExecuteBotton_S2.setText(_translate("PhanMemThue", "Execute"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.GetDataTab), _translate("PhanMemThue", "Get Data"))
        self.menuPh_n_m_m_thu.setTitle(_translate("PhanMemThue", ""))

    def select_script_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory()
        self.ScriptsFolderPath.setText(folder)
        if self.ScriptsFolderPath.text() != '':
            for i in reversed(range(self.LayoutVScroll.count())): 
                self.LayoutVScroll.itemAt(i).widget().setParent(None)
            self.select_all = QtWidgets.QCheckBox('SELECT ALL')
            self.select_all.stateChanged.connect(self.check_all_file)
            self.select_all.setStyleSheet("color: #009900; font-weight: bold;")
            self.LayoutVScroll.addWidget(self.select_all)#, alignment=QtCore.Qt.AlignCenter)
            
            list_files = listdir(self.ScriptsFolderPath.text())
            list_files = [file for file in listdir(self.ScriptsFolderPath.text()) if file[-4:]=='.VBS']
            for file in list_files:
                cb = QtWidgets.QCheckBox(file[:-4])
                self.LayoutVScroll.addWidget(cb)
    
    def select_template_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory()
        self.TemplateFolderPath.setText(folder)
    
    def select_excel_file(self):
        file, _ = QtWidgets.QFileDialog.getOpenFileName()
        self.OutputExcelPath.setText(file)

    def select_template_folder_tab2(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory()
        self.TemplateFolderPath_S2.setText(folder)
        if self.TemplateFolderPath_S2.text() != '' and self.ExcelFilePath_S2.text() != '':
            for i in reversed(range(self.LayoutVScroll_S2.count())): 
                self.LayoutVScroll_S2.itemAt(i).widget().setParent(None)
            self.select_all_S2 = QtWidgets.QCheckBox('SELECT ALL')
            self.select_all_S2.setStyleSheet("color: #009900; font-weight: bold;")
            self.select_all_S2.stateChanged.connect(self.check_all_file_S2)
            self.LayoutVScroll_S2.addWidget(self.select_all_S2) #, alignment=QtCore.Qt.AlignCenter)
            script_files = get_matching_scripts(self.TemplateFolderPath_S2.text(), self.ExcelFilePath_S2.text())
            for file in script_files:
                cb = QtWidgets.QCheckBox(file[:-4])
                self.LayoutVScroll_S2.addWidget(cb)

    def select_excel_file_tab2(self):
        file, _ = QtWidgets.QFileDialog.getOpenFileName()
        self.ExcelFilePath_S2.setText(file)
        if self.TemplateFolderPath_S2.text() != '' and self.ExcelFilePath_S2.text() != '':
            for i in reversed(range(self.LayoutVScroll_S2.count())): 
                self.LayoutVScroll_S2.itemAt(i).widget().setParent(None)
            self.select_all_S2 = QtWidgets.QCheckBox('SELECT ALL')
            self.select_all_S2.stateChanged.connect(self.check_all_file_S2)
            self.LayoutVScroll_S2.addWidget(self.select_all_S2, alignment=QtCore.Qt.AlignCenter)
            script_files = get_matching_scripts(self.TemplateFolderPath_S2.text(), self.ExcelFilePath_S2.text())
            for file in script_files:
                cb = QtWidgets.QCheckBox(file)
                self.LayoutVScroll_S2.addWidget(cb)
    
    def check_all_file(self):
        for i in range(self.LayoutVScroll.count()):
            if self.select_all.isChecked():
                self.LayoutVScroll.itemAt(i).widget().setChecked(True)
            else:
                self.LayoutVScroll.itemAt(i).widget().setChecked(False)

    def check_all_file_S2(self):
        for i in range(self.LayoutVScroll_S2.count()):
            if self.select_all_S2.isChecked():
                self.LayoutVScroll_S2.itemAt(i).widget().setChecked(True)
            else:
                self.LayoutVScroll_S2.itemAt(i).widget().setChecked(False)
    
    def handle_script_files(self):
        # Get list of script files need to handle
        list_script_files = []
        for i in range(self.LayoutVScroll.count()):
            wid = self.LayoutVScroll.itemAt(i).widget()
            if(wid.text() != 'SELECT ALL'):
                if wid.isChecked():
                    list_script_files.append(wid.text()+'.VBS')
    
        # Call function to handle
        if self.ScriptsFolderPath.text() == '':
            msg = "Please Select Script Folder!"
            self.show_msg_error(msg)
        elif self.TemplateFolderPath.text() == '':
            msg = "Please Select Template Folder!"
            self.show_msg_error(msg)
        elif self.OutputExcelPath.text() == '':
            msg = "Please Select Excel File!"
            self.show_msg_error(msg)
        else: 
            if -1 == handle_script_folder(list_script_files, self.ScriptsFolderPath.text(), self.TemplateFolderPath.text(), self.OutputExcelPath.text()):
                msg = f'Permission denied to write to {self.OutputExcelPath.text()}'
                self.show_msg_error(msg)
            else:
                msg = "Generate Template Done!"
                self.show_msg_error(msg, title='Success')
                # Change tab
                self.tabWidget.setCurrentIndex(1)

    def run_script_file(self):
        # Get list of script files need to handle
        list_script_files = []
        for i in range(self.LayoutVScroll_S2.count()):
            wid = self.LayoutVScroll_S2.itemAt(i).widget()
            if(wid.text() != 'SELECT ALL'):
                if wid.isChecked():
                    list_script_files.append(wid.text()+'.VBS')

        # Call function to handle
        if self.TemplateFolderPath_S2.text() == '':
            msg = "Please Select Template Folder!"
            self.show_msg_error(msg)
        elif self.ExcelFilePath_S2.text() == '':
            msg = "Please Select Excel File!"
            self.show_msg_error(msg)
        else: 
            if -1 == download_data(self.ExcelFilePath_S2.text(), self.TemplateFolderPath_S2.text(), list_script_files):
                msg = f'Permission denied to write to {self.ExcelFilePath_S2.text()}'
                self.show_msg_error(msg)
            else:
                msg = "Generate Template Done!"
                self.show_msg_error(msg, title='Success')
            # # Change tab
            # self.tabWidget.setCurrentIndex(1)
    
    def show_msg_error(self, msg, title='Error'):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowTitle(title)
        if title == 'Error':
            msg_box.setIcon(QtWidgets.QMessageBox.Critical)
        else:
            msg_box.setIcon(QtWidgets.QMessageBox.Information)
        msg_box.setText(msg)
        msg_box.exec_()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    PhanMemThue = QtWidgets.QMainWindow()
    ui = Ui_PhanMemThue()
    ui.setupUi(PhanMemThue)
    PhanMemThue.show()
    sys.exit(app.exec_())
