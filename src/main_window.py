from os import path
from typing import Any
from PyQt5.QtWidgets import (
    QAction, QComboBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
    QLabel, QMainWindow, QMessageBox, QPushButton, QTabWidget, QToolBar,
    QVBoxLayout, QWidget)
from PyQt5.QtCore import QSize, Qt
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import QKeySequence
from src.excel_tab_ui import ExeclTab
import sys
import os


class CustomBtn(QPushButton):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setStyleSheet("""
                           width: 100%;
                           color: #fff;
                           background-color: #275efe;
                           padding: 16px 2px;
                           border-radius: 10px;
                           """)


class MainTab(QWidget):
    WELCOMEGROUP = 1
    FORMGROUP = 2
    FIRSTSTANDARD = 3

    def __init__(self, tabs, *args) -> None:
        super().__init__(*args)

        self.tabs = tabs
        self.opt = None

        self.main_layout = QHBoxLayout()
        self.createBar()
        self.createFormGroupBox()
        self.createWelcome()
        self.main_layout.addWidget(self.barGroup)
        self.main_layout.addWidget(self.welcomeGroup)

        self.setLayout(self.main_layout)

    def createWelcome(self):
        self.welcomeGroup = QGroupBox("Welcome")

        self.qb = QVBoxLayout()
        lb1 = QLabel("Welcome")
        self.qb.addWidget(lb1)

        self.welcomeGroup.setLayout(self.qb)

    def createBar(self):
        self.barGroup = QGroupBox('Menu')
        self.btn1 = CustomBtn('Welcome')
        self.btn2 = CustomBtn("Selecting Files")
        self.btn3 = CustomBtn("First Standard")

        self.btn2.clicked.connect(
            lambda: self.handelFormBtn(MainTab.FORMGROUP))
        self.btn1.clicked.connect(
            lambda: self.handelFormBtn(MainTab.WELCOMEGROUP))
        self.btn3.clicked.connect(lambda x: self.handelStandards(x, 1))

        self.hb = QVBoxLayout()
        self.hb.addWidget(self.btn1)
        self.hb.addWidget(self.btn2)
        self.hb.addWidget(self.btn3)
        self.hb.addStretch(1)
        self.hb.setAlignment(Qt.AlignTop)
        self.barGroup.setMaximumWidth(150)

        self.barGroup.setLayout(self.hb)

    def handelFormBtn(self, clearOpt):
        self.clearLayoutAndSetNewOne(clearOpt)

    def handelStandards(self, _, number):
        self.exels_files = []
        self.exels_files.append(str(os.path.join(os.path.dirname(
            __file__), 'data', f'standard_{number}.xlsx')))
        print(self.exels_files)
        for path in self.exels_files:
            if not os.path.exists(path):
                MainWindow.errorHandler(f"directory_{path} not exits")
                return
        
        self.createTabs(None)

    def clearLayoutAndSetNewOne(self, clearOpt):
        if self.opt != clearOpt:
            item = None
            if self.layout() is not None:
                for i in range(1, self.layout().count()):
                    item = self.layout().itemAt(i).widget()
                    item.setParent(None)
            item.close()
            if clearOpt == MainTab.FORMGROUP:
                self.createFormGroupBox()
                self.main_layout.addWidget(self.formGroupBox)
                self.opt = MainTab.FORMGROUP
            elif clearOpt == MainTab.WELCOMEGROUP:
                self.createWelcome()
                self.main_layout.addWidget(self.welcomeGroup)
                self.opt = MainTab.WELCOMEGROUP

    def setComobox(self):
        self.options = {}
        for i in range(1, 11):
            self.options[f'{i} file'] = {
                'value': i,
                'label': f'Insert Address of file {i}'
            }

    def createFormGroupBox(self):
        self.formGroupBox = QGroupBox("Form layout")
        self.formLayout = QFormLayout()
        self.combo = QComboBox(self)
        self.combo.setStyleSheet("""background: #275efe;
                                 border-radius: 10px;
                                 width:100%;""")
        self.combo.addItem('select')
        self.setComobox()

        for val in self.options:
            self.combo.addItem(val, userData=QtCore.QVariant(
                str(self.options[val]['value'])))

        self.filebtn = QPushButton('Select File', self)
        self.filebtn.setStyleSheet("""background: #275efe;
                                   border-radius: 10px;
                                   width: 100%; padding: 2px""")
        self.filebtn.clicked.connect(self.fileBtnHandler)
        self.tabCreatorBtn = QPushButton('create Tabs', self)
        self.tabCreatorBtn.setStyleSheet("""background: #275efe;
                                         border-radius: 10px;
                                         width:100%; padding: 2px;""")

        self.formLayout.addRow(
            QLabel('Enter The Folder Which Contains The Execl Files.'), self.filebtn)
        self.formLayout.addRow(
            QLabel("Enter the number of files you want to load from directory"), self.combo)
        self.formLayout.addRow(
            QLabel('Create Tabs For Execl files.'), self.tabCreatorBtn)

        self.tabCreatorBtn.clicked.connect(self.createTabs)
        self.combo.currentIndexChanged.connect(self.comboHandler)
        self.formGroupBox.setLayout(self.formLayout)

    def fileBtnHandler(self, x):
        """
            Handling directory which selected by user if directory was not selected the send error msg
            in QMessage window if everything was good it sets file_path for directories; 
            parameter:
                x: clicked or not
            returns: 
                void
        """
        self.file_path = str(
            QFileDialog.getExistingDirectory(self, 'select Directory'))

        if self.file_path == None or self.file_path == " " or len(self.file_path) == 0:
            self.file_path = " "
            MainWindow.errorHandler(
                "Directory Not Selected!\n\nplease select directory.")
            return

        self.exels_files = []
        files = os.listdir(self.file_path)

        for file in files:
            if os.path.splitext(file)[1] == ".xlsx":
                self.exels_files.append(path.join(self.file_path, file))

        if len(self.exels_files) == 0:
            MainWindow.errorHandler(
                "Directory Contains No .xlsx files.", "Directory Execl file")
            return

        foundedFiles = "["
        lbl = QLabel()

        for file in self.exels_files:
            foundedFiles += str(file).split('/')[-1] + ', '

        lbl.setText(foundedFiles + ']')
        self.formLayout.addRow(QLabel("Founded Files"), lbl)

    def comboHandler(self, value):
        self.combo_value = value

        if getattr(self, 'exels_files', None) == None or len(self.exels_files) == 0:
            MainWindow.errorHandler("please select Directory First.")
            return

        if value >= len(self.exels_files):
            self.combo_value = len(self.exels_files)

        itemValue = self.combo.itemData(value)
        itemText = self.combo.itemText(value)
        print(itemValue, "====", itemText, "====", value)

    def keyPressEvent(self, a0: QtGui.QKeyEvent) -> None:
        if a0.key() == QKeySequence("Ctrl+Q"):
            sys.exit(1)

        if a0.key() == QKeySequence('Command+Q'):
            sys.exit(1)

    def createTabs(self, _=None):
        if getattr(self, 'exels_files', None) == None or len(self.exels_files) == 0:
            MainWindow.errorHandler("please Select Directory First.")
            return

        if getattr(self, 'combo_value', None) == None:
            self.combo_value = len(self.exels_files)

        for file in self.exels_files[:self.combo_value]:
            tab = ExeclTab(file)
            tab.setMaximumWidth(self.width())
            self.tabs.addTab(tab, str(file).split('/')[-1])
            self.tabs.setCurrentWidget(tab)

        self.exels_files = None
        del self.exels_files
        # self.mainLayout.addWidget(self.tabs)


# this is main Window
class MainWindow(QMainWindow):
    NumGridRows = 3
    NumButtons = 4
    _instance = None

    def __new__(cls, *args, **kwargs) -> Any:
        if not isinstance(MainWindow._instance, cls):
            cls._instance = super().__new__(cls, *args, **kwargs)
        return cls._instance

    @staticmethod
    def getInstance():
        if MainWindow._instance != None:
            return MainWindow._instance
        else:
            MainWindow.errorHandler(
                "Internal Error please Inform Some One!", "Internal Error")

    @staticmethod
    def errorHandler(text, title="Directory Select MessageBox"):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("Warning!")
        msg.setInformativeText(text)
        msg.setWindowTitle(title)
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.exec_()

    def __init__(self, *args) -> None:
        super().__init__(*args)
        # self.setMinimumWidth(2000)
        self.setStyleSheet("background-color: rgba(0,0,0,0.5); color: #fff;")
        self.setWindowTitle('Main Window')
        self.createToolbar()

        self.mainLayout = QVBoxLayout()

        self.tabs = QTabWidget()
        self.mainTab = MainTab(self.tabs)
        self.tabs.addTab(self.mainTab, 'Main Tab')

        self.mainLayout.addWidget(self.tabs)

        widget = QWidget()
        widget.setLayout(self.mainLayout)
        self.setCentralWidget(widget)

    def createToolbar(self):
        self.toolbar = QToolBar('Main Toolbar')
        self.toolbar.setIconSize(QSize(16, 16))

        exit_button = QAction('Exit', self)
        exit_button.setStatusTip('Exit')
        exit_button.triggered.connect(lambda x: sys.exit(1))

        self.toolbar.addAction(exit_button)

        self.addToolBar(self.toolbar)

    def keyPressEvent(self, a0: QtGui.QKeyEvent) -> None:
        if a0.key() == QtCore.Qt.Key_Q:
            sys.exit(1)
