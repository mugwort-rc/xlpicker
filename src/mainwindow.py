# -*- coding: utf-8 -*-

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QFileDialog
import win32com.client

XL = win32com.client.Dispatch("Excel.Application")

import models
import utils.excel

from ui.mainwindow import Ui_MainWindow

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.currentModel = models.ChartStyleModel(self)
        self.ui.treeView.setModel(self.currentModel)

    def getOpenFileName(self, filter=''):
        return QFileDialog.getOpenFileName(self, '', filter)

    def getSaveFileName(self, filter=''):
        return QFileDialog.getSaveFileName(self, '', filter)

    @pyqtSlot()
    def on_pushButtonPickup_clicked(self):
        self.currentModel.setStyles(
            utils.excel.collect_styles(XL.ActiveChart)
        )
