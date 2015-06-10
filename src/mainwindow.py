# -*- coding: utf-8 -*-

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QProgressBar
import pythoncom
import win32com.client

XL = win32com.client.Dispatch("Excel.Application")

import models
import utils.excel
from utils.progress import QtProgressObject

from ui.mainwindow import Ui_MainWindow


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # model
        self.currentModel = models.ChartStyleModel(self)
        self.ui.treeView.setModel(self.currentModel)

        # general_progress
        self.general_progress = QProgressBar()
        self.ui.statusbar.addWidget(self.general_progress)
        self.general_progress.setVisible(False)
        # progress & progObj
        self.progress = QProgressBar()
        self.ui.statusbar.addWidget(self.progress)
        self.progress.setVisible(False)
        self.progObj = QtProgressObject()
        self.progObj.initialized.connect(self.on_progress_initialized)
        self.progObj.updated.connect(self.on_progress_updated)
        self.progObj.finished.connect(self.on_progress_finished)

    def getOpenFileName(self, filter=''):
        return QFileDialog.getOpenFileName(self, '', filter)

    def getSaveFileName(self, filter=''):
        return QFileDialog.getSaveFileName(self, '', filter)

    @pyqtSlot()
    def on_pushButtonPickup_clicked(self):
        try:
            styles = utils.excel.collect_styles(XL.ActiveChart,
                                                prog=self.progObj)
            self.currentModel.setStyles(styles)
        except pythoncom.com_error:
            QMessageBox.warning(self, self.tr("Collect error"),
                self.tr("Failed collecting from the ActiveChart."))

    @pyqtSlot()
    def on_pushButtonApplyChart_clicked(self):
        styles = self.currentModel.styles()
        try:
            utils.excel.apply_styles(XL.ActiveChart, styles, prog=self.progObj)
        except pythoncom.com_error:
            QMessageBox.warning(self, self.tr("Apply error"),
                self.tr("Failed applied in the ActiveChart."))

    @pyqtSlot(int)
    def on_progress_initialized(self, maximum):
        self.progress.setValue(0)
        self.progress.setRange(0, maximum)
        self.progress.setVisible(True)

    @pyqtSlot(int)
    def on_progress_updated(self, value):
        self.progress.setValue(value)

    @pyqtSlot()
    def on_progress_finished(self):
        self.progress.setVisible(False)
        self.progress.setValue(0)
