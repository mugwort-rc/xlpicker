# -*- coding: utf-8 -*-

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtCore import QStringListModel
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
        self.pickedModel = models.ChartStyleModel(self)
        self.pickedDelegate = models.ChartPatternDelegate(self)
        self.ui.treeView.setModel(self.pickedModel)
        self.ui.treeView.setItemDelegate(self.pickedDelegate)
        self.targetModel = QStringListModel([
            self.tr("ActiveBook"),
            self.tr("ActiveSheet"),
        ], self)
        self.typeModel = QStringListModel([
            self.tr("All Chart"),
            self.tr("Type of ActiveChart"),
        ], self)
        self.ui.comboBoxTarget.setModel(self.targetModel)
        self.ui.comboBoxType.setModel(self.typeModel)

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
            self.pickedModel.setStyles(styles)
        except pythoncom.com_error:
            QMessageBox.warning(self, self.tr("Collect error"),
                self.tr("Failed collecting from the ActiveChart."))

    @pyqtSlot()
    def on_pushButtonApplyChart_clicked(self):
        styles = self.pickedModel.styles()
        try:
            utils.excel.apply_styles(XL.ActiveChart, styles, prog=self.progObj)
        except pythoncom.com_error:
            QMessageBox.warning(self, self.tr("Apply error"),
                self.tr("Failed applied in the ActiveChart."))

    @pyqtSlot()
    def on_pushButtonApplyBook_clicked(self):
        styles = self.pickedModel.styles()
        target_mode = self.ui.comboBoxTarget.currentIndex()
        target_type = self.ui.comboBoxType.currentIndex()
        self.general_progress.setValue(0)
        self.general_progress.setRange(0, 0)
        self.general_progress.setVisible(True)
        try:
            type_filter = None  # default: All Chart
            if target_type == 1:  # ActiveChart Type
                if XL.ActiveChart is None:
                    QMessageBox.information(self, "", self.tr("Please select the Chart."))
                    return
                type_filter = XL.ActiveChart.ChartType
            book = XL.ActiveWorkbook
            sheets = []
            # ActiveBook
            if target_mode == 0:
                for i in utils.excel.com_range(book.Worksheets.Count):
                    sheets.append(book.Worksheets(i))
            # ActiveSheet
            elif target_mode == 1:
                sheets = [XL.ActiveSheet]
            # calc chart count
            self.on_progress_initialized(len(sheets))
            charts = []
            for c, sheet in enumerate(sheets, 1):
                self.on_progress_updated(c)
                for i in utils.excel.com_range(sheet.ChartObjects().Count):
                    chart = sheet.ChartObjects(i).Chart
                    if type_filter is not None and chart.ChartType != type_filter:
                        continue
                    charts.append(chart)
            self.general_progress.setRange(0, len(charts))
            for i, chart in enumerate(charts):
                self.general_progress.setValue(i)
                try:
                    utils.excel.apply_styles(chart, styles, prog=self.progObj)
                except pythoncom.com_error:
                    chart.Parent.Activate()
                    QMessageBox.warning(self, self.tr("Apply error"),
                        self.tr("Failed applied in the ActiveChart."))

        except pythoncom.com_error:
            QMessageBox.warning(self, self.tr("Apply error"),
                self.tr("Failed applied."))
        self.general_progress.setVisible(False)

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
