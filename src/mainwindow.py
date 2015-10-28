# -*- coding: utf-8 -*-

import sys

from PyQt4.QtCore import QT_VERSION_STR

from PyQt4.QtCore import pyqtSlot
from PyQt4.QtCore import Qt
from PyQt4.QtGui import QStringListModel
from PyQt4.QtCore import QModelIndex
from PyQt4.QtGui import QColor
from PyQt4.QtGui import QColorDialog
from PyQt4.QtGui import QDialog
from PyQt4.QtGui import QMainWindow
from PyQt4.QtGui import QMessageBox
from PyQt4.QtGui import QFileDialog
from PyQt4.QtGui import QProgressBar
import pythoncom
import win32com.client

import models
import utils.excel
from utils.progress import QtProgressObject

from ui.mainwindow import Ui_MainWindow


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        try:
            self.XL = win32com.client.Dispatch("Excel.Application")
        except pythoncom.com_error:
            QMessageBox.warning(self, self.tr("Error"), self.tr("Excel isn't installed."))
            sys.exit(0)

        # model
        self.pickedModel = models.ChartStyleModel(self)
        self.pickedDelegate = models.ChartPatternDelegate(self)
        self.ui.treeView.setModel(self.pickedModel)
        self.ui.treeView.setItemDelegate(self.pickedDelegate)
        if QT_VERSION_STR.startswith("5."):
            self.ui.treeView.header().setSectionsMovable(False)
        else:
            self.ui.treeView.header().setMovable(False)
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

    def getActiveChart(self):
        ActiveChart = None
        try:
            ActiveChart = self.XL.ActiveChart
        except AttributeError:
            pass
        if ActiveChart is None:
            QMessageBox.warning(self, self.tr("Error"), self.tr("Chart is not active currently."))
            return None
        return ActiveChart

    @pyqtSlot()
    def on_actionNewStyle_triggered(self):
        self.pickedModel.insertRow(self.pickedModel.rowCount(), QModelIndex())
        self.pickedModel.fetchMore()

    @pyqtSlot()
    def on_actionRemoveStyle_triggered(self):
        current = self.ui.treeView.currentIndex()
        if not current.isValid():
            return
        self.pickedModel.removeRow(current.row(), QModelIndex())

    @pyqtSlot()
    def on_actionUpStyle_triggered(self):
        current = self.ui.treeView.currentIndex()
        if not current.isValid():
            return
        self.pickedModel.upRow(current.row())

    @pyqtSlot()
    def on_actionDownStyle_triggered(self):
        current = self.ui.treeView.currentIndex()
        if not current.isValid():
            return
        self.pickedModel.downRow(current.row())

    @pyqtSlot()
    def on_pushButtonPickup_clicked(self):
        try:
            chart = self.getActiveChart()
            if chart is None:
                return
            styles = utils.excel.collect_styles(chart,
                                                prog=self.progObj)
            self.pickedModel.setStyles(styles)
        except pythoncom.com_error:
            QMessageBox.warning(self, self.tr("Collect error"),
                self.tr("Failed collecting from the ActiveChart."))

    @pyqtSlot()
    def on_pushButtonApplyChart_clicked(self):
        styles = self.pickedModel.styles()
        try:
            chart = self.getActiveChart()
            if chart is None:
                return
            utils.excel.apply_styles(chart, styles, prog=self.progObj)
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
                chart = self.getActiveChart()
                if chart is None:
                    return
                type_filter = chart.ChartType
            book = self.XL.ActiveWorkbook
            if book is None:
                raise pythoncom.com_error
            sheets = []
            # ActiveBook
            if target_mode == 0:
                for i in utils.excel.com_range(book.Worksheets.Count):
                    sheets.append(book.Worksheets(i))
            # ActiveSheet
            elif target_mode == 1:
                sheets = [self.XL.ActiveSheet]
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
            aborted = False
            for i, chart in enumerate(charts):
                if aborted:
                    break
                self.general_progress.setValue(i)
                try:
                    utils.excel.apply_styles(chart, styles, prog=self.progObj)
                except pythoncom.com_error:
                    chart.Parent.Activate()
                    buttons = QMessageBox.Ok | QMessageBox.Abort
                    ret = QMessageBox.warning(self, self.tr("Apply error"),
                        self.tr("Failed applied in the ActiveChart."), buttons)
                    if ret == QMessageBox.Abort:
                        aborted = True

        except pythoncom.com_error:
            QMessageBox.warning(self, self.tr("Apply error"),
                self.tr("Failed applied."))
        self.general_progress.setVisible(False)

    @pyqtSlot(QModelIndex)
    def on_treeView_doubleClicked(self, index):
        if index.column() not in [1, 2, 4]:
            return
        color = QColor(index.data(Qt.EditRole))
        dialog = QColorDialog(self)
        dialog.setCurrentColor(color)
        if dialog.exec_() == QDialog.Accepted:
            self.pickedModel.setData(index, dialog.currentColor())

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
