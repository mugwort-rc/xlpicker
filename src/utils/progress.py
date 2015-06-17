# -*- coding: utf-8 -*-

from PyQt4.QtCore import pyqtSignal
from PyQt4.QtCore import QObject


class ProgressObject(object):
    def initialize(self, maximum):
        raise NotImplementedError

    def update(self, value):
        raise NotImplementedError

    def finish(self):
        raise NotImplementedError


class QtProgressObject(ProgressObject, QObject):

    initialized = pyqtSignal(int)
    updated = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, parent=None):
        super(QtProgressObject, self).__init__(parent)

    def initialize(self, maximum):
        self.initialized.emit(maximum)

    def update(self, value):
        self.updated.emit(value)

    def finish(self):
        self.finished.emit()
