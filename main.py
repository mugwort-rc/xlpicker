#! /usr/bin/env python
# -*- coding: utf-8 -*-

import sys

from PyQt4 import QtCore
from PyQt4 import QtGui

from src.mainwindow import MainWindow

def main(argv):
    app = QtGui.QApplication(argv)

    qtTr = QtCore.QTranslator()
    if qtTr.load("qt_" + QtCore.QLocale.system().name(), ":/i18n/"):
        app.installTranslator(qtTr)

    xlpickerTr = QtCore.QTranslator()
    if xlpickerTr.load("xlpicker_" + QtCore.QLocale.system().name(), ":/i18n/"):
        app.installTranslator(xlpickerTr)

    win = MainWindow()
    win.show()

    return app.exec_()

if __name__ == '__main__':
    sys.exit(main(sys.argv))

