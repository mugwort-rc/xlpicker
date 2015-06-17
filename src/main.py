#! /usr/bin/env python
# -*- coding: utf-8 -*-

import sys

from PyQt4 import QtGui

from mainwindow import MainWindow

def main(argv):
    app = QtGui.QApplication(argv)

    win = MainWindow()
    win.show()

    return app.exec_()

if __name__ == '__main__':
    sys.exit(main(sys.argv))

