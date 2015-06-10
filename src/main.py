#! /usr/bin/env python3
# -*- coding: utf-8 -*-

import sys

from PyQt5 import QtWidgets

from mainwindow import MainWindow

def main(argv):
    app = QtWidgets.QApplication(argv)

    win = MainWindow()
    win.show()

    return app.exec_()

if __name__ == '__main__':
    sys.exit(main(sys.argv))

