#!/usr/bin/env bash
pyuic5 src/ui/mainwindow.ui -o src/ui/mainwindow.py
pyrcc5 src/resource.qrc -o src/resource_rc.py
