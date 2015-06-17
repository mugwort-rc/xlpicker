#!/usr/bin/env bash
pyuic4 src/ui/mainwindow.ui -o src/ui/mainwindow.py
pyrcc4 src/resource.qrc -o src/resource_rc.py
