#! /usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import print_function

import argparse
import sys
import os.path

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import win32com.client
import yaml

import excel


def main(argv):
    parser = argparse.ArgumentParser()
    args = parser.parse_args(argv[1:])

    XL = win32com.client.Dispatch("Excel.Application")
    if XL.ActiveChart is None:
        print("ActiveChart is None", file=sys.stderr)
        return 1

    colorLambda = lambda x: "#{:02x}{:02x}{:02x}".format(x&0xff, (x>>8)&0xff, (x>>16)&0xff) if x >= 0 else x
    colorCollector = {
        "RGB": colorLambda,
    }
    formatCollector = {
        "Fill": {
            "BackColor": colorCollector,
            "ForeColor": colorCollector,
            "Type": None,
            "Visible": None,
        },
    }
    collector = {
        "SeriesCollection": (excel.COLLECT_METHOD, {
            "Format": formatCollector,
            "Points": (excel.COLLECT_METHOD, {
                "Format": formatCollector,
            }),
            "Border": {
                "Color": colorLambda,
                "ColorIndex": None,
                "LineStyle": None,
                "Weight": None,
            },
        }),
    }
    print(yaml.safe_dump(excel.collect(XL.ActiveChart, collector)))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
