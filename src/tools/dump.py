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
    colorCollectors = [
        excel.RGBCollector(),
        excel.SchemeColorCollector(),
    ]

    collector = excel.SeriesCollectionCollector(
        excel.SeriesCollector([
            excel.FormatCollector([
                excel.FillCollector([
                    excel.BackColorCollector(colorCollectors),
                    excel.ForeColorCollector(colorCollectors),
                    excel.PatternCollector(),
                ])
            ])
        ])
    )
    print(yaml.safe_dump(collector.collect(XL.ActiveChart)))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
