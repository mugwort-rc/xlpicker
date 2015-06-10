# -*- coding: utf-8 -*-

from . import constants
import win32com.client


POINTS_TYPE = [
    # pie charts
    win32com.client.constants.xl3DPie,
    win32com.client.constants.xl3DPieExploded,
    win32com.client.constants.xlBarOfPie,
    win32com.client.constants.xlPie,
    win32com.client.constants.xlPieExploded,
    win32com.client.constants.xlPieOfPie,
]

class Style(object):
    def __init__(self, pattern, fore, back):
        self.pattern = pattern
        self.fore = fore
        self.back = back

    def isFilled(self):
        return self.pattern <= 0 or self.pattern > 48

    @staticmethod
    def from_com_object(obj):
        return Style(
            obj.Pattern,
            lergb2rgb(obj.ForeColor.RGB),
            lergb2rgb(obj.BackColor.RGB)
        )


def com_range(n):
    return range(1, n+1)


def lergb2rgb(lergb):
    r = lergb & 0xff
    g = (lergb >> 8) & 0xff
    b = (lergb >> 16) & 0xff
    return (r, g, b)


def rgb2lergb(r, g, b):
    return (
        (b << 16) |
        (g << 8) |
        r
    )


def collect_styles(chart):
    if chart.ChartType in POINTS_TYPE:
        return list(_collect_fill(chart.SeriesCollection(1).Points))
    else:
        return list(_collect_fill(chart.SeriesCollection))


def _collect_fill(method):
    for i in com_range(method().Count):
        obj = method(i).Format.Fill
        yield Style.from_com_object(obj)


def apply_styles(chart, styles):
    if chart.ChartType in POINTS_TYPE:
        _apply_fill(chart.SeriesCollection(1).Points, styles)
    else:
        _apply_fill(chart.SeriesCollection, styles)


def _apply_fill(method, styles):
    maximum = min(method().Count, len(styles))
    for i in com_range(maximum):
        style = styles[i-1]  # 0-based
        obj = method(i).Format.Fill
        if style.isFilled():
            obj.Solid()
            obj.ForeColor.RGB = rgb2lergb(*style.fore)
            #obj.BackColor.RGB = rgb2lergb(0xff, 0xff, 0xff)
        else:
            obj.Patterned(style.pattern)
            obj.ForeColor.RGB = rgb2lergb(*style.fore)
            obj.BackColor.RGB = rgb2lergb(*style.back)
