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
    def __init__(self, pattern, fore, back, style, color, weight):
        # format.fill
        self.pattern = pattern
        self.fore = fore
        self.back = back
        # border
        self.style = style
        self.color = color
        self.weight = weight

    def isFilled(self):
        return self.pattern <= 0 or self.pattern > 48

    @staticmethod
    def create():
        # Format.Fill: None, black, white
        return Style(-2, lergb2rgb(0), lergb2rgb(0xffffff),
        # Border: Auto, None, None
                      -4105, -2, -4142)

    @staticmethod
    def from_com_object(obj):
        return Style(
            obj.Format.Fill.Pattern,
            lergb2rgb(obj.Format.Fill.ForeColor.RGB),
            lergb2rgb(obj.Format.Fill.BackColor.RGB),
            obj.Border.LineStyle,
            obj.Border.Color,
            obj.Border.Weight,
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


def collect_styles(chart, prog=None):
    if chart.ChartType in POINTS_TYPE:
        return list(_collect_style(chart.SeriesCollection(1).Points, prog=prog))
    else:
        return list(_collect_style(chart.SeriesCollection, prog=prog))


def _collect_style(method, prog=None):
    maximum = method().Count
    if prog:
        prog.initialize(maximum)
    for i in com_range(maximum):
        if prog:
            prog.update(i)
        obj = method(i)
        yield Style.from_com_object(obj)
    if prog:
        prog.finish()


def apply_styles(chart, styles, prog=None):
    if chart.ChartType in POINTS_TYPE:
        _apply_style(chart.SeriesCollection(1).Points, styles, prog=prog)
    else:
        _apply_style(chart.SeriesCollection, styles, prog=prog)


def _apply_style(method, styles, prog=None):
    maximum = min(method().Count, len(styles))
    if prog:
        prog.initialize(maximum)
    for i in com_range(maximum):
        if prog:
            prog.update(i)
        style = styles[i-1]  # 0-based
        obj = method(i)
        # Format.Fill
        if style.isFilled():
            obj.Format.Fill.Solid()
            obj.Format.Fill.ForeColor.RGB = rgb2lergb(*style.fore)
        else:
            obj.Format.Fill.Patterned(style.pattern)
            obj.Format.Fill.ForeColor.RGB = rgb2lergb(*style.fore)
            obj.Format.Fill.BackColor.RGB = rgb2lergb(*style.back)
        # Border
        obj.Border.LineStyle = style.style
        if style.style not in [-4142, -4105]:
            obj.Border.Color = style.color
            obj.Border.Weight = style.weight
    if prog:
        prog.finish()
