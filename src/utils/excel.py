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
    # doughnut charts
    win32com.client.constants.xlDoughnut,
    win32com.client.constants.xlDoughnutExploded,
]

class Style(object):
    def __init__(self, pattern, fore, back, style, dash, line, color, weight):
        # format.fill
        self.pattern = pattern
        self.fore = fore
        self.back = back
        # border
        self.style = style  # Border style
        self.dash = dash  # Dash style
        self.line = line  # Line style
        self.color = color
        self.weight = weight
        # TODO: marker option for LineMarkers

    def isFilled(self):
        return self.pattern <= 0 or self.pattern > 48

    def copy(self):
        ret = Style(self.pattern, self.fore, self.back,
                    self.style, self.dash, self.line, self.color, self.weight)
        return ret

    def dump(self):
        return {
            "pattern": self.pattern,
            "fore": rgb2lergb(*self.fore),
            "back": rgb2lergb(*self.back),
            "style": self.style,
            "dash": self.dash,
            "line": self.line,
            "color": rgb2lergb(*self.color),
            "weight": self.weight,
        }

    @staticmethod
    def create():
        # Format.Fill: None, black, white
        return Style(-2, lergb2rgb(0), lergb2rgb(0xffffff),
        # Border: Auto, Solid, Single, None, Normal
                      -4105, 1, 1, lergb2rgb(0), -4138)

    @staticmethod
    def from_com_object(obj):
        return Style(
            obj.Format.Fill.Pattern,
            lergb2rgb(obj.Format.Fill.ForeColor.RGB),
            lergb2rgb(obj.Format.Fill.BackColor.RGB),
            obj.Border.LineStyle,
            obj.Format.Line.DashStyle,
            obj.Format.Line.Style,
            lergb2rgb(obj.Border.Color) if obj.Border.Color >= 0 else lergb2rgb(0),
            obj.Border.Weight,
        )

    @staticmethod
    def from_dump(data):
        return Style(data["pattern"], lergb2rgb(data["fore"]), lergb2rgb(data["back"]),
                        data["style"], data["dash"], data["line"], lergb2rgb(data["color"]), data["weight"])


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
        if style.style not in [-4142, -4105]:  # None, Auto
            obj.Format.Line.DashStyle = style.dash
            if style.line != -2:  # not supported
                obj.Format.Line.Style = style.line
            obj.Border.Color = rgb2lergb(*style.color)
            obj.Border.Weight = style.weight
    if prog:
        prog.finish()
