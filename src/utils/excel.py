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


def collect_styles(obj):
    if obj.Type in POINTS_TYPE:
        return list(_collect_impl(obj.SeriesCollection(1).Points))
    else:
        return list(_collect_impl(obj.SeriesCollection))


def _collect_impl(method):
    for i in com_range(method().Count):
        obj = method(i).Format.Fill
        yield Style.from_com_object(obj)
