# -*- coding: utf-8 -*-


def excel_range(n):
    return range(1, n+1)


class Collector(object):
    def __init__(self, name):
        """
        :type name: str
        """
        self.name = name

    def collect(self, parent):
        """
        :type parent: COMObject
        :return: collected data
        """
        raise NotImplementedError

    @property
    def name(self):
        """
        :rtype: str
        """
        return self._name

    @name.setter
    def name(self, name):
        """
        :type name: str
        """
        self._name = name


class SequenceCollector(Collector):
    def __init__(self, name, collector):
        """
        :type name: str
        :type collector: Collector
        """
        super(SequenceCollector, self).__init__(name)
        self.collector = collector

    def collect(self, parent):
        """
        :type parent: COMObject
        :return: collected data
        """
        return [self.collector.collect(child) for child in self.childrens(parent)]

    def childrens(self, parent):
        """
        :type parent: COMObject
        :rtype: collections.Iterable[COMObject]
        """
        raise NotImplementedError


class AttributeSequenceCollector(SequenceCollector):

    ATTRIBUTE_NAME = None

    def __init__(self, collector):
        if self.ATTRIBUTE_NAME is None:
            raise NotImplementedError
        super(AttributeSequenceCollector, self).__init__(self.ATTRIBUTE_NAME, collector)

    def childrens(self, parent):
        attrib = getattr(parent, self.ATTRIBUTE_NAME)
        for i in excel_range(attrib.Count):
            yield attrib(i)

class MethodSequenceCollector(AttributeSequenceCollector):
    def childrens(self, parent):
        attrib = getattr(parent, self.ATTRIBUTE_NAME)
        for i in excel_range(attrib().Count):
            yield attrib(i)


class AttributeCollector(Collector):

    ATTRIBUTE_NAME = None

    def __init__(self):
        if self.ATTRIBUTE_NAME is None:
            raise NotImplementedError
        super(AttributeCollector, self).__init__(self.ATTRIBUTE_NAME)

    def collect(self, parent):
        raise NotImplementedError

class AttributeValueCollector(AttributeCollector):
    def collect(self, parent):
        return getattr(parent, self.ATTRIBUTE_NAME)


class SetCollector(AttributeCollector):
    def __init__(self, collectors):
        super(SetCollector, self).__init__()
        self.collectors = collectors

class ObjectCollector(SetCollector):
    def collect(self, parent):
        # directly
        return dict([(c.name, c.collect(parent)) for c in self.collectors])

class PropertyCollector(SetCollector):
    def collect(self, parent):
        obj = getattr(parent, self.ATTRIBUTE_NAME)
        return dict([(c.name, c.collect(obj)) for c in self.collectors])

class MethodCollector(SetCollector):
    def collect(self, parent):
        obj = getattr(parent, self.ATTRIBUTE_NAME)()
        return dict([(c.name, c.collect(obj)) for c in self.collectors])


# Values

class BrightnessCollector(AttributeValueCollector):
    ATTRIBUTE_NAME = "Brightness"

class ObjectThemeColorCollector(AttributeValueCollector):
    ATTRIBUTE_NAME = "ObjectThemeColor"

class PatternCollector(AttributeValueCollector):
    ATTRIBUTE_NAME = "Pattern"

class RGBCollector(AttributeValueCollector):
    ATTRIBUTE_NAME = "RGB"

class SchemeColorCollector(AttributeValueCollector):
    ATTRIBUTE_NAME = "SchemeColor"

class TypeCollector(AttributeValueCollector):
    ATTRIBUTE_NAME = "Type"


class ColorFormatCollector(PropertyCollector):
    pass

class BackColorCollector(ColorFormatCollector):
    ATTRIBUTE_NAME = "BackColor"

class ForeColorCollector(ColorFormatCollector):
    ATTRIBUTE_NAME = "ForeColor"


class FillFormatCollector(PropertyCollector):
    pass

class FillCollector(FillFormatCollector):
    ATTRIBUTE_NAME = "Fill"

class ChartFormatCollector(PropertyCollector):
    pass

class FormatCollector(ChartFormatCollector):
    ATTRIBUTE_NAME = "Format"


# Charts

class SeriesCollector(ObjectCollector):
    ATTRIBUTE_NAME = "Series"  # DUMMY

class SeriesCollectionCollector(MethodSequenceCollector):
    ATTRIBUTE_NAME = "SeriesCollection"


# books

class WorkbooksCollector(AttributeSequenceCollector):
    ATTRIBUTE_NAME = "Workbooks"

class WorksheetsCollector(AttributeSequenceCollector):
    ATTRIBUTE_NAME = "Worksheets"

