# -*- coding: utf-8 -*-

import copy

from PyQt5.QtCore import Qt
from PyQt5.QtCore import QModelIndex
from PyQt5.QtCore import QVariant
from PyQt5.QtCore import QStringListModel
from PyQt5.QtGui import QColor
from PyQt5.QtGui import QPixmap
from PyQt5.QtGui import QPainter
from PyQt5.QtWidgets import QComboBox
from PyQt5.QtWidgets import QItemDelegate
from PyQt5.QtWidgets import QUndoStack

from . import abstract
from ..undo import ResetBlock
from ..utils.excel import Style


def qcolor2rgb(color):
    return (color.red(), color.green(), color.blue())


class ChartStyleModel(abstract.AbstractItemModel):
    def __init__(self, parent=None):
        super(ChartStyleModel, self).__init__(parent)
        self._styles = []
        self.patternImage = QPixmap(":/img/patterns.png")
        self.setItemSize(0, 6)  # 6: Pattern, Fore, Back, LineStyle, Color, Weight
        self._undoStack = QUndoStack()

    def setStyles(self, styles):
        """
        :type styles: list[utils.excel.Style]
        """
        self.beginResetModel()
        with ResetBlock(self._undoStack, self):
            self._setStyles(styles)
        self.endResetModel()

    def _setStyles(self, styles):
        """
        :type styles: list[utils.excel.Style]
        """
        self.beginResetModel()
        self._styles = styles
        self.setItemSize(len(self._styles),
                         # Pattern, Fore, Back, LineStyle, Color, Weight
                         6)
        self.endResetModel()

    def styles(self):
        return self._styles

    def copy_styles(self):
        return [x.copy() for x in self.styles()]

    def undoStack(self):
        return self._undoStack

    def setUndoStack(self, undoStack):
        self._undoStack = undoStack

    def replicate(self, index):
        self.beginResetModel()
        with ResetBlock(self._undoStack, self):
            newobj = copy.deepcopy(self._styles[index])
            self._styles.append(newobj)
            self.setItemSize(len(self._styles),
                             # Pattern, Fore, Back, LineStyle, Color, Weight
                             6)
        self.endResetModel()

    def flags(self, index):
        result = super(ChartStyleModel, self).flags(index)
        if index.column() in [0, 3, 5]:
            return result | Qt.ItemIsEditable
        return result

    def data(self, index=QModelIndex(), role=Qt.DisplayRole):
        if role not in [Qt.DisplayRole, Qt.DecorationRole, Qt.EditRole]:
            return QVariant()
        style = self._styles[index.row()]
        fore = QColor(*style.fore)
        back = QColor(*style.back)
        border_color = QColor(*style.color)
        if role == Qt.EditRole:
            if index.column() == 0:
                if style.isFilled():
                    return 0  # solid special case
                return style.pattern
            elif index.column() == 1:
                return fore
            elif index.column() == 2:
                return back
            elif index.column() == 3:
                if style.style == -4105:
                    return 0  # Auto special case
                elif style.style == -4142:
                    return 9  # None special case
                return style.dash  # see also MsoLineDashStyle
            elif index.column() == 4:
                return border_color
            elif index.column() == 5:
                if style.weight == -4138:
                    return 3  # Medium special case
                return style.weight  # see also XlBorderWeight
            else:
                return QVariant()
        elif role == Qt.DecorationRole:
            # Pattern
            if index.column() == 0:
                img = None
                if style.isFilled():
                    img = QPixmap(48, 48)
                    img.fill(Qt.black)
                else:
                    img = self.patternImage.copy(48*(style.pattern-1), 0, 48, 48)
                foreMask = img.createMaskFromColor(QColor(Qt.white))
                backMask = img.createMaskFromColor(QColor(Qt.black))
                p = QPainter(img)
                p.setPen(fore)
                p.drawPixmap(img.rect(), foreMask, foreMask.rect())
                p.setPen(back)
                p.drawPixmap(img.rect(), backMask, backMask.rect())
                p.end()
                return img
        elif role == Qt.DisplayRole:
            # Line style
            if index.column() == 3:
                value = self.data(index, role=Qt.EditRole)
                if value not in range(10):
                    return self.tr("Unknown")
                return {
                    0: self.tr("Auto"),
                    1: self.tr("Solid"),
                    2: self.tr("SquareDot"),
                    3: self.tr("RoundDot"),
                    4: self.tr("LineDash"),
                    5: self.tr("DashDot"),
                    6: self.tr("DashDotDot"),
                    7: self.tr("LongDash"),
                    8: self.tr("LongDashDot"),
                    9: self.tr("None"),
                }[value]
            elif index.column() == 5:
                return {
                    1: self.tr("Hairline"),
                    2: self.tr("Thin"),
                    3: self.tr("Medium"),
                    4: self.tr("Thick"),
                }[self.data(index, role=Qt.EditRole)]
        # Fore
        if index.column() == 1:
            return fore
        # Back
        elif index.column() == 2:
            return back
        # border color
        elif index.column() == 4:
            return border_color
        return QVariant()

    def setData(self, index, value, role=Qt.EditRole):
        if role == Qt.EditRole:
            with ResetBlock(self._undoStack, self):
                if index.column() == 0:
                    if isinstance(value, QVariant):
                        value = value.toInt()[0]
                    self._styles[index.row()].pattern = value
                elif index.column() == 1:
                    if isinstance(value, QVariant):
                        value = QColor(value)
                    self._styles[index.row()].fore = qcolor2rgb(value)
                elif index.column() == 2:
                    if isinstance(value, QVariant):
                        value = QColor(value)
                    self._styles[index.row()].back = qcolor2rgb(value)
                elif index.column() == 3:
                    if isinstance(value, QVariant):
                        value = value.toInt()[0]
                    if value in [0, 9]:
                        if value == 0:
                            self._styles[index.row()].style = -4105
                        elif vlaue == 9:
                            self._styles[index.row()].style = -4142
                        self._styles[index.row()].dash = 1
                    else:
                        self._styles[index.row()].style = {
                            1: 1,     # Solid
                            2: -4118, # SquareDot
                            3: -4118, # RoundDot
                            4: -4115, # LineDash
                            5: 4,     # DashDot
                            6: 5,     # DashDotDot
                            7: -4115, # LongDash
                            8: 4,     # LongDashDot
                        }[value]
                        self._styles[index.row()].dash = value
                elif index.column() == 4:
                    if isinstance(value, QVariant):
                        value = QColor(value)
                    self._styles[index.row()].color = qcolor2rgb(value)
                elif index.column() == 5:
                    if isinstance(value, QVariant):
                        value = value.toInt()[0]
                    if value == 3:
                        value = -4138
                    self._styles[index.row()].weight = value
            return True
        return False

    def headerData(self, section, orientation, role):
        if role != Qt.DisplayRole or orientation == Qt.Vertical:
            return QVariant()
        if section == 0:
            return self.tr("Pattern")
        elif section == 1:
            return self.tr("Fore")
        elif section == 2:
            return self.tr("Back")
        elif section == 3:
            return self.tr("Border")
        elif section == 4:
            return self.tr("Border color")
        elif section == 5:
            return self.tr("Border weight")
        return QVariant()

    def insertRows(self, row, count, parent=QModelIndex()):
        self.beginInsertRows(parent, row, row+count)
        with ResetBlock(self._undoStack, self):
            for i in range(count):
                style = Style.create()
                if row == len(self._styles):
                    self._styles.append(style)
                else:
                    self._styles.insert(row+i, style)
            self.row.insert(row, count)
        self.endInsertRows()
        return True

    def removeRows(self, row, count, parent=QModelIndex()):
        self.beginRemoveRows(parent, row, row+count)
        with ResetBlock(self._undoStack, self):
            for i in range(count):
                if len(self._styles) <= row:
                    return False
                del self._styles[row]
            self.row.remove(row, count)
        self.endRemoveRows()
        return True

    def upRow(self, row):
        if row < 1:
            return
        if len(self._styles) <= row:
            return
        self.beginMoveRows(QModelIndex(), row, row, QModelIndex(), row-1)
        with ResetBlock(self._undoStack, self):
            self._styles.insert(row-1, self._styles.pop(row))
        self.endMoveRows()

    def downRow(self, row):
        if row < 0:
            return
        if len(self._styles) <= row+1:
            return
        with ResetBlock(self._undoStack, self):
            self.upRow(row+1)

    def parent(self, index):
        return QModelIndex()

    def hasChildren(self, parent):
        if parent.isValid():
            return False
        # root item
        return True


class ChartPatternModel(abstract.AbstractListModel):
    def __init__(self, parent=None):
        super(ChartPatternModel, self).__init__(parent)
        self.patternImage = QPixmap(":/img/patterns.png")
        self._foreColor = QColor(Qt.black)
        self._backColor = QColor(Qt.white)

        self.beginResetModel()
        self.setRowSize(48+1)  # patterns + one solid
        self.endResetModel()

    def setForeColor(self, color):
        self._foreColor = color

    def setBackColor(self, color):
        self._backColor = color

    def foreColor(self):
        return self._foreColor

    def backColor(self):
        return self._backColor

    def data(self, index=QModelIndex(), role=Qt.DisplayRole):
        if role not in [Qt.DecorationRole, Qt.EditRole]:
            return QVariant()
        if role == Qt.DecorationRole:
            img = None
            # solid special case
            if index.row() == 0:
                img = QPixmap(48, 48)
                img.fill(Qt.black)
            # pattern case
            else:
                img = self.patternImage.copy(48*(index.row()-1), 0, 48, 48)
            foreMask = img.createMaskFromColor(QColor(Qt.white))
            backMask = img.createMaskFromColor(QColor(Qt.black))
            p = QPainter(img)
            p.setPen(self.foreColor())
            p.drawPixmap(img.rect(), foreMask, foreMask.rect())
            p.setPen(self.backColor())
            p.drawPixmap(img.rect(), backMask, backMask.rect())
            p.end()
            return img
        elif role == Qt.EditRole:
            if index.row() == 0:
                return -2  # solid special case
            else:
                return index.row()
        return QVariant()


class ChartPatternDelegate(QItemDelegate):
    def __init__(self, parent=None):
        super(ChartPatternDelegate, self).__init__(parent)

        self.lineStrings = [
            self.tr("Auto"),
            self.tr("Solid"),
            self.tr("SquareDot"),
            self.tr("RoundDot"),
            self.tr("LineDash"),
            self.tr("DashDot"),
            self.tr("DashDotDot"),
            self.tr("LongDash"),
            self.tr("LongDashDot"),
            self.tr("None"),
        ]
        self.weightStrings = [
            self.tr("Hairline"),
            self.tr("Thin"),
            self.tr("Medium"),
            self.tr("Thick"),
        ]

    def createEditor(self, parent, option, index):
        if index.column() == 0:
            editor = QComboBox(parent)
            editorDelegate = PatternImageDelegate(editor)
            editor.setItemDelegate(editorDelegate)
            model = ChartPatternModel(editor)
            #model.setForeColor(index.sibling(index.row(), 1).data(Qt.EditRole))
            #model.setBackColor(index.sibling(index.row(), 2).data(Qt.EditRole))
            model.fetchMore()
            editor.setModel(model)
            return editor
        elif index.column() == 3:
            editor = QComboBox(parent)
            model = QStringListModel(editor)
            model.setStringList(self.lineStrings)
            editor.setModel(model)
            return editor
        elif index.column() == 5:
            editor = QComboBox(parent)
            model = QStringListModel(editor)
            model.setStringList(self.weightStrings)
            editor.setModel(model)
            return editor
        return super(ChartPatternDelegate, self).createEditor(parent, option, index)

    def setEditorData(self, editor, index):
        data = index.model().data(index, Qt.EditRole)
        if index.column() == 0:
            if data < 1 or data > 48:
                data = 0
        elif index.column() == 5:
            data -= 1
        editor.setCurrentIndex(data)

    def setModelData(self, editor, model, index):
        value = editor.currentIndex()
        if index.column() == 0:
            if value == 0:
                value = -2
        elif index.column() == 5:
            value += 1
        model.setData(index, value, Qt.EditRole)


class PatternImageDelegate(QItemDelegate):
    def paint(self, painter, option, index):
        pixmap = QPixmap(index.data(Qt.DecorationRole))
        painter.drawPixmap(option.rect, pixmap)
        super(PatternImageDelegate, self).paint(painter, option, index)

    def sizeHint(self, option, index):
        pixmap = QPixmap(index.data(Qt.DecorationRole))
        return pixmap.size()
