# -*- coding: utf-8 -*-

from PyQt5.QtCore import Qt
from PyQt5.QtCore import QModelIndex
from PyQt5.QtCore import QVariant
from PyQt5.QtGui import QColor
from PyQt5.QtGui import QPixmap
from PyQt5.QtGui import QPainter
from PyQt5.QtWidgets import QComboBox
from PyQt5.QtWidgets import QItemDelegate

from .. import abstract


def qcolor2rgb(color):
    return (color.red(), color.green(), color.blue())


class ChartStyleModel(abstract.AbstractItemModel):
    def __init__(self, parent=None):
        super(ChartStyleModel, self).__init__(parent)
        self._styles = []
        self.patternImage = QPixmap(":/img/patterns.png")

    def setStyles(self, styles):
        """
        :type styles: list[utils.excel.Style]
        """
        self.beginResetModel()
        self._styles = styles
        self.setItemSize(len(self._styles),
                         # Pattern, Fore, Back
                         3 if len(self._styles) > 0 else 0)
        self.endResetModel()

    def styles(self):
        return self._styles

    def flags(self, index):
        result = super(ChartStyleModel, self).flags(index)
        if index.column() == 0:
            return result | Qt.ItemIsEditable
        return result

    def data(self, index=QModelIndex(), role=Qt.DisplayRole):
        if role not in [Qt.DisplayRole, Qt.DecorationRole, Qt.EditRole]:
            return QVariant()
        style = self._styles[index.row()]
        fore = QColor(*style.fore)
        back = QColor(*style.back)
        if role == Qt.EditRole:
            if index.column() == 0:
                if style.isFilled():
                    return 0  # solid special case
                return style.pattern
            elif index.column() == 1:
                return fore
            elif index.column() == 2:
                return back
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
        # Fore
        if index.column() == 1:
            return fore
        # Back
        elif index.column() == 2:
            return back
        return QVariant()

    def setData(self, index, value, role=Qt.EditRole):
        if role == Qt.EditRole:
            if index.column() == 0:
                self._styles[index.row()].pattern = value
            elif index.column() == 1:
                self._styles[index.row()].fore = qcolor2rgb(value)
            elif index.column() == 2:
                self._styles[index.row()].back = qcolor2rgb(value)
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
        return super(ChartPatternDelegate, self).createEditor(parent, option, index)

    def setEditorData(self, editor, index):
        pattern = index.model().data(index, Qt.EditRole)
        if pattern < 1 or pattern > 48:
            pattern = 0
        editor.setCurrentIndex(pattern)

    def setModelData(self, editor, model, index):
        pattern = editor.currentIndex()
        if pattern == 0:
            pattern = -2
        model.setData(index, pattern, Qt.EditRole)


class PatternImageDelegate(QItemDelegate):
    def paint(self, painter, option, index):
        pixmap = index.data(Qt.DecorationRole)
        painter.drawPixmap(option.rect, pixmap)
        super(PatternImageDelegate, self).paint(painter, option, index)

    def sizeHint(self, option, index):
        pixmap = index.data(Qt.DecorationRole)
        return pixmap.size()
