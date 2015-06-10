# -*- coding: utf-8 -*-

from PyQt5.QtCore import Qt
from PyQt5.QtCore import QModelIndex
from PyQt5.QtCore import QVariant
from PyQt5.QtGui import QColor
from PyQt5.QtGui import QPixmap
from PyQt5.QtGui import QPainter

from .. import abstract


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

    def data(self, index=QModelIndex(), role=Qt.DisplayRole):
        if role not in [Qt.DisplayRole, Qt.DecorationRole]:
            return QVariant()
        style = self._styles[index.row()]
        fore = QColor(*style.fore)
        back = QColor(*style.back)
        if role == Qt.DecorationRole:
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
