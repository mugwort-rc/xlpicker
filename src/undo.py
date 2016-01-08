# -*- coding: utf-8 -*-

from PyQt4.QtGui import QUndoCommand


class ResetBlock:
    def __init__(self, stack, model):
        self.stack = stack
        self.model = model
        self.before = None

    def __enter__(self):
        self.before = self.model.copy_styles()

    def __exit__(self, types, values, traceback):
        if types is not None:
            return
        after = self.model.copy_styles()
        self.stack.push(StyleResetCommand(self.model, self.before, after))


class StyleResetCommand(QUndoCommand):
    def __init__(self, model, before, after):
        super(StyleResetCommand, self).__init__()
        self.model = model
        self.before = before
        self.after = after

    def undo(self):
        self.model._setStyles(self.before)

    def redo(self):
        self.model._setStyles(self.after)
