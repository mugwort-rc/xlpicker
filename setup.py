import py2exe
from distutils.core import setup

py2exe_options = {
    "compressed": 1,
    "optimize": 2,
    "bundle_files": 1,
    "includes" : ["sip", "PyQt4.QtCore", "PyQt4.QtGui",
                  "src.utils.constants", "src.utils.excel",
                  "src.utils.progress"]
}
setup(
    options={"py2exe" : py2exe_options},
    windows=[{"script" : "main.py",
              "dest_base": "xlpicker"}],
    zipfile=None,
)
