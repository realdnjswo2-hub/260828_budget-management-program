"""Collect the local Tcl/Tk runtime without executing tkinter during analysis."""

import os
import sys


_root = sys.base_prefix
binaries = []
datas = []

for _name in ("tcl86t.dll", "tk86t.dll"):
    _source = os.path.join(_root, "DLLs", _name)
    if os.path.exists(_source):
        binaries.append((_source, "."))

for _folder, _destination in (("tcl8.6", "_tcl_data"), ("tk8.6", "_tk_data")):
    _source = os.path.join(_root, "tcl", _folder)
    if os.path.isdir(_source):
        datas.append((_source, _destination))
