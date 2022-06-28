"""
This module contains the file input GUI.
"""


# Third-party modules:
import wx

# Module functions:
from declawb import *


# GUI:
app = wx.App()

filename = wx.FileSelector("Choose an Excel file", wildcard="*.xlsx")
if check_excel_file(filename):
    create_text_file(filename, get_declawb(filename))

app.MainLoop()
