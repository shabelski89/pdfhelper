"""
Provides to convert docx file(s) into pdf using docx2pdf
"""
__author__ = "Aleksandr Shabelsky"
__version__ = "1.0.0"
__email__ = "a.shabelsky@gmail.com"
# Requirement docx2pdf: pip install docx2pdf
# Usage python docx_to_pdf.py -i file1.docx file2.docx

import wx
from wxgui.main_widget import MainFrame
import multiprocessing


class GenApp(wx.App):
    def __init__(self, redirect=False, filename=None):
        wx.App.__init__(self, redirect, filename)

    def OnInit(self):
        frame = MainFrame()
        frame.Center()
        frame.Show()
        return True


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = GenApp(redirect=False)
    app.MainLoop()
