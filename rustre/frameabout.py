import platform

import wx
import pandas as pd
import openpyxl as op

from rustre.bitmap import rustre_icon
from rustre.version import BRANCH_NAME
from rustre.version import COMMIT_ID
from rustre.version import COMMIT_NUMBER
from rustre.version import VERSION_MAJOR_MINOR



class FrameAbout(wx.Dialog):  # pragma: no cover

    def __init__(self, parent, programname):
        wx.Dialog.__init__(self, parent, id=wx.ID_ANY, title=u"About", pos=wx.DefaultPosition, size=wx.DefaultSize,
                           style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        self._create_controls(programname + " v{}".format(VERSION_MAJOR_MINOR))

        # set version number
        self.m_textCtrl3.AppendText("Commit id: " + COMMIT_ID + "\n")
        self.m_textCtrl3.AppendText("Commit number: " + COMMIT_NUMBER + "\n")
        self.m_textCtrl3.AppendText("Branch: " + BRANCH_NAME + "\n")
        self.m_textCtrl3.AppendText("Openpyxl: " + op.__version__ + "\n")
        self.m_textCtrl3.AppendText("Pandas: " + pd.__version__ + "\n")
        self.m_textCtrl3.AppendText("Python: " + platform.python_version() + "\n")
        self.m_textCtrl3.AppendText("wxWidgets: " + wx.version())
        self.m_btn_cancel.SetFocus()

    def __del__(self):
        pass

    def _create_controls(self, program_name):
        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        b_sizer8 = wx.BoxSizer(wx.VERTICAL)

        self.m_bitmap1 = wx.StaticBitmap(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, 0)
        b_sizer8.Add(self.m_bitmap1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_staticText2 = wx.StaticText(self, wx.ID_ANY, program_name, wx.DefaultPosition, wx.DefaultSize, 0)
        b_sizer8.Add(self.m_staticText2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl3 = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(300, 150),
                                       wx.TE_MULTILINE)
        b_sizer8.Add(self.m_textCtrl3, 1, wx.EXPAND | wx.ALL, 5)

        self.m_staticText3 = wx.StaticText(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize)
        b_sizer8.Add(self.m_staticText3, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        m_sdbSizer3 = wx.StdDialogButtonSizer()
        self.m_btn_cancel = wx.Button(self, wx.ID_CANCEL)
        m_sdbSizer3.AddButton(self.m_btn_cancel)
        m_sdbSizer3.Realize()

        b_sizer8.Add(m_sdbSizer3, 0, wx.EXPAND | wx.ALL, 5)

        # append bitmap
        my_bitmap = rustre_icon.GetBitmap()
        my_image = my_bitmap.ConvertToImage()
        my_image.Rescale(64, 64)
        my_small_bmp = wx.Bitmap(my_image)
        self.m_bitmap1.SetBitmap(my_small_bmp)

        # change font and compute minimum size for text
        my_font = wx.SWISS_FONT
        my_font.SetPointSize(my_font.GetPointSize() + 3)
        self.m_staticText2.SetFont(my_font)
        dc = wx.ScreenDC()
        dc.SetFont(my_font)
        my_size = dc.GetTextExtent(program_name)
        my_size += wx.Size(10, 10)
        self.m_staticText2.SetMinSize(my_size)

        # set copyright for the current year
        my_year = wx.DateTime.Now().GetCurrentYear()
        self.m_staticText3.SetLabel("(c) Lucien SCHREIBER, " + str(my_year))
        my_font.SetPointSize(my_font.GetPointSize() - 5)
        self.m_staticText3.SetFont(my_font)

        self.SetSizer(b_sizer8)
        self.Layout()
        b_sizer8.Fit(self)

        self.Centre(wx.BOTH)
