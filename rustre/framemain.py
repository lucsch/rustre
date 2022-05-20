import wx
from rustre.xlsxmerge import XlsxMerge


class FrameMain(wx.Dialog):  # pragma: no cover

    def __init__(self):
        wx.Dialog.__init__(self, None, id=wx.ID_ANY, title="RUSTRE", pos=wx.DefaultPosition,
                           size=wx.DefaultSize, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self._create_controls()

        # bind events
        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.Bind(wx.EVT_BUTTON, self.on_button_paste, id=self.m_btn_paste.GetId())
        self.Bind(wx.EVT_BUTTON, self.on_button_clear, id=self.m_btn_clear.GetId())
        self.Bind(wx.EVT_BUTTON, self.on_button_merge, id=self.m_btn_do_merge.GetId())

    def on_close(self, event):
        self.Destroy()

    def on_button_paste(self, event):
        if not wx.TheClipboard.IsOpened():
            do = wx.FileDataObject()
            wx.TheClipboard.Open()
            success = wx.TheClipboard.GetData(do)
            wx.TheClipboard.Close()
            if success:
                self.m_ctrl_list.Clear()
                for f in do.GetFilenames():
                    self.m_ctrl_list.Append(f)
            else:
                wx.MessageBox("""There is no data in the clipboard in the required format""")

    def on_button_clear(self, event):
        self.m_ctrl_list.Clear()

    def on_button_merge(self, event):
        try:
            xmerge = XlsxMerge(self.m_ctrl_list.GetItems(), self.m_ctrl_sheet.GetValue(), self.m_ctrl_header.GetValue())
        except ValueError:
            wx.LogError("Error merging files, check your data!")
            return
        if not xmerge.merge(self.m_ctrl_result.GetPath()):
            wx.LogError("Merging failed!")
            return
        wx.LogMessage("Merging done!")

    def _create_controls(self):
        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        self.m_ctrl_notebook = wx.Notebook(self, wx.ID_ANY, wx.DefaultPosition, wx.Size(800, 600), 0)
        self.m_page_merge = wx.Panel(self.m_ctrl_notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                     wx.TAB_TRAVERSAL)
        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        sbSizer1 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_merge, wx.ID_ANY, u"Files"), wx.VERTICAL)

        m_ctrl_listChoices = []
        self.m_ctrl_list = wx.ListBox(sbSizer1.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                      m_ctrl_listChoices, wx.LB_MULTIPLE)
        sbSizer1.Add(self.m_ctrl_list, 1, wx.ALL | wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_btn_paste = wx.Button(sbSizer1.GetStaticBox(), wx.ID_ANY, u"Paste", wx.DefaultPosition, wx.DefaultSize,
                                     0)
        bSizer4.Add(self.m_btn_paste, 0, wx.ALL, 5)

        self.m_btn_clear = wx.Button(sbSizer1.GetStaticBox(), wx.ID_ANY, u"Clear", wx.DefaultPosition, wx.DefaultSize,
                                     0)
        bSizer4.Add(self.m_btn_clear, 0, wx.ALL, 5)

        sbSizer1.Add(bSizer4, 0, wx.EXPAND, 5)

        bSizer3.Add(sbSizer1, 1, wx.EXPAND | wx.ALL, 5)

        sbSizer2 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_merge, wx.ID_ANY, u"Options"), wx.VERTICAL)

        fgSizer1 = wx.FlexGridSizer(0, 2, 0, 0)
        fgSizer1.AddGrowableCol(1)
        fgSizer1.SetFlexibleDirection(wx.BOTH)
        fgSizer1.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.m_staticText1 = wx.StaticText(sbSizer2.GetStaticBox(), wx.ID_ANY, u"sheet index (zero based): ",
                                           wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText1.Wrap(-1)

        fgSizer1.Add(self.m_staticText1, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_sheet = wx.SpinCtrl(sbSizer2.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                        wx.DefaultSize, wx.SP_ARROW_KEYS, 0, 20, 0)
        fgSizer1.Add(self.m_ctrl_sheet, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.m_staticText2 = wx.StaticText(sbSizer2.GetStaticBox(), wx.ID_ANY, u"header index (start at 1):",
                                           wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText2.Wrap(-1)

        fgSizer1.Add(self.m_staticText2, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_header = wx.SpinCtrl(sbSizer2.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                         wx.DefaultSize, wx.SP_ARROW_KEYS, 1, 100, 1)
        fgSizer1.Add(self.m_ctrl_header, 0, wx.ALL | wx.EXPAND | wx.ALIGN_CENTER_VERTICAL, 5)

        sbSizer2.Add(fgSizer1, 1, wx.EXPAND, 5)

        bSizer3.Add(sbSizer2, 0, wx.EXPAND | wx.ALL, 5)

        sbSizer3 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_merge, wx.ID_ANY, u"Result"), wx.VERTICAL)

        self.m_ctrl_result = wx.FilePickerCtrl(sbSizer3.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file",
                                               u"*.xlsx", wx.DefaultPosition, wx.DefaultSize,
                                               wx.FLP_SAVE | wx.FLP_USE_TEXTCTRL)
        sbSizer3.Add(self.m_ctrl_result, 0, wx.ALL | wx.EXPAND, 5)

        bSizer3.Add(sbSizer3, 0, wx.EXPAND | wx.ALL, 5)

        self.m_btn_do_merge = wx.Button(self.m_page_merge, wx.ID_ANY, u"Merge", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer3.Add(self.m_btn_do_merge, 0, wx.ALL, 5)

        self.m_page_merge.SetSizer(bSizer3)
        self.m_page_merge.Layout()
        bSizer3.Fit(self.m_page_merge)
        self.m_ctrl_notebook.AddPage(self.m_page_merge, u"Merge", False)
        self.m_page_compare = wx.Panel(self.m_ctrl_notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                       wx.TAB_TRAVERSAL)
        self.m_ctrl_notebook.AddPage(self.m_page_compare, u"Compare", False)

        bSizer1.Add(self.m_ctrl_notebook, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(bSizer1)
        self.Layout()
        bSizer1.Fit(self)

        self.Centre(wx.BOTH)
