import os.path
import pathlib

import wx
from rustre.xlsxmerge import XlsxMerge
from rustre.xlsxcompare import XlsxCompare
from rustre.xlsxduplicate import XlsxDuplicate
from rustre.version import COMMIT_ID
from rustre.version import COMMIT_NUMBER
from rustre.version import VERSION_MAJOR_MINOR
from rustre.bitmap import rustre_icon
from rustre.frameabout import FrameAbout


class FrameMain(wx.Frame):  # pragma: no cover

    def __init__(self):
        wx.Frame.__init__(self, None, id=wx.ID_ANY, title="RUSTRE", pos=wx.DefaultPosition,
                          size=wx.DefaultSize, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self._create_controls()
        self._create_menubar()
        self._create_statusbar()

        # add icon
        icon = wx.Icon()
        icon.CopyFromBitmap(rustre_icon.GetBitmap())
        self.SetIcon(icon)

        # bind events
        self.Bind(wx.EVT_BUTTON, self.on_button_paste, id=self.m_btn_paste.GetId())
        self.Bind(wx.EVT_BUTTON, self.on_button_clear, id=self.m_btn_clear.GetId())
        self.Bind(wx.EVT_BUTTON, self.on_button_merge, id=self.m_btn_do_merge.GetId())
        self.Bind(wx.EVT_BUTTON, self.on_button_compare, id=self.m_btn_compare.GetId())
        self.Bind(wx.EVT_BUTTON, self.on_button_duplicates, id=self.m_btn_duplicates.GetId())
        self.Bind(wx.EVT_MENU, self.on_menu_exit, id=self.m_menu_item_exit.GetId())
        self.Bind(wx.EVT_MENU, self.on_menu_about, id=self.m_menu_item_about.GetId())
        self.Bind(wx.EVT_MENU, self.on_menu_documentation, id=self.m_menu_item_doc.GetId())

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
        # add xlsx extension if not present
        result_filename = self.m_ctrl_result_union.GetPath()
        if pathlib.Path(result_filename).suffix != ".xlsx":
            result_filename += ".xlsx"

        if not xmerge.merge(result_filename):
            wx.LogError("Merging failed!")
            return
        wx.LogMessage("Merging done!")

    def on_button_duplicates(self, event):
        # get cols
        cols_str = self.m_ctrl_columns_list_dup.GetValue()
        if cols_str == "":
            wx.LogError("Error, no columns defined!")
            return
        col_arr = cols_str.split(",")
        xdup = XlsxDuplicate(self.m_ctrl_src_file_dup.GetPath(), self.m_ctrl_log_file_dup.GetPath(),
                             col_arr, self.m_ctrl_sheet_index_dup.GetValue(),
                             self.m_ctrl_header_index_dup.GetValue())
        if not xdup.check_duplicate():
            wx.LogError("Error checking duplicates... check your data!")
            return
        wx.LogMessage("Checking duplicate done!")

    def on_button_compare(self, event):
        my_ref_path = self.m_ctrl_reference.GetPath()
        my_target_path = self.m_ctrl_target.GetPath()
        my_model_path = self.m_ctrl_model.GetPath()
        my_log_path = self.m_ctrl_result.GetPath()

        berror = False
        if not os.path.exists(my_ref_path):
            wx.LogError("Reference path not specified!")
            berror = True
        if not os.path.exists(my_target_path):
            wx.LogError("Target path not specified!")
            berror = True
        if not os.path.exists(my_model_path):
            wx.LogError("Model path not specified!")
            berror = True
        if my_log_path == "":
            wx.LogError("Log path not specified!")
            berror = True

        if berror:
            return

        # add xlsx extension if not present
        if pathlib.Path(my_log_path).suffix != ".xlsx":
            my_log_path += ".xlsx"

        cursor = wx.BusyCursor()
        xcomp = XlsxCompare(my_model_path, my_ref_path, my_target_path)
        if xcomp.do_compare(my_log_path) is False:
            wx.LogError("Comparing files failed! please check source, target and and ini files")
            return
        wx.LogMessage("Compare done!")

    def on_menu_exit(self, event):
        self.Destroy()

    def on_menu_about(self, event):
        frm = FrameAbout(self, self.GetTitle())
        frm.ShowModal()

    def on_menu_documentation(self, event):
        wx.LaunchDefaultBrowser("https://rustre.readthedocs.io")

    def _create_statusbar(self):
        self.CreateStatusBar()
        self.SetStatusBarPane(-1)  # don't display menu hints
        self.SetStatusText("version: {}.{} ({})".format(VERSION_MAJOR_MINOR, COMMIT_NUMBER, COMMIT_ID))

    def _create_menubar(self):
        self.m_menubar1 = wx.MenuBar(0)
        self.m_menu_file = wx.Menu()
        self.m_menu_item_exit = wx.MenuItem(self.m_menu_file, wx.ID_EXIT, u"Exit", wx.EmptyString, wx.ITEM_NORMAL)
        self.m_menu_file.Append(self.m_menu_item_exit)

        self.m_menubar1.Append(self.m_menu_file, u"File")

        self.m_menu_help = wx.Menu()
        self.m_menu_item_about = wx.MenuItem(self.m_menu_help, wx.ID_ABOUT, u"About...", wx.EmptyString, wx.ITEM_NORMAL)
        self.m_menu_help.Append(self.m_menu_item_about)

        self.m_menu_item_doc = wx.MenuItem(self.m_menu_help, wx.ID_ANY, u"Documentation...", wx.EmptyString,
                                           wx.ITEM_NORMAL)
        self.m_menu_help.Append(self.m_menu_item_doc)

        self.m_menubar1.Append(self.m_menu_help, u"Help")
        self.SetMenuBar(self.m_menubar1)

    def _create_controls(self):
        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel8 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer15 = wx.BoxSizer(wx.VERTICAL)

        self.m_ctrl_notebook = wx.Notebook(self.m_panel8, wx.ID_ANY, wx.DefaultPosition, wx.Size(800, 600), 0)
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

        self.m_ctrl_result_union = wx.FilePickerCtrl(sbSizer3.GetStaticBox(), wx.ID_ANY, wx.EmptyString,
                                                     u"Select a file",
                                                     u"*.xlsx", wx.DefaultPosition, wx.DefaultSize,
                                                     wx.FLP_SAVE | wx.FLP_USE_TEXTCTRL)
        sbSizer3.Add(self.m_ctrl_result_union, 0, wx.ALL | wx.EXPAND, 5)

        bSizer3.Add(sbSizer3, 0, wx.EXPAND | wx.ALL, 5)

        self.m_btn_do_merge = wx.Button(self.m_page_merge, wx.ID_ANY, u"Merge", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer3.Add(self.m_btn_do_merge, 0, wx.ALL, 5)

        self.m_page_merge.SetSizer(bSizer3)
        self.m_page_merge.Layout()
        bSizer3.Fit(self.m_page_merge)
        self.m_ctrl_notebook.AddPage(self.m_page_merge, u"Union", False)

        self.m_page_duplicate = wx.Panel(self.m_ctrl_notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                         wx.TAB_TRAVERSAL)
        bSizer7 = wx.BoxSizer(wx.VERTICAL)

        sbSizer6 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_duplicate, wx.ID_ANY, u"Input"), wx.VERTICAL)

        fgSizer41 = wx.FlexGridSizer(0, 2, 0, 0)
        fgSizer41.AddGrowableCol(1)
        fgSizer41.SetFlexibleDirection(wx.BOTH)
        fgSizer41.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.m_staticText81 = wx.StaticText(sbSizer6.GetStaticBox(), wx.ID_ANY, u"Source file (xlsx):",
                                            wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText81.Wrap(-1)

        fgSizer41.Add(self.m_staticText81, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_src_file_dup = wx.FilePickerCtrl(sbSizer6.GetStaticBox(), wx.ID_ANY, wx.EmptyString,
                                                     u"Select a file", u"*.xlsx", wx.DefaultPosition, wx.DefaultSize,
                                                     wx.FLP_DEFAULT_STYLE | wx.FLP_OPEN)
        fgSizer41.Add(self.m_ctrl_src_file_dup, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText91 = wx.StaticText(sbSizer6.GetStaticBox(), wx.ID_ANY, u"sheet index (start at 0):",
                                            wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText91.Wrap(-1)

        fgSizer41.Add(self.m_staticText91, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_sheet_index_dup = wx.SpinCtrl(sbSizer6.GetStaticBox(), wx.ID_ANY, wx.EmptyString,
                                                  wx.DefaultPosition, wx.DefaultSize, wx.SP_ARROW_KEYS, 0, 10, 0)
        fgSizer41.Add(self.m_ctrl_sheet_index_dup, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText10 = wx.StaticText(sbSizer6.GetStaticBox(), wx.ID_ANY, u"header index (start at 1):",
                                            wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText10.Wrap(-1)

        fgSizer41.Add(self.m_staticText10, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_header_index_dup = wx.SpinCtrl(sbSizer6.GetStaticBox(), wx.ID_ANY, wx.EmptyString,
                                                   wx.DefaultPosition, wx.DefaultSize, wx.SP_ARROW_KEYS, 1, 100, 1)
        fgSizer41.Add(self.m_ctrl_header_index_dup, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText11 = wx.StaticText(sbSizer6.GetStaticBox(), wx.ID_ANY, u"Columns:", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText11.Wrap(-1)

        fgSizer41.Add(self.m_staticText11, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_columns_list_dup = wx.TextCtrl(sbSizer6.GetStaticBox(), wx.ID_ANY, wx.EmptyString,
                                                   wx.DefaultPosition, wx.DefaultSize, 0)
        fgSizer41.Add(self.m_ctrl_columns_list_dup, 0, wx.ALL | wx.EXPAND, 5)

        fgSizer41.Add((0, 0), 1, wx.EXPAND, 5)

        self.m_staticText12 = wx.StaticText(sbSizer6.GetStaticBox(), wx.ID_ANY,
                                            u"Comma separeted list of columns index : 0,2,4", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText12.Wrap(-1)

        fgSizer41.Add(self.m_staticText12, 0, wx.ALL, 5)

        sbSizer6.Add(fgSizer41, 1, wx.EXPAND, 5)

        bSizer7.Add(sbSizer6, 1, wx.EXPAND | wx.ALL, 5)

        sbSizer7 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_duplicate, wx.ID_ANY, u"Results"), wx.VERTICAL)

        bSizer8 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText13 = wx.StaticText(sbSizer7.GetStaticBox(), wx.ID_ANY, u"Log file (xlsx)", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText13.Wrap(-1)

        bSizer8.Add(self.m_staticText13, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_log_file_dup = wx.FilePickerCtrl(sbSizer7.GetStaticBox(), wx.ID_ANY, wx.EmptyString,
                                                     u"Select a file", u"*.xlsx", wx.DefaultPosition, wx.DefaultSize,
                                                     wx.FLP_SAVE | wx.FLP_USE_TEXTCTRL)
        bSizer8.Add(self.m_ctrl_log_file_dup, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        sbSizer7.Add(bSizer8, 0, wx.EXPAND, 5)

        bSizer7.Add(sbSizer7, 1, wx.EXPAND | wx.ALL, 5)

        self.m_btn_duplicates = wx.Button(self.m_page_duplicate, wx.ID_ANY, u"Search duplicates", wx.DefaultPosition,
                                          wx.DefaultSize, 0)
        bSizer7.Add(self.m_btn_duplicates, 0, wx.ALL, 5)

        self.m_page_duplicate.SetSizer(bSizer7)
        self.m_page_duplicate.Layout()
        bSizer7.Fit(self.m_page_duplicate)
        self.m_ctrl_notebook.AddPage(self.m_page_duplicate, u"Duplicates", True)

        self.m_page_compare = wx.Panel(self.m_ctrl_notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                       wx.TAB_TRAVERSAL)
        bSizer17 = wx.BoxSizer(wx.VERTICAL)

        sbSizer10 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_compare, wx.ID_ANY, u"Input"), wx.VERTICAL)

        fgSizer4 = wx.FlexGridSizer(3, 2, 0, 0)
        fgSizer4.AddGrowableCol(1)
        fgSizer4.SetFlexibleDirection(wx.BOTH)
        fgSizer4.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.m_staticText7 = wx.StaticText(sbSizer10.GetStaticBox(), wx.ID_ANY, u"Reference:", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText7.Wrap(-1)

        fgSizer4.Add(self.m_staticText7, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_reference = wx.FilePickerCtrl(sbSizer10.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file",
                                                  u"*.xlsx", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        fgSizer4.Add(self.m_ctrl_reference, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText8 = wx.StaticText(sbSizer10.GetStaticBox(), wx.ID_ANY, u"Target:", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText8.Wrap(-1)

        fgSizer4.Add(self.m_staticText8, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_target = wx.FilePickerCtrl(sbSizer10.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file",
                                               u"*.xlsx", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        fgSizer4.Add(self.m_ctrl_target, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText9 = wx.StaticText(sbSizer10.GetStaticBox(), wx.ID_ANY, u"Model:", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText9.Wrap(-1)

        fgSizer4.Add(self.m_staticText9, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_model = wx.FilePickerCtrl(sbSizer10.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file",
                                              u"*.ini", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        fgSizer4.Add(self.m_ctrl_model, 0, wx.ALL | wx.EXPAND, 5)

        sbSizer10.Add(fgSizer4, 1, wx.EXPAND, 5)

        sbSizer10.Add((0, 0), 1, wx.EXPAND, 5)

        bSizer17.Add(sbSizer10, 0, wx.ALL | wx.EXPAND, 5)

        sbSizer11 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_compare, wx.ID_ANY, u"Results"), wx.VERTICAL)

        bSizer18 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText10 = wx.StaticText(sbSizer11.GetStaticBox(), wx.ID_ANY, u"Log file:", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText10.Wrap(-1)

        bSizer18.Add(self.m_staticText10, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_result = wx.FilePickerCtrl(sbSizer11.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file",
                                               u"*.xlsx", wx.DefaultPosition, wx.DefaultSize,
                                               wx.FLP_SAVE | wx.FLP_USE_TEXTCTRL)
        bSizer18.Add(self.m_ctrl_result, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        sbSizer11.Add(bSizer18, 0, wx.EXPAND, 5)

        bSizer17.Add(sbSizer11, 1, wx.EXPAND | wx.ALL, 5)

        self.m_btn_compare = wx.Button(self.m_page_compare, wx.ID_ANY, u"Compare", wx.DefaultPosition, wx.DefaultSize,
                                       0)
        bSizer17.Add(self.m_btn_compare, 0, wx.ALL, 5)

        self.m_page_compare.SetSizer(bSizer17)
        self.m_page_compare.Layout()
        bSizer17.Fit(self.m_page_compare)
        self.m_ctrl_notebook.AddPage(self.m_page_compare, u"People", False)

        self.m_page_join = wx.Panel(self.m_ctrl_notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                    wx.TAB_TRAVERSAL)
        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        sbSizer8 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_join, wx.ID_ANY, u"First file"), wx.VERTICAL)

        fgSizer42 = wx.FlexGridSizer(0, 2, 0, 0)
        fgSizer42.AddGrowableCol(1)
        fgSizer42.SetFlexibleDirection(wx.BOTH)
        fgSizer42.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.m_staticText131 = wx.StaticText(sbSizer8.GetStaticBox(), wx.ID_ANY, u"File", wx.DefaultPosition,
                                             wx.DefaultSize, 0)
        self.m_staticText131.Wrap(-1)

        fgSizer42.Add(self.m_staticText131, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_file1 = wx.FilePickerCtrl(sbSizer8.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file",
                                                   u"*.*", wx.DefaultPosition, wx.DefaultSize,
                                                   wx.FLP_DEFAULT_STYLE | wx.FLP_USE_TEXTCTRL)
        fgSizer42.Add(self.m_ctrl_join_file1, 0, wx.ALL | wx.EXPAND | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_staticText14 = wx.StaticText(sbSizer8.GetStaticBox(), wx.ID_ANY, u"Sheet", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText14.Wrap(-1)

        fgSizer42.Add(self.m_staticText14, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_sheet1 = wx.SpinCtrl(sbSizer8.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                              wx.DefaultSize, wx.SP_ARROW_KEYS, 0, 20, 0)
        fgSizer42.Add(self.m_ctrl_join_sheet1, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText15 = wx.StaticText(sbSizer8.GetStaticBox(), wx.ID_ANY, u"Skip rows", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText15.Wrap(-1)

        fgSizer42.Add(self.m_staticText15, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_row1 = wx.SpinCtrl(sbSizer8.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                            wx.DefaultSize, wx.SP_ARROW_KEYS, 0, 50, 0)
        fgSizer42.Add(self.m_ctrl_join_row1, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText16 = wx.StaticText(sbSizer8.GetStaticBox(), wx.ID_ANY, u"Id column index", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText16.Wrap(-1)

        fgSizer42.Add(self.m_staticText16, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_id1 = wx.SpinCtrl(sbSizer8.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                           wx.DefaultSize, wx.SP_ARROW_KEYS, 0, 1000, 0)
        fgSizer42.Add(self.m_ctrl_join_id1, 0, wx.ALL | wx.EXPAND | wx.ALIGN_CENTER_VERTICAL, 5)

        sbSizer8.Add(fgSizer42, 1, wx.EXPAND, 5)

        bSizer9.Add(sbSizer8, 1, wx.EXPAND | wx.ALL, 5)

        sbSizer9 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_join, wx.ID_ANY, u"Second file"), wx.VERTICAL)

        fgSizer5 = wx.FlexGridSizer(0, 2, 0, 0)
        fgSizer5.AddGrowableCol(1)
        fgSizer5.SetFlexibleDirection(wx.BOTH)
        fgSizer5.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.m_staticText17 = wx.StaticText(sbSizer9.GetStaticBox(), wx.ID_ANY, u"File", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText17.Wrap(-1)

        fgSizer5.Add(self.m_staticText17, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_file2 = wx.FilePickerCtrl(sbSizer9.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file",
                                                   u"*.*", wx.DefaultPosition, wx.DefaultSize,
                                                   wx.FLP_DEFAULT_STYLE | wx.FLP_USE_TEXTCTRL)
        fgSizer5.Add(self.m_ctrl_join_file2, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText18 = wx.StaticText(sbSizer9.GetStaticBox(), wx.ID_ANY, u"Sheet", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText18.Wrap(-1)

        fgSizer5.Add(self.m_staticText18, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_sheet2 = wx.SpinCtrl(sbSizer9.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                              wx.DefaultSize, wx.SP_ARROW_KEYS, 0, 20, 0)
        fgSizer5.Add(self.m_ctrl_join_sheet2, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.m_staticText19 = wx.StaticText(sbSizer9.GetStaticBox(), wx.ID_ANY, u"Skip rows", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText19.Wrap(-1)

        fgSizer5.Add(self.m_staticText19, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_row2 = wx.SpinCtrl(sbSizer9.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                            wx.DefaultSize, wx.SP_ARROW_KEYS, 0, 50, 0)
        fgSizer5.Add(self.m_ctrl_join_row2, 0, wx.ALL | wx.EXPAND, 5)

        self.m_staticText20 = wx.StaticText(sbSizer9.GetStaticBox(), wx.ID_ANY, u"Id column index", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText20.Wrap(-1)

        fgSizer5.Add(self.m_staticText20, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_id2 = wx.SpinCtrl(sbSizer9.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                           wx.DefaultSize, wx.SP_ARROW_KEYS, 0, 1000, 0)
        fgSizer5.Add(self.m_ctrl_join_id2, 0, wx.ALL | wx.EXPAND, 5)

        sbSizer9.Add(fgSizer5, 1, wx.EXPAND, 5)

        bSizer9.Add(sbSizer9, 1, wx.EXPAND | wx.ALL, 5)

        sbSizer101 = wx.StaticBoxSizer(wx.StaticBox(self.m_page_join, wx.ID_ANY, u"Results"), wx.VERTICAL)

        fgSizer6 = wx.FlexGridSizer(0, 2, 0, 0)
        fgSizer6.AddGrowableCol(1)
        fgSizer6.SetFlexibleDirection(wx.BOTH)
        fgSizer6.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.m_staticText21 = wx.StaticText(sbSizer101.GetStaticBox(), wx.ID_ANY, u"Result", wx.DefaultPosition,
                                            wx.DefaultSize, 0)
        self.m_staticText21.Wrap(-1)

        fgSizer6.Add(self.m_staticText21, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.m_ctrl_join_result = wx.FilePickerCtrl(sbSizer101.GetStaticBox(), wx.ID_ANY, wx.EmptyString,
                                                    u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize,
                                                    wx.FLP_SAVE | wx.FLP_USE_TEXTCTRL)
        fgSizer6.Add(self.m_ctrl_join_result, 0, wx.ALL | wx.EXPAND, 5)

        sbSizer101.Add(fgSizer6, 1, wx.EXPAND, 5)

        bSizer9.Add(sbSizer101, 0, wx.EXPAND | wx.ALL, 5)

        self.m_ctrl_join_btn = wx.Button(self.m_page_join, wx.ID_ANY, u"Join", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer9.Add(self.m_ctrl_join_btn, 0, wx.ALL, 5)

        self.m_page_join.SetSizer(bSizer9)
        self.m_page_join.Layout()
        bSizer9.Fit(self.m_page_join)
        self.m_ctrl_notebook.AddPage(self.m_page_join, u"Join", True)

        bSizer15.Add(self.m_ctrl_notebook, 1, wx.EXPAND | wx.ALL, 5)

        self.m_panel8.SetSizer(bSizer15)
        self.m_panel8.Layout()
        bSizer15.Fit(self.m_panel8)
        bSizer1.Add(self.m_panel8, 1, wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()
        bSizer1.Fit(self)

        self.m_ctrl_notebook.SetSelection(0)

        self.Centre(wx.BOTH)
