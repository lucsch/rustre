#!/usr/bin/env python3
import wx

from rustre.framemain import FrameMain


class RustreApp(wx.App):
    """
    Main application class
    init the 'framemain' and the main loop
    """

    def OnInit(self):
        dlg = FrameMain()
        dlg.Show(True)
        self.SetTopWindow(dlg)
        return True


if __name__ == '__main__':
    app = RustreApp()
    app.MainLoop()
