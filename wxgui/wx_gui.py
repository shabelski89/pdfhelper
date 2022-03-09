import wx
from wxgui.wx_widget import MainFrame


class GenApp(wx.App):
    def __init__(self, redirect=True, filename=None):
        wx.App.__init__(self, redirect, filename)

    def OnInit(self):
        frame = MainFrame()
        frame.Center()
        frame.Show()
        return True


def main():
    app = GenApp(redirect=False)
    app.MainLoop()


if __name__ == "__main__":
    main()
