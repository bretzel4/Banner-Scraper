#!/usr/bin/env python

"""
ZetCode wxPython tutorial

In this example we create a rename layout
with wx.GridBagSizer.

author: Jan Bodnar
website: www.zetcode.com
last modified: July 2020
"""

import wx


class Example(wx.Frame):

    def __init__(self, parent, title):
        super(Example, self).__init__(parent, title=title)

        self.InitUI()
        self.Centre()

    def InitUI(self):

        panel = wx.Panel(self)
        sizer = wx.GridBagSizer(4, 4)

        text = wx.StaticText(panel, label="Paste text from Banner")
        sizer.Add(text, pos=(0, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        # staticIcon = wx.Button(panel, size=(24, 24))
        # staticIcon.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_WARNING))
        staticIcon = wx.StaticBitmap(panel, id=wx.ID_INFO, bitmap=wx.ArtProvider.GetBitmap(wx.ART_QUESTION, size=(16, 16)),
             size=(24, 24), style=0, name='')
        staticIcon.SetToolTip('Ctrl+A on Banner then Ctrl+V here. If there are multiple pages on Banner, paste them sequentially here.')
        sizer.Add(staticIcon, pos=(0, 1), border=0, flag=wx.CENTER)

        tc = wx.TextCtrl(panel, style=wx.TE_MULTILINE)
        sizer.Add(tc, pos=(1, 0), span=(5, 5),
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)
        self.tc = tc

        buttonSave = wx.Button(panel, wx.ID_SAVE, label="Scrape", size=(90, 28))
        buttonClose = wx.Button(panel, wx.ID_CLEAR, label="Reset", size=(90, 28))
        sizer.Add(buttonSave, pos=(6, 3))
        sizer.Add(buttonClose, pos=(6, 4), flag=wx.RIGHT|wx.BOTTOM, border=10)

        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(2)
        panel.SetSizer(sizer)

        self.Bind(wx.EVT_BUTTON, self.OnSaveAs, id=wx.ID_SAVE)
        self.Bind(wx.EVT_BUTTON, self.OnClear, id=wx.ID_CLEAR)

    def Scrape(self):
        progressMax = 100
        with wx.ProgressDialog("A progress box", "Time remaining", progressMax,
                    style=wx.PD_CAN_ABORT | wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME) as progressDialog:
            keepGoing = True
            count = 0
            while keepGoing and count < progressMax:
                count = count + 1
                wx.Sleep(1)
                keepGoing = progressDialog.Update(count)[0]
                print(keepGoing)

    def OnSaveAs(self, event):
        self.Scrape()

        with wx.DirDialog(None, "Choose save directory", "",
                    wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST) as dirDialog:

            if dirDialog.ShowModal() == wx.ID_CANCEL:
                return     # the user changed their mind

            # save the current contents in the file
            pathname = dirDialog.GetPath() + '\\test.txt'
            print(pathname)
            try:
                with open(pathname, 'w') as file:
                    file.write(self.tc.GetValue())
            except IOError:
                wx.LogError("Cannot save current data in file '%s'." % pathname)

    def OnClear(self, event):
        self.tc.Clear()


def main():

    app = wx.App()
    ex = Example(None, title='Student Scraper')
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()