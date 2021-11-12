#!/usr/bin/env python

import wx
import wx.adv
import scrapy # web scraping
from scrapy import signals
from scrapy.crawler import CrawlerProcess
import json   # dumping and loading variabls
import re     # regex
from pandas import DataFrame # excel
from pandas import ExcelWriter
from bs4 import BeautifulSoup

df1 = False
df2 = False
prof = False

class BannerParser:
    raw = ''
    students_list = []
    students_stripped = []
    subject = ''
    number = ''
    term = ''
    crn = ''
    section = ''

    def __init__(self, text):
        self.raw = text

        self.students_list = re.findall('\d{9}\n(.+?(?=,).+)', self.raw)
        self.students_stripped = list(map(self.FormatName, self.students_list))

        info_line = re.search('Term.+\n', self.raw).group(0).split(' ')
        self.subject = info_line[5]
        self.number = info_line[6]
        self.term = info_line[1]
        self.crn = info_line[4]
        self.section = info_line[7]
        print([self.subject, self.number, self.term, self.crn, self.section])

    def FormatName(self, name):
        name = name.split(',')
        name.reverse()
        name = ' '.join(name)
        name = name.strip()
        name = re.sub('\(.+\)', '', name)
        name = re.sub('\s+', ' ', name)
        name = re.sub('\.', '', name)
        return name

class StudentSpider(scrapy.Spider):
    # errors = ''
    name = 'uvm'
    start_urls = []
    names = []
    emails = []
    depyears = []
    error_names = []
    error_errors = []
    name_from_url = {}

    def parse(self, response):
        responses = response.json()['data']
        people = list(map(lambda x: {'name': x['cn']['0'], 'email': x['mail']['0'], 'depyears': x['ou']['0']}, responses))

        pset = set()
        for p in people:
            pset.add(json.dumps(p, sort_keys='true'))
        people = list(map(lambda x: json.loads(x), pset))

        if len(people)==0:
            self.error_names.append(self.name_from_url[response.url])
            self.error_errors.append('email not found')
        elif len(people)>1:
            self.error_names.append(self.name_from_url[response.url])
            self.error_errors.append('too many matches for name')
        else:
            self.names.append(self.name_from_url[response.url])
            self.emails.append(people[0]['email'])
            self.depyears.append(people[0]['depyears'])

    def closed(self, reason):
        # with open('errors.txt', 'w') as f:
        #     f.write(self.errors)
        global df1
        df1 = DataFrame({'Name': self.names, 'Email': self.emails, 'Department/Year': self.depyears})
        df1.sort_values(by=['Name'], inplace=True)
        global df2
        df2 = DataFrame({'Name': self.error_names, 'Error': self.error_errors})
        df2.sort_values(by=['Name'], inplace=True)

class ProfSpider(scrapy.Spider):
    name = 'prof'
    crn = ''
    start_urls = []

    def parse(self, response):
        print(self.start_urls)
        print(self.crn)
        print(response)
        soup = BeautifulSoup(str(response.body))
        table = soup.find('table', id='sections')
        for row in table.find_all('tr'):
            if self.crn in str(row):
                global prof
                prof = row.select_one('td[data-label="Instructor"]').get_text().rstrip('\\t').split()[-1]
                print(prof)

class GuiManager(wx.Frame):
    def __init__(self, parent, title):
        super(GuiManager, self).__init__(parent, title=title)

        self.InitUI()
        self.Centre()

    def InitUI(self):
        font = wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        font_title = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)

        panel = wx.Panel(self)
        sizer = wx.GridBagSizer(5, 4)

        title = wx.StaticText(panel, label="Banner Scraper by Lilac Damon")
        title.SetFont(font_title)
        sizer.Add(title, pos=(0, 0), span=(1, 10), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        text = wx.StaticText(panel, label="Paste text from Banner")
        text.SetFont(font)
        sizer.Add(text, pos=(1, 0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)

        # staticIcon = wx.Button(panel, size=(24, 24))
        # staticIcon.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_WARNING))
        staticIcon = wx.StaticBitmap(panel, id=wx.ID_INFO, bitmap=wx.ArtProvider.GetBitmap(wx.ART_QUESTION, size=(16, 16)),
             size=(24, 24), style=0, name='')
        staticIcon.SetToolTip('Ctrl+A on Banner then Ctrl+V here. If there are multiple pages on Banner, paste them sequentially here.')
        sizer.Add(staticIcon, pos=(1, 1), border=0, flag=wx.CENTER)

        tc = wx.TextCtrl(panel, style=wx.TE_MULTILINE)
        sizer.Add(tc, pos=(2, 0), span=(5, 5),
            flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)
        self.tc = tc

        buttonSave = wx.Button(panel, wx.ID_SAVE, label="Scrape", size=(90, 28))
        buttonClose = wx.Button(panel, wx.ID_CLEAR, label="Reset", size=(90, 28))
        buttonInfo = wx.Button(panel, wx.ID_INFO, label="Info and Privacy", size=(110, 28))
        sizer.Add(buttonInfo, pos=(7, 0), flag=wx.LEFT|wx.BOTTOM, border=5)
        sizer.Add(buttonSave, pos=(7, 3))
        sizer.Add(buttonClose, pos=(7, 4), flag=wx.RIGHT|wx.BOTTOM, border=5)

        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(2)
        panel.SetSizer(sizer)

        self.Bind(wx.EVT_BUTTON, self.OnInfo, id=wx.ID_INFO)
        self.Bind(wx.EVT_BUTTON, self.OnSaveAs, id=wx.ID_SAVE)
        self.Bind(wx.EVT_BUTTON, self.OnClear, id=wx.ID_CLEAR)

    def Scrape(self, banner_parser):
        base_url = 'https://www.uvm.edu/directory/api/query_results.php?name='
        urls = list(map(lambda x: base_url + x.replace(" ", "%20"), banner_parser.students_stripped))
        name_from_url = dict(zip(urls, banner_parser.students_list))

        process = CrawlerProcess(settings={})
        process.crawl(StudentSpider, start_urls=urls, name_from_url=name_from_url)

        prof_url = 'https://www.uvm.edu/coursedirectory/search.php?subject=' + banner_parser.subject + \
                        '&number=' + banner_parser.number + \
                        '&term=' + banner_parser.term + '&section'
        print(prof_url)
        process.crawl(ProfSpider, start_urls=[prof_url], crn=banner_parser.crn)
        process.start()

    def OnSaveAs(self, event):
        banner_parser = BannerParser(self.tc.GetValue())
        self.Scrape(banner_parser)

        print(df1)
        print(df2)
        print(prof)

        with wx.DirDialog(None, "Choose save directory", "",
                    wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST) as dirDialog:
            if dirDialog.ShowModal() == wx.ID_CANCEL:
                return # the user changed their mind

            # save the current contents in the file
            pathname = dirDialog.GetPath() + '\\' + banner_parser.subject + ' ' + banner_parser.number + ' ' + \
                                             banner_parser.section + ' (' + prof + ').xlsx'
            print(pathname)
            try:
                with ExcelWriter(pathname) as writer:
                    df1.to_excel(writer, sheet_name="Results", index=False)
                    df2.to_excel(writer, sheet_name="Errors", index=False)
            except IOError:
                wx.LogError("Cannot save current data in file '%s'." % pathname)

    def OnClear(self, event):
        self.tc.Clear()

    def OnInfo(self, event):
        description = """Created by Lilac Damon http://www.meltedlilacs.com

This program uses pasted info from Banner in order to compile a list of student emails and other details.
The source code is freely available at the below url. This repository is unlisted but publicly accessible.

If you have privacy concerns, please note:
1) This program never directly connects to Banner. It therefore has no way to access or manipulate privileged information.
2) This program only connects to public websites (UVM directories).
3) The source code does not reveal any information about Banner other than basic information
    about where items are located on the page (ex: student names appear near 9-digit numbers).
4) All of these claims may be verified by looking at the source code which is available at the below url."""

        licence = """Copyright 2021 Lilac Damon

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."""


        info = wx.adv.AboutDialogInfo()

        info.SetName('Banner Scraper')
        info.SetVersion('1.0')
        info.SetDescription(description)
        info.SetWebSite('[url]')
        info.SetLicence(licence)

        wx.adv.AboutBox(info)

def main():

    app = wx.App()
    ex = GuiManager(None, title='Banner Scraper')
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()