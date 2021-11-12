#!/usr/bin/env python

import scrapy # web scraping
from scrapy import signals
from scrapy.crawler import CrawlerProcess
import json   # dumping and loading variabls
import re     # regex
from pandas import DataFrame # excel
from pandas import ExcelWriter
from bs4 import BeautifulSoup
import logging
import sys
import os
import socket

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

        try:
            self.students_list = re.findall('\d{9}\n(.+?(?=,).+)', self.raw)
            self.students_stripped = list(map(self.FormatName, self.students_list))
            info_line = re.search('Term.+\n', self.raw).group(0).split(' ')
        except AttributeError:
            input("ERROR: No student names found.\nPress Enter to continue.")
            sys.exit()

        self.subject = info_line[5]
        self.number = info_line[6]
        self.term = info_line[1]
        self.crn = info_line[4]
        self.section = info_line[7]
        # print([self.subject, self.number, self.term, self.crn, self.section])

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
    custom_settings = {'LOG_FILE': 'log.txt'}

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
    custom_settings = {'LOG_FILE': 'log.txt'}

    def parse(self, response):
        # print(self.start_urls)
        # print(self.crn)
        # print(response)
        soup = BeautifulSoup(str(response.body))
        table = soup.find('table', id='sections')
        for row in table.find_all('tr'):
            if self.crn in str(row):
                global prof
                prof = row.select_one('td[data-label="Instructor"]').get_text().rstrip('\\t').split()[-1]
                # print(prof)

def check_connection():
    try:
        socket.create_connection(("www.google.com", 443))
        return True
    except:
        pass
    return False

def Scrape(banner_parser):
    if(not check_connection()):
        input("ERROR: No internet connection.\nPress Enter to continue.")
        sys.exit()

    base_url = 'https://www.uvm.edu/directory/api/query_results.php?name='
    urls = list(map(lambda x: base_url + x.replace(" ", "%20"), banner_parser.students_stripped))
    name_from_url = dict(zip(urls, banner_parser.students_list))

    os.remove('.\\log.txt')

    process = CrawlerProcess()
    process.crawl(StudentSpider, start_urls=urls, name_from_url=name_from_url)

    prof_url = 'https://www.uvm.edu/coursedirectory/search.php?subject=' + banner_parser.subject + \
                    '&number=' + banner_parser.number + \
                    '&term=' + banner_parser.term + '&section'
    # print(prof_url)
    process.crawl(ProfSpider, start_urls=[prof_url], crn=banner_parser.crn)
    process.start()

def main():
    logging.getLogger('scrapy').setLevel(logging.WARNING)

    with open('students.txt') as f:
        banner_parser = BannerParser(f.read())

    Scrape(banner_parser)

    print(df1)
    print(df2)
    # print(prof)

    # save the current contents in the file
    pathname = '.\\' + banner_parser.subject + ' ' + banner_parser.number + ' ' + \
                       banner_parser.section + ' (' + prof + ').xlsx'
    print(pathname)
    try:
        with ExcelWriter(pathname) as writer:
            df1.to_excel(writer, sheet_name="Results", index=False)
            df2.to_excel(writer, sheet_name="Errors", index=False)
    except IOError:
        input("ERROR: Cannot save current data in file " + pathname + ".\n Press Enter to continue.")

if __name__ == '__main__':
    main()