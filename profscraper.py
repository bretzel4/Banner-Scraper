#!/usr/bin/env python

import scrapy
from scrapy.crawler import CrawlerProcess
from bs4 import BeautifulSoup

crn = '94371'

class ProfSpider(scrapy.Spider):
    name = 'prof'
    start_urls = [
        'https://www.uvm.edu/coursedirectory/search.php?subject=ECLD&number=056&term=202109&section',
    ]

    def parse(self, response):
        soup = BeautifulSoup(str(response.body))
        table = soup.find('table', id='sections')
        for row in table.find_all('tr'):
            print('aaaaaaaaaaaaaa' + str(row))
            if crn in str(row):
                print('AAAAAAAAAAAAAA')
                print(row.select_one('td[data-label="Instructor"]').get_text().rstrip('\\t').split()[-1])
            # if crn in row.td.text:
            #     prof = row.td.a.text
            #     print(prof)
            #     break

process = CrawlerProcess(settings={})
process.crawl(ProfSpider)
process.start()