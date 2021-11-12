import scrapy # web scraping
from scrapy import signals
from scrapy.crawler import CrawlerProcess
import json   # dumping and loading variabls
import re     # regex
from pandas import DataFrame # excel
from pandas import ExcelWriter

def format_name(name):
    name = name.split(',')
    name.reverse()
    name = ' '.join(name)
    name = name.strip()
    name = re.sub('\(.+\)', '', name)
    name = re.sub('\s+', ' ', name)
    name = re.sub('\.', '', name)
    return name

with open('students.txt') as f:
    students_raw = f.read()
    students_list = re.findall('\d{9}\n(.+?(?=,).+)', students_raw)
    students_stripped = list(map(format_name, students_list))

base_url = 'https://www.uvm.edu/directory/api/query_results.php?name='
urls = list(map(lambda x: base_url + x.replace(" ", "%20"), students_stripped))
name_from_url = dict(zip(urls, students_list))
print(name_from_url)

class StudentSpider(scrapy.Spider):
    # errors = ''
    name = 'uvm'
    start_urls = urls
    names = []
    emails = []
    depyears = []
    error_names = []
    error_errors = []

    def parse(self, response):
        # print('AAAAAAAAAAAAAAAAAAAAAAA')
        # print(name_from_url[response.url])
        # print('AAAAAAAAAAAAAAAAAAAAAAA')
        # self.logger.info('aaaaaaaaaaaaaaaa' + str(response.json()))

        responses = response.json()['data']
        people = list(map(lambda x: {'name': x['cn']['0'], 'email': x['mail']['0'], 'depyears': x['ou']['0']}, responses))

        pset = set()
        for p in people:
            pset.add(json.dumps(p, sort_keys='true'))
        people = list(map(lambda x: json.loads(x), pset))

        # self.logger.info(people)
        if len(people)==0:
            # self.errors += 'name: ' + response.json()['search'][0] + '\n\terror: email not found\n'
            self.error_names.append(name_from_url[response.url])
            self.error_errors.append('email not found')
        elif len(people)>1:
            # self.errors += 'name: ' + response.json()['search'][0] + '\n\terror: too many matches for name\n'
            self.error_names.append(name_from_url[response.url])
            self.error_errors.append('too many matches for name')
        else:
            # yield people[0]
            self.names.append(name_from_url[response.url])
            self.emails.append(people[0]['email'])
            self.depyears.append(people[0]['depyears'])

    # @classmethod
    # def from_crawler(cls, crawler, *args, **kwargs):
    #     spider = super(StudentSpider, cls).from_crawler(crawler, *args, **kwargs)
    #     crawler.signals.connect(spider.spider_closed, signal=signals.spider_closed)
    #     return spider

    def closed(self, reason):
        # with open('errors.txt', 'w') as f:
        #     f.write(self.errors)
        df1 = DataFrame({'Name': self.names, 'Email': self.emails, 'Department/Year': self.depyears})
        df1.sort_values(by=['Name'], inplace=True)
        df2 = DataFrame({'Name': self.error_names, 'Error': self.error_errors})
        df2.sort_values(by=['Name'], inplace=True)
        print(df1)
        print(df2)
        # df.to_excel('emails.xlsx', sheet_name='sheet1', index=False)
        with ExcelWriter("emails.xlsx") as writer:
            df1.to_excel(writer, sheet_name="Results", index=False)
            df2.to_excel(writer, sheet_name="Errors", index=False)

process = CrawlerProcess(settings={})
process.crawl(StudentSpider)
process.start()