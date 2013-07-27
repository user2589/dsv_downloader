#!/usr/bin/python
# -*- coding: utf-8 -*-

import urllib2, cookielib
from urllib import urlencode
from optparse import make_option, OptionParser
import sys
from datetime import datetime, timedelta
import csv

try:
    from lxml import html
except ImportError:
    sys.stderr.write("""This script requires lxml. You can install it by either
    # easy_install lxml
or
    # apt-get install python-lxml
""")

DEFAULT_ACC = 'YOUR_ACC'
DEFAULT_PIN = 'YOUR_PIN'
DATE_FORMAT = "%m.%Y"
DEFAULT_DATE = (datetime.now() -timedelta(days=datetime.now().day+1)).strftime('%m.%Y')
LOGIN_URL       = 'https://issa.dsv.ru/Account/LogOnByAccount'
LOCAL_CALLS_URL = 'https://issa.dsv.ru/detail/apus'
EXT_CALLS_URL   = 'https://issa.dsv.ru/detail/mts'
REGIONS = (
        ('423', u'Приморский край - по умолчанию'),
        ('421', u'Хабаровский край'),
        ('424', u'Сахалинская область'),
        ('416', u'Амурская область'),
        ('415', u'Камчатская область'),
        ('413', u'Магаданская область'),
    )

class Downloader(object):
    def get_url(self, url, *args):
        try:
            response = self._opener.open(url, *args)
        except IOError:
            sys.stderr.write('Network error: '+url)
            exit(1)

        response_text = response.read().decode('cp1251')
        tree = html.fromstring(response_text)
        return tree

    def __init__(self, account, pin, region):
        #prepare cookie handler
        cookie_policy = cookielib.DefaultCookiePolicy(allowed_domains=['issa.dsv.ru',])
        cj = cookielib.CookieJar(cookie_policy)
        self._opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))

        post = urlencode({'Model.Region' : region,
                          'Model.Account': account,
                          'Model.Pin'    : pin})

        #Login in to get session cookies and store in _opener
        self.get_url(LOGIN_URL, post)

    def get_phones(self, month):
        url = '?'.join((LOCAL_CALLS_URL, urlencode({
                'month' : month.strftime('%m'),
                'year'  : month.strftime('%Y'),
                })))
        tree = self.get_url(url)
        return tree.xpath('//select[@name="Phone"]/option/@value')

    def _history(self, history_url, phone, month):
        url = '?'.join((history_url, urlencode({
                'month' : month.strftime('%m'),
                'year'  : month.strftime('%Y'),
                'phone' : phone,
                })))
        history = []
        page = 1
        while True:
            tree = self.get_url('&page='.join((url,str(page))))
            rows = tree.xpath('//form/table/tr/td/table/tr/td/div/table/tr')
            records = []
            for row in rows:
                cells = [c for c in row.xpath('./td/text()') if c.strip()]
                if cells and phone.endswith(cells[0]): #ie not line separator/header/sum/pager
                    records.append(cells)
            if not records: #ie page without records reached
                break
            history.extend(records)
            page += 1
        return history

    def local_history(self, phone, month):
        return self._history(LOCAL_CALLS_URL, phone, month)

    def ext_history(self, phone, month, year):
        return self._history(EXT_CALLS_URL, phone, month)

    def totals(self, phone, month):
        GET = urlencode({
                'month' : month.strftime('%m'),
                'year'  : month.strftime('%Y'),
                'phone' : phone,
                })
        tree = self.get_url('?'.join((LOCAL_CALLS_URL, GET)))
        record = [phone,]
        for row in tree.xpath('//form/table/tr/td/table/tr/td/div/table/tr'):
            cells = [c for c in row.xpath('./td/strong/text()') if c.strip()]
            if cells and cells[0] == u'ВСЕГО ЗА МЕСЯЦ:':
                record.extend([cells[1], cells[2]])
        tree = self.get_url('?'.join((EXT_CALLS_URL, GET)))
        for row in tree.xpath('//form/table/tr/td/table/tr/td/div/table/tr'):
            cells = [c for c in row.xpath('./td/strong/text()') if c.strip()]
            if cells and cells[0] == u'ВСЕГО ЗА МЕСЯЦ:':
                record.extend([cells[1], cells[2]])
        return record

if __name__ == "__main__":
    option_list = (
        make_option("-r", "--region", action="store", choices=dict(REGIONS).keys(), type="choice", dest="region", default='423',
                help = ("\t"*5).join([": ".join(r) for r in REGIONS])),
        make_option("-a", "--account", action="store", type="string", dest="account", default=DEFAULT_ACC,
                help = u"Номер счета в Ростелекоме (можно посмотреть в счете)"),
        make_option("-p", "--pin", action="store", type="string", dest="pin", default=DEFAULT_PIN,
                help = u"Пин-код ИССА, можно узнать в местном отделении РТ"),
        make_option("-d", "--date", action="store", type="string", dest="date", default=DEFAULT_DATE,
                help = u"Отчетный месяц в формате %s, например %s. Предыдущий месяц по умолчанию"%(DATE_FORMAT, datetime.now().strftime(DATE_FORMAT))),
        make_option("--total", action="store_const", dest="action", const='total',
                help = u"Получить только общуя статистику - сумма местных и междугородних звонков по телефонам"),
        make_option("--local", action="store_const", dest="action", const='local',
                help = u"Получить статистику только локальных звонков (по умолчанию)"),
        make_option("--external", "--ild", action="store_const", dest="action", const='ext',
                help = u"Получить статистику только звонков на межгород/мобильные"),
        make_option("--list", action="store_const", dest="action", const='phones',
                help = u"Получить список телефонов"),
        make_option("--phone", action="store", dest="phone",
                help = u"Получить статистику только по указанному номеру (номер нужно указывать вместе с кодом города)"),
            )
    parser = OptionParser(usage=u"""Usage: %prog -a <account_no> -p <pin> [options]
    Скрипт скачивает статистику по телефонии Ростелекома за месяц, указанный в параметре --date
    Выдача в CSV в стандартный вывод""",
                        option_list=option_list)
    parser.set_defaults(action="local")
    options, args = parser.parse_args()

    try:
        month = datetime.strptime(options.date, DATE_FORMAT)
    except ValueError:
        sys.stderr.write("Wrong date format")
        exit(2)

    downloader = Downloader(options.account, options.pin, options.region)

    out = csv.writer(sys.stdout)

    phones = downloader.get_phones(month)
    if options.phone:
        if options.phone in phones:
            phones = [options.phone,]
        else:
            sys.stderr.write("Указанного телефона нет в списке. Доступные номера:\n")
            sys.stderr.write("\n".join(phones))
            exit(3)

    if options.action == 'local':
        for phone in phones:
            for record in downloader.local_history(phone, month):
                out.writerow(record)
    elif options.action == 'ext':
        for phone in phones:
            for record in downloader.ext_history(phone, month):
                out.writerow(record)
    elif options.action == 'phones':
        for phone in phones:
            out.writerow([phone, ])
    elif options.action == 'total':
        out.writerow(['Телефон', 'локальные, сек', 'локальные, руб', 'межгород, мин', 'межгород, руб'])
        for phone in phones:
            out.writerow(downloader.totals(phone, month))
