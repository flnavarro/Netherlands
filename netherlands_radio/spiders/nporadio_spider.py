import os
import calendar
import scrapy
import time
import xlwt, xlrd
from scrapy.exceptions import CloseSpider
from scrapy.http import HtmlResponse

import settings


class NpoRadioSpider(scrapy.Spider):
    name = "NpoRadio"

    def __init__(self, radio_station, day_begin, day_end, repair_opt=False):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.repair_opt = repair_opt

        if self.day_begin == '':
            if self.radio_station == 'NpoRadio1':
                self.day_begin = '01-01-2016'
            elif self.radio_station == 'NpoRadio2' or self.radio_station == 'NpoRadio5':
                self.day_begin = '01-01-2015'
            elif self.radio_station == 'NpoRadio4':
                self.day_begin = '01-01-2012'

        if self.radio_station == 'NpoRadio1' or self.radio_station == 'NpoRadio4':
            if self.radio_station == 'NpoRadio1':
                self.root_url = 'http://www.nporadio1.nl/muziek/'
            elif self.radio_station == 'NpoRadio4':
                self.root_url = 'http://www.radio4.nl/speellijst/'

            self.day = str(int(self.day_end[6:])) + '-' + \
                str(int(self.day_end[3:5])).zfill(2) + '-' + \
                str(int(self.day_end[:2])).zfill(2)

        elif self.radio_station == 'NpoRadio2' or self.radio_station == 'NpoRadio5':
            if self.radio_station == 'NpoRadio2':
                self.root_url = 'http://www.nporadio2.nl/playlist?date='
            elif self.radio_station == 'NpoRadio5':
                self.root_url = 'http://www.nporadio5.nl/playlist?show=all&date='

            self.day = self.day_end

        self.urls = []
        self.build_urls()

        if self.repair_opt:
            self.urls_to_repair = []
            self.build_repair_urls()
            if len(self.urls_to_repair) == 0:
                print('Could not find any url to repair')

        self.all_tracks = []
        self.list_xls = xlwt.Workbook()
        self.sheet = self.list_xls.add_sheet(self.radio_station + 'Playlist')
        self.sheet.write(0, 0, 'Track Title')
        self.sheet.write(0, 1, 'Track Artist')
        self.sheet.write(0, 2, 'Count')
        self.sheet.write(0, 3, 'URL From')
        self.row = 1

    def build_repair_urls(self):
        raw_file_path = self.radio_station + '_repair.xls'
        if os.path.exists(raw_file_path):
            xls = xlrd.open_workbook(raw_file_path, formatting_info=True)
            n_rows = xls.sheet_by_index(0).nrows
            sheet_read = xls.sheet_by_index(0)
            urls_read = []
            for row in range(1, n_rows):
                url_from = sheet_read.cell(row, 3).value
                if url_from not in urls_read:
                    urls_read.append(url_from)
            for url in self.urls:
                if url not in urls_read:
                    print('Found url to repair: ' + url)
                    self.urls_to_repair.append(url)
            self.urls = self.urls_to_repair
        else:
            print('Could not find a repair list.')

    def build_urls(self):
        print('Building urls...')
        is_within_requested_days = True
        while True:
            if is_within_requested_days:
                url = self.root_url + self.day
                print('NEW URL -> ' + url)
                self.urls.append(url)
                is_within_requested_days = self.get_previous_day()
            else:
                break

    def get_previous_day(self):
        print('Getting previous day...')

        if self.radio_station == 'NpoRadio1' or self.radio_station == 'NpoRadio4':
            year_int = int(self.day[:4])
            month_int = int(self.day[5:7])
            day_int = int(self.day[8:])
        elif self.radio_station == 'NpoRadio2' or self.radio_station == 'NpoRadio5':
            day_int = int(self.day[:2])
            month_int = int(self.day[3:5])
            year_int = int(self.day[6:])

        if day_int > 1:
            day_int -= 1
        else:
            if month_int > 1:
                month_int -= 1
            else:
                year_int -= 1
                month_int = 12
            day_int = calendar.monthrange(year_int, month_int)[1]

        begin_day_int = int(self.day_begin[:2])
        begin_month_int = int(self.day_begin[3:5])
        begin_year_int = int(self.day_begin[6:])

        if year_int < begin_year_int \
            or (year_int == begin_year_int and month_int < begin_month_int) \
                or (year_int == begin_year_int and month_int == begin_month_int and day_int < begin_day_int):
            print('LIMIT OF DATES RANGE.')
            is_within_requested_days = False
        else:
            if self.radio_station == 'NpoRadio1' or self.radio_station == 'NpoRadio4':
                self.day = str(year_int) + '-' + str(month_int).zfill(2) + '-' + str(day_int).zfill(2)
            elif self.radio_station == 'NpoRadio2' or self.radio_station == 'NpoRadio5':
                self.day = str(day_int).zfill(2) + '-' + str(month_int).zfill(2) + '-' + str(year_int)
            print('NEW DAY -> ' + self.day)
            is_within_requested_days = True

        return is_within_requested_days

    def start_requests(self):
        for url in self.urls:
            yield scrapy.Request(url=url, callback=self.parse, errback=self.parse_error)
            # time.sleep(5)

    def parse(self, response):
        print('Getting tracks for url...' + response.url)

        if self.radio_station == 'NpoRadio1':
            tracks = response.css('div.responsive-item__content').css('h1.heading--small').extract()
        elif self.radio_station == 'NpoRadio2':
            tracks = response.css('p.fn-song::text').extract()
            artists = response.css('p.fn-artist::text').extract()
        elif self.radio_station == 'NpoRadio4':
            tracks = response.css('div.l-content').css('span.title::text').extract()
            artists = response.css('div.l-content').css('strong::text').extract()
        elif self.radio_station == 'NpoRadio5':
            tracks = response.css('p.fn-song::text').extract()
            artists = response.css('h5.fn-artist::text').extract()

        for track in tracks:
            if self.radio_station == 'NpoRadio1':
                new_response = HtmlResponse(url='My Url', body=track, encoding='utf-8')
                if len(new_response.css('a')) > 0:
                    raw_title = new_response.css('a::text').extract_first()
                else:
                    raw_title = new_response.css('h1.heading--small::text').extract_first()
                raw_title = raw_title.split('-')
                title = raw_title[0].rstrip().lstrip()
                artist = raw_title[1].rstrip().lstrip()
            elif self.radio_station == 'NpoRadio2' \
                    or self.radio_station == 'NpoRadio4' \
                    or self.radio_station == 'NpoRadio5':
                title = track
                artist = artists[tracks.index(track)]

            print('Appending Track -> ' + '"' + title + '"' + ' by ' + artist)
            self.sheet.write(self.row, 0, title)
            self.sheet.write(self.row, 1, artist)
            self.sheet.write(self.row, 3, response.url)
            self.row += 1
            self.all_tracks.append([title, artist, response.url])
        self.list_xls.save(self.radio_station + '_raw.xls')

        print('...tracks of url [ ' + response.url + ' ] finished.')

    def parse_error(self, response):
        print('There was an error.')

    def save_and_close_spider(self):
        pass