import os
import calendar
import scrapy
import time
import xlwt, xlrd
import json
from scrapy.exceptions import CloseSpider
from scrapy.http import HtmlResponse

import settings


class RtvSpider(scrapy.Spider):
    name = "Rtv"

    def __init__(self, radio_station, day_begin, day_end, repair_opt=False):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.day = day_end
        self.repair_opt = repair_opt

        if self.radio_station == 'Drenthe':
            self.root_url = 'http://www.rtvdrenthe.nl/ajax/RadioTv/GetPlaylistDay?date='
            if self.day_begin == '':
                self.day_begin = '04-03-2016'
        elif self.radio_station == 'Noord':
            self.root_url = 'http://www.rtvnoord.nl/ajax/RadioTv/GetPlaylistDay?date='
            if self.day_begin == '':
                self.day_begin = '12-10-2015'

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
                day_url = self.day[3:5] + '-' + self.day[:2] + '-' + self.day[6:]
                url = self.root_url + day_url
                print('NEW URL -> ' + url)
                self.urls.append(url)
                is_within_requested_days = self.get_previous_day()
            else:
                break

    def get_previous_day(self):
        print('Getting previous day...')

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
        day_data = json.loads(response.body)
        for part_data in day_data:
            for hour in part_data['hours']:
                for track in hour['tracks']:
                    title = track['title'].title()
                    artist = track['artist'].title()

                    print('Appending Track -> ' + '"' + title + '"' +
                          ' by ' + artist)
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