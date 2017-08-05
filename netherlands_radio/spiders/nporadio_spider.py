import os
import calendar
import scrapy
import time
import xlwt
from scrapy.exceptions import CloseSpider
from scrapy.http import HtmlResponse

import settings


class NpoRadioSpider(scrapy.Spider):
    name = "NpoRadio"

    def __init__(self, radio_station, day_begin, day_end):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.day = str(int(self.day_end[6:])) + '-' + \
            str(int(self.day_end[3:5])).zfill(2) + '-' + \
            str(int(self.day_end[:2])).zfill(2)

        if self.radio_station == 'NpoRadio1':
            self.root_url = 'http://www.nporadio1.nl/muziek/'

        self.url = ''
        self.urls = []
        self.build_urls()

        self.all_tracks = []
        self.list_xls = xlwt.Workbook()
        self.sheet = self.list_xls.add_sheet(self.radio_station + 'Playlist')
        self.sheet.write(0, 0, 'Track Title')
        self.sheet.write(0, 1, 'Track Artist')
        self.sheet.write(0, 2, 'Count')
        self.sheet.write(0, 3, 'URL From')
        self.row = 1

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
        year_int = int(self.day[:4])
        month_int = int(self.day[5:7])
        day_int = int(self.day[8:])
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
            self.day = str(year_int) + '-' + str(month_int).zfill(2) + '-' + str(day_int).zfill(2)
            print('NEW DAY -> ' + self.day)
            is_within_requested_days = True

        return is_within_requested_days

    def start_requests(self):
        for url in self.urls:
            self.url = url
            yield scrapy.Request(url=url, callback=self.parse, errback=self.parse_error)
            time.sleep(5)
        self.list_xls.save(self.radio_station + '.xls')

    def parse(self, response):
        print('Getting tracks for url...' + self.url)
        tracks_data = response.css('div.responsive-item__content').css('h1.heading--small').extract()
        for track_data in tracks_data:
            new_response = HtmlResponse(url='My Url', body=track_data, encoding='utf-8')
            if len(new_response.css('a')) > 0:
                raw_title = new_response.css('a::text').extract_first()
            else:
                raw_title = new_response.css('h1.heading--small::text').extract_first()
            raw_title = raw_title.split('-')
            title = raw_title[0].rstrip().lstrip()
            artist = raw_title[1].rstrip().lstrip()
            print('Appending Track -> ' + '"' + title + '"' + ' by ' + artist)
            self.sheet.write(self.row, 0, title)
            self.sheet.write(self.row, 1, artist)
            self.sheet.write(self.row, 3, self.url)
            self.row += 1
            self.all_tracks.append([title, artist, self.url])
        print('...tracks of url [ ' + self.url + ' ] finished.')

    def parse_error(self, response):
        print('There was an error.')

    def save_and_close_spider(self):
        pass