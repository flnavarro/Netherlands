import os
import calendar
import scrapy
import time
import xlwt
from scrapy.exceptions import CloseSpider
from scrapy.http import HtmlResponse

import settings


class NpoRadio6Spider(scrapy.Spider):
    name = "NpoRadio6"

    def __init__(self, radio_station, day_begin, day_end):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.day = day_end

        # TODO: CHECK WHICH DAY IS FIRST
        if self.day_begin == '':
            self.day_begin = '00-00-0000'

        self.root_url = 'http://www.nposoulenjazz.nl/playlist'

        # self.urls = []
        # self.build_urls()

        self.all_tracks = []
        self.list_xls = xlwt.Workbook()
        self.sheet = self.list_xls.add_sheet(self.radio_station + 'Playlist')
        self.sheet.write(0, 0, 'Track Title')
        self.sheet.write(0, 1, 'Track Artist')
        self.sheet.write(0, 2, 'Count')
        self.sheet.write(0, 3, 'URL From')
        self.row = 1

    # TODO: CHANGE FOR BUILD FORM REQUESTS?
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

    # TODO: CHECK
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
        yield scrapy.Request(url=self.root_url, callback=self.form_requests, errback=self.parse_error)
        # for url in self.urls:
        #     yield scrapy.Request(url=url, callback=self.parse, errback=self.parse_error)
        #     # time.sleep(5)

    def form_requests(self, response):
        # TODO: MAKE A FOR HERE
        print('something')
        yield scrapy.FormRequest.from_response(response,
                                               formxpath="//form[@class='form_full'][@action='/playlist/zoeken#search']",
                                               formdata={"daletDay": "7",
                                                         "daletMonth": "08",
                                                         "daletYear": "2017",
                                                         "daletHour": "18"},
                                               clickdata={"type": "submit"},
                                               callback=self.parse)

    def parse(self, response):
        # TODO: Substitute response.url for DAY and HOUR
        # print('Getting tracks for url...' + response.url)

        titles = response.css('span.artist::text').extract()
        artists = response.css('span.track::text').extract()

        for title in titles:
            artist = artists[titles.index(title)].title()
            title = title.title()

            print('Appending Track -> ' + '"' + title + '"' + ' by ' + artist)
            self.sheet.write(self.row, 0, title)
            self.sheet.write(self.row, 1, artist)
            # self.sheet.write(self.row, 3, response.url)
            self.row += 1
            # self.all_tracks.append([title, artist, response.url])
            self.list_xls.save(self.radio_station + '.xls')

        # print('...tracks of url [ ' + response.url + ' ] finished.')

    def parse_error(self, response):
        print('There was an error.')

    def save_and_close_spider(self):
        pass