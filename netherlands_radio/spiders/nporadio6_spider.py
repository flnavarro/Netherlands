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

        if self.day_begin == '':
            self.day_begin = '01-01-2016'

        self.root_url = 'http://www.nposoulenjazz.nl/playlist'

        self.forms = []
        self.build_forms()

        self.all_tracks = []
        self.list_xls = xlwt.Workbook()
        self.sheet = self.list_xls.add_sheet(self.radio_station + 'Playlist')
        self.sheet.write(0, 0, 'Track Title')
        self.sheet.write(0, 1, 'Track Artist')
        self.sheet.write(0, 2, 'Count')
        self.sheet.write(0, 3, 'URL From')
        self.row = 1

    def build_forms(self):
        print('Building forms...')
        is_within_requested_days = True
        while True:
            if is_within_requested_days:
                form_day = str(int(self.day[:2]))
                form_month = self.day[3:5]
                form_year = self.day[6:]
                for hour in range(0, 24):
                    form_hour = str(hour).zfill(2)
                    print('NEW FORM -> Day: ' + form_day + ' // Month: ' + form_month +
                          ' // Year: ' + form_year + ' // Hour: ' + form_hour)
                    form = [form_day, form_month, form_year, form_hour]
                    self.forms.append(form)
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
        yield scrapy.Request(url=self.root_url, callback=self.form_requests, errback=self.parse_error)

    def form_requests(self, response):
        for form in self.forms:
            yield scrapy.FormRequest.from_response(response,
                                                   formxpath="//form[@class='form_full'][@action='/playlist/zoeken#search']",
                                                   formdata={"daletDay": form[0],
                                                             "daletMonth": form[1],
                                                             "daletYear": form[2],
                                                             "daletHour": form[3]},
                                                   clickdata={"type": "submit"},
                                                   callback=self.parse)

    def parse(self, response):
        print('Getting tracks for form request... ' + response.request.body)

        titles = response.css('span.artist::text').extract()
        artists = response.css('span.track::text').extract()

        for title in titles:
            artist = artists[titles.index(title)].title()
            title = title.title()

            print('Appending Track -> ' + '"' + title + '"' + ' by ' + artist)
            self.sheet.write(self.row, 0, title)
            self.sheet.write(self.row, 1, artist)
            self.sheet.write(self.row, 3, response.request.body)
            self.row += 1
            self.all_tracks.append([title, artist, response.request.body])
            self.list_xls.save(self.radio_station + '.xls')

        print('...tracks of form request [ ' + response.request.body + ' ] finished.')

    def parse_error(self, response):
        print('There was an error.')

    def save_and_close_spider(self):
        pass