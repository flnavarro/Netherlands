import os
import calendar
import scrapy
import time
import xlwt, xlrd
from scrapy.exceptions import CloseSpider
from scrapy.http import HtmlResponse

import settings


class NpoRadio6Spider(scrapy.Spider):
    name = "NpoRadio6"

    def __init__(self, radio_station, day_begin, day_end, repair_opt=False):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.day = day_end
        self.repair_opt = repair_opt

        if self.day_begin == '':
            self.day_begin = '01-01-2016'

        self.root_url = 'http://www.nposoulenjazz.nl/playlist'

        self.forms = []
        self.build_forms()

        if self.repair_opt:
            self.forms_to_repair = []
            self.build_repair_forms()
            if len(self.forms_to_repair) == 0:
                print('Could not find any url to repair')

        self.all_tracks = []
        self.list_xls = xlwt.Workbook()
        self.sheet = self.list_xls.add_sheet(self.radio_station + 'Playlist')
        self.sheet.write(0, 0, 'Track Title')
        self.sheet.write(0, 1, 'Track Artist')
        self.sheet.write(0, 2, 'Count')
        self.sheet.write(0, 3, 'URL From')
        self.row = 1

    def build_repair_forms(self):
        raw_file_path = self.radio_station + '_repair.xls'
        if os.path.exists(raw_file_path):
            xls = xlrd.open_workbook(raw_file_path, formatting_info=True)
            n_rows = xls.sheet_by_index(0).nrows
            sheet_read = xls.sheet_by_index(0)
            forms_read = []
            for row in range(1, n_rows):
                form = sheet_read.cell(row, 3).value
                if form not in forms_read:
                    forms_read.append(form)
            for form in self.forms:
                string_form = 'daletMonth=' + form[1] + \
                               '&daletYear=' + form[2] + \
                               '&daletHour=' + form[3] + \
                               '&daletDay=' + form[0]
                if string_form not in forms_read:
                    print('Found form for url to repair: ' + string_form)
                    self.forms_to_repair.append(form)
            self.forms = self.forms_to_repair
        else:
            print('Could not find a repair list.')

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

        titles = response.css('span.track::text').extract()
        artists = response.css('span.artist::text').extract()

        for title in titles:
            artist = artists[titles.index(title)].title()
            title = title.title()

            print('Appending Track -> ' + '"' + title + '"' + ' by ' + artist)
            self.sheet.write(self.row, 0, title)
            self.sheet.write(self.row, 1, artist)
            self.sheet.write(self.row, 3, response.request.body)
            self.row += 1
            self.all_tracks.append([title, artist, response.request.body])
        self.list_xls.save(self.radio_station + '_raw.xls')

        print('...tracks of form request [ ' + response.request.body + ' ] finished.')

    def parse_error(self, response):
        print('There was an error.')

    def save_and_close_spider(self):
        pass