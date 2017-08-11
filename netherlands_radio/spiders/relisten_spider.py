import os
import calendar
import scrapy
import xlwt, xlrd
from scrapy.exceptions import CloseSpider

import settings


class RelistenSpider(scrapy.Spider):
    name = "Relisten"

    def __init__(self, radio_station, day_begin, day_end, repair_opt=False):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.day = self.day_end
        self.repair_opt = repair_opt

        if self.day_begin == '':
            if self.radio_station == 'veronica' or self.radio_station == 'skyradio':
                self.day_begin = '10-08-2010'
            elif self.radio_station == '3fm' or self.radio_station == 'qmusic' \
                    or self.radio_station == '100p' or self.radio_station == '538':
                self.day_begin = '01-05-2010'
            elif self.radio_station == 'radio10':
                self.day_begin = '19-02-2014'
            elif self.radio_station == 'slamfm':
                self.day_begin = '30-04-2010'
            elif self.radio_station == 'sublimefm':
                self.day_begin = '06-04-2014'
            elif self.radio_station == 'radionl':
                self.day_begin = '28-12-2011'

        self.root_url = 'https://www.relisten.nl/playlists/' + self.radio_station + '/'

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
                url = self.root_url + self.day + '.html'
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

    def parse(self, response):
        print('Getting tracks for url...' + response.url)
        track_titles = response.css('li.media').css('h4.media-heading').css('span::text').extract()
        track_artists = response.css('li.media').css('a').css('span::text').extract()
        for title in track_titles:
            index = track_titles.index(title)
            artist = track_artists[index]
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