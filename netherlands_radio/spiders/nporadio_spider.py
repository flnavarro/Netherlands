import os
import calendar
import scrapy
import xlwt
import xlrd
from scrapy.http import HtmlResponse


class NpoRadioSpider(scrapy.Spider):
    name = "NpoRadio"

    def __init__(self, radio_station, day_begin, day_end, folder_path, repair_opt=False):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.folder_path = folder_path
        self.repair_opt = repair_opt

        if self.day_begin == '':
            if self.radio_station == 'NpoRadio1':
                self.day_begin = '01-01-2016'
            elif self.radio_station == 'NpoRadio2' or self.radio_station == 'NpoRadio5':
                self.day_begin = '01-01-2015'
            elif self.radio_station == 'NpoRadio4':
                self.day_begin = '01-01-2012'

        self.saving_code = self.day_begin[:2] + self.day_begin[3:5] + self.day_begin[6:] + '-' + \
                           self.day_end[:2] + self.day_end[3:5] + self.day_end[6:]

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

        self.all_tracks = []
        self.list_xls = None
        self.sheet = None
        self.row = 0

        if not self.repair_opt:
            self.build_urls()
        else:
            self.build_repair()

        self.repair_xls = xlwt.Workbook()
        self.repair_sheet = self.repair_xls.add_sheet(self.radio_station + 'URL Repair List')
        self.repair_sheet.write(0, 0, 'URL to Repair')
        self.repair_row = 1

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

    def build_repair(self):
        repair_file_path = self.folder_path + self.radio_station + '_' + self.saving_code + '_repair.xls'
        if os.path.exists(repair_file_path):
            print('Loading repair list.')
            xls = xlrd.open_workbook(repair_file_path, formatting_info=True)
            n_rows = xls.sheet_by_index(0).nrows
            sheet_read = xls.sheet_by_index(0)
            for row in range(1, n_rows):
                self.urls.append(sheet_read.cell(row, 0).value)

            if len(self.urls) == 0:
                print('Could not find any url to repair')
            else:
                path = self.folder_path + self.radio_station + '_' + self.saving_code + '.xls'
                if os.path.exists(path):
                    print('Loading existing list for this radio station.')
                    xls = xlrd.open_workbook(path)
                    n_sheets = xls.nsheets
                    for sheet_index in range(0, n_sheets):
                        n_rows = xls.sheet_by_index(sheet_index).nrows
                        sheet_read = xls.sheet_by_index(sheet_index)
                        for row in range(1, n_rows):
                            title = sheet_read.cell(row, 0).value
                            artist = sheet_read.cell(row, 1).value
                            count = int(sheet_read.cell(row, 2).value)
                            self.all_tracks.append([title, artist, count])
                    os.remove(repair_file_path)
                else:
                    print('Could not find any list with the name: ' + path + '!')
                    self.urls = []
        else:
            print('Could not find a repair list.')

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

            already_in_tracks = False
            index_in_tracks = 0
            for track_in_list in self.all_tracks:
                if title == track_in_list[0] and artist == track_in_list[1]:
                    already_in_tracks = True
                    index_in_tracks = self.all_tracks.index(track_in_list)
                    break
            if already_in_tracks:
                self.all_tracks[index_in_tracks][2] += 1
            else:
                self.all_tracks.append([title, artist, 1])
        print('...tracks of url [ ' + response.url + ' ] finished.')
        self.save_list()

    def parse_error(self, response):
        if response.value.response.status == 408 or response.value.response.status == 500 \
                or response.value.response.status == 503:
            # Error 408 -> Request Timeout
            # Error 500 -> Internal Server Error
            # Error 503 -> Service Unavailable
            self.repair_sheet.write(self.repair_row, 0, response.request.url)
            self.repair_xls.save(self.folder_path + self.radio_station + '_repair.xls')
            self.repair_row += 1
            print('There was an error of type -> ' + str(response.value.response.status))
            print('Saved url to repair -> ' + str(response.url))
        else:
            # Other
            # Error 400 -> Bad Request
            # Error 404 -> Not Found
            print('There was an error of type -> ' + str(response.value.response.status))

    def save_list(self):
        self.list_xls = xlwt.Workbook()
        self.sheet = self.list_xls.add_sheet(self.radio_station + 'Playlist')
        self.sheet.write(0, 0, 'Track Title')
        self.sheet.write(0, 1, 'Track Artist')
        self.sheet.write(0, 2, 'Count')
        # self.sheet.write(0, 3, 'URL From')
        self.row = 1
        list_index = 1
        for track in self.all_tracks:
            self.sheet.write(self.row, 0, track[0])
            self.sheet.write(self.row, 1, track[1])
            self.sheet.write(self.row, 2, track[2])
            # self.sheet.write(self.row, 3, response.url)
            self.row += 1
            if self.row == 30001:
                list_index += 1
                self.sheet = self.list_xls.add_sheet(self.radio_station + 'Playlist (' + str(list_index) + ')')
                self.sheet.write(0, 0, 'Track Title')
                self.sheet.write(0, 1, 'Track Artist')
                self.sheet.write(0, 2, 'Play Count')
                self.row = 1
        if not self.repair_opt:
            self.list_xls.save(self.folder_path + self.radio_station + '_' + self.saving_code + '.xls')
        else:
            self.list_xls.save(self.folder_path + self.radio_station + '_' + self.saving_code + '_repaired.xls')
