import os
import calendar
import scrapy
import xlwt
import xlrd


class FryslanSpider(scrapy.Spider):
    name = "Fryslan"

    def __init__(self, radio_station, day_begin, day_end, folder_path, repair_opt=False):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.day = self.day_end
        self.folder_path = folder_path
        self.repair_opt = repair_opt

        if self.day_begin == '':
            self.day_begin = '02-02-2015'

        self.saving_code = self.day_begin[:2] + self.day_begin[3:5] + self.day_begin[6:] + '-' + \
                           self.day_end[:2] + self.day_end[3:5] + self.day_end[6:]

        self.root_url = 'http://www.omropfryslan.nl/utstjoering/'

        self.programmes = []
        self.get_programmes()

        self.months = []
        self.month_index = None
        self.get_months()

        self.all_urls = []

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

    def get_programmes(self):
        self.programmes = [
            # ['alvestedetochtrige', 'NONE'],
            # ['befrijingsfestival-fryslan', 'NONE'],
            ['buro-de-vries', '1100'],
            ['buro-de-vries', '2100'],
            # ['datwiedoe', 'NONE'],
            ['de-dei-foarby', '2200'],
            ['de-dei-foarby', '2230'],
            # ['de-flier-is-fan-jim', 'NONE'],
            # ['de-gouden-finylmiddei', 'NONE'],
            # ['de-grutte-nostalgy-show', 'NONE'],
            ['de-jun-fan-fryslan', '1900'],
            # ['de-krystdagen-fan-fryslan', 'NONE'],
            ['de-middei-fan-fryslan', '1300'],
            # ['de-rjochtsaak', 'NONE'],
            # ['de-top-fan-de-jierren-70', 'NONE'],
            # ['deadebetinking-ljouwert', 'NONE'],
            # ['fersykplaten', 'NONE'],
            # ['fierljep-kafee', 'NONE'],
            # ['fierljepkafee', 'NONE'],
            # ['finale-stimdei', 'NONE'],
            # ['fryske-muzykwike', 'NONE'],
            ['fryske-top-100', '1230'],
            # ['fryske-tsjerketsjinst', 'NONE'],
            # ['fryslan-aktueel-ferkiezings', 'NONE'],
            ['fryslan-fan-e-moarn', '0600'],
            # ['fryslan-kiest', 'NONE'],
            ['fryslan-nonstop', '0000'],
            ['fytsalvestedetocht', '0800'],
            # ['hjoed-1700-oere', 'NONE'],
            # ['in-fleanende-start', 'NONE'],
            ['it-jier-ut', '0900'],
            # ['jazzkafee-wijnbergen', 'NONE'],
            # ['junpraters', 'NONE'],
            ['koperkanaal-fm', '0600'],
            # ['krystkonsert-radio', 'NONE'],
            # ['linkk', 'NONE'],
            # ['lotting-pc', 'NONE'],
            # ['mei-douwe', 'NONE'],
            ['met-het-oog-op-morgen', '2300'],
            # ['muzyk-maskelyn', 'NONE'],
            ['muzyk-yn-bedriuw', '0900'],
            ['no-yn-fryslan', '0800'],
            ['no-yn-fryslan', '1200'],
            # ['noardewyn', 'NONE'],
            # ['noardewyn-live', 'NONE'],
            # ['oer-de-grins', 'NONE'],
            # ['omnium', 'NONE'],
            # ['omrop-fryslan-aktueel', 'NONE'],
            ['omrop-fryslan-sport', '1800'],
            # ['op-ult', 'NONE'],
            # ['pc-kafee', 'NONE'],
            # ['piter-wilkens-de-fleanende-hollanner', 'NONE'],
            # ['prelude', 'NONE'],
            ['radio-froskepole', '1800'],
            ['radio-froskepole', '0000'],
            # ['reade-tried', 'NONE'],
            # ['rients', 'NONE'],
            # ['rnc-jieroersjoch', 'NONE'],
            ['simmer-yn-fryslan', '0700'],
            ['simmer-yn-fryslan', '0900'],
            ['simmer-yn-fryslan', '1300'],
            # ['skutsjekafee', 'NONE'],
            ['snein-yn-fryslan', '1300'],
            ['sneon-yn-fryslan', '1300'],
            ['weistra-op-wei', '1600']
        ]

    def get_months(self):
        self.months = [
            'jannewaris',
            'febrewaris',
            'maart',
            'april',
            'maaie',
            'juny',
            'july',
            'augustus',
            'septimber',
            'oktober',
            'novimber',
            'desimber'
        ]

    def build_urls(self):
        print('Building urls...')
        is_within_requested_days = True
        while True:
            if is_within_requested_days:
                day = str(int(self.day[:2]))
                month = self.months[int(self.day[3:5])-1]
                year = self.day[6:]
                for programme in self.programmes:
                    hours = []
                    if programme[1] == 'NONE':
                        for hour in range(0, 24):
                            hour_min = str(hour).zfill(2) + '00'
                            hours.append(hour_min)
                    else:
                        hours.append(programme[1])

                    for hour in hours:
                        url = self.root_url + programme[0] + '-fan-' \
                              + day + '-' + month + '-' + year + '-' + hour
                        print('NEW URL -> ' + url)
                        self.all_urls.append(url)

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
                self.all_urls.append(sheet_read.cell(row, 0).value)

            if len(self.all_urls ) == 0:
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
                    self.all_urls = []
        else:
            print('Could not find a repair list.')

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
        for url in self.all_urls:
            yield scrapy.Request(url=url, callback=self.parse, errback=self.parse_error)

    def parse(self, response):
        print('Getting tracks for url...' + response.url)
        track_titles = response.css('div.field-name-field-uitzending-playlist').css('p::text').extract()[1::2]
        track_artists = response.css('div.field-name-field-uitzending-playlist').css('strong::text').extract()
        for title in track_titles:
            index = track_titles.index(title)
            artist = track_artists[index]
            title = title[3:]
            print('Appending Track -> ' + '"' + title + '"' + ' by ' + artist)

            already_in_tracks = False
            index_in_tracks = 0
            for track in self.all_tracks:
                if title == track[0] and artist == track[1]:
                    already_in_tracks = True
                    index_in_tracks = self.all_tracks.index(track)
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