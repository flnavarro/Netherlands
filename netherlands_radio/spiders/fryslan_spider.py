import os
import calendar
import scrapy
import xlwt
from scrapy.exceptions import CloseSpider

import settings


class FryslanSpider(scrapy.Spider):
    name = "Fryslan"

    def __init__(self, radio_station, day_begin, day_end):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.day = self.day_end

        # TODO: CHECK THIS!
        if self.day_begin == '':
            self.day_begin = 'NONE'

        self.root_url = 'http://www.omropfryslan.nl/utstjoering/'

        self.programmes = []
        self.get_programmes()

        self.months = []
        self.month_index = None
        self.get_months()

        self.all_urls = []
        self.valid_urls = []
        self.build_urls()

        self.all_tracks = []
        self.list_xls = xlwt.Workbook()
        self.sheet = self.list_xls.add_sheet(self.radio_station + 'Playlist')
        self.sheet.write(0, 0, 'Track Title')
        self.sheet.write(0, 1, 'Track Artist')
        self.sheet.write(0, 2, 'Count')
        self.sheet.write(0, 3, 'URL From')
        self.row = 1

    def get_programmes(self):
        self.programmes = [
            ['alvestedetochtrige', 'NONE'],
            ['befrijingsfestival-fryslan', 'NONE'],
            ['buro-de-vries', '1100'],
            ['buro-de-vries', '2100'],
            ['datwiedoe', 'NONE'],
            ['de-dei-foarby', '2200'],
            ['de-dei-foarby', '2230'],
            ['de-flier-is-fan-jim', 'NONE'],
            ['de-gouden-finylmiddei', 'NONE'],
            ['de-grutte-nostalgy-show', 'NONE'],
            ['de-jun-fan-fryslan', '1900'],
            ['de-krystdagen-fan-fryslan', 'NONE'],
            ['de-middei-fan-fryslan', '1300'],
            ['de-rjochtsaak', 'NONE'],
            ['de-top-fan-de-jierren-70', 'NONE'],
            ['deadebetinking-ljouwert', 'NONE'],
            ['fersykplaten', 'NONE'],
            ['fierljep-kafee', 'NONE'],
            ['fierljepkafee', 'NONE'],
            ['finale-stimdei', 'NONE'],
            ['fryske-muzykwike', 'NONE'],
            ['fryske-top-100', 'NONE'],
            ['fryske-tsjerketsjinst', 'NONE'],
            ['fryslan-aktueel-ferkiezings', 'NONE'],
            ['fryslan-fan-e-moarn', '0600'],
            ['fryslan-kiest', 'NONE'],
            ['fryslan-nonstop', '0000'],
            ['fytsalvestedetocht', '0800'],
            ['hjoed-1700-oere', 'NONE'],
            ['in-fleanende-start', 'NONE'],
            ['it-jier-ut', 'NONE'],
            ['jazzkafee-wijnbergen', 'NONE'],
            ['junpraters', 'NONE'],
            ['koperkanaal-fm', '0600'],
            ['krystkonsert-radio', 'NONE'],
            ['linkk', 'NONE'],
            ['lotting-pc', 'NONE'],
            ['mei-douwe', 'NONE'],
            ['met-het-oog-op-morgen', '2300'],
            ['muzyk-maskelyn', 'NONE'],
            ['muzyk-yn-bedriuw', '0900'],
            ['no-yn-fryslan', '0800'],
            ['no-yn-fryslan', '1200'],
            ['noardewyn', 'NONE'],
            ['noardewyn-live', 'NONE'],
            ['oer-de-grins', 'NONE'],
            ['omnium', 'NONE'],
            ['omrop-fryslan-aktueel', 'NONE'],
            ['omrop-fryslan-sport', '1800'],
            ['op-ult', 'NONE'],
            ['pc-kafee', 'NONE'],
            ['piter-wilkens-de-fleanende-hollanner', 'NONE'],
            ['prelude', 'NONE'],
            ['radio-froskepole', '1800'],
            ['radio-froskepole', '0000'],
            ['reade-tried', 'NONE'],
            ['rients', 'NONE'],
            ['rnc-jieroersjoch', 'NONE'],
            ['simmer-yn-fryslan', 'NONE'],
            ['skutsjekafee', 'NONE'],
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
            self.sheet.write(self.row, 0, title)
            self.sheet.write(self.row, 1, artist)
            self.sheet.write(self.row, 3, response.url)
            self.row += 1
            self.all_tracks.append([title, artist, response.url])
            self.list_xls.save(self.radio_station + '.xls')

        print('...tracks of url [ ' + response.url + ' ] finished.')

    def parse_error(self, response):
        print('There was an error.')

    def save_and_close_spider(self):
        pass