import argparse
import os
import xlrd
import xlwt

from scrapy.crawler import CrawlerProcess

from spiders import settings
from spiders.relisten_spider import RelistenSpider
from spiders.nporadio_spider import NpoRadioSpider
from spiders.classicfm_spider import ClassicFmSpider
from spiders.omroep_spider import OmroepSpider
from spiders.fryslan_spider import FryslanSpider
from spiders.nporadio6_spider import NpoRadio6Spider
from spiders.rtv_spider import RtvSpider


class NetherlandsCrawler(object):

    def __init__(self, spider_name, radio_station, day_begin, day_end, folder_path, repair_opt=False):
        self.spider = None
        if spider_name == 'Relisten':
            self.spider = RelistenSpider
        elif spider_name == 'NpoRadio':
            self.spider = NpoRadioSpider
        elif spider_name == 'ClassicFm':
            self.spider = ClassicFmSpider
        elif spider_name == 'Omroep':
            self.spider = OmroepSpider
        elif spider_name == 'Fryslan':
            self.spider = FryslanSpider
        elif spider_name == 'NpoRadio6':
            self.spider = NpoRadio6Spider
        elif spider_name == 'Rtv':
            self.spider = RtvSpider

        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        if folder_path == '':
            self.folder_path = folder_path
        else:
            self.folder_path = folder_path + '/'
        self.repair_opt = repair_opt

    def get_lists(self):
        process = CrawlerProcess({
            'USER_AGENT': settings.USER_AGENT
        })
        process.crawl(self.spider, radio_station=self.radio_station,
                      day_begin=self.day_begin, day_end=self.day_end,
                      folder_path=self.folder_path,
                      repair_opt=self.repair_opt)
        process.start()


class InputParser(object):
    def __init__(self):
        self.parser = argparse.ArgumentParser(description='Download playlists from Dutch radio stations.')
        self.path = ''
        self.station = ''
        self.day_begin = ''
        self.day_end = ''
        self.repair_opt = False

    def add_arguments(self):
        # Arguments to parse
        self.parser.add_argument('-path', action='store', dest='path', default='',
                                 help='Path where: lists will be stored & repair lists will be read')
        self.parser.add_argument('-station', action='store', dest='station', required=True,
                                 help='Available radio stations [REQUIRED]: '
                                      'veronica, 3fm, skyradio, qmusic, 100p, '
                                      'radio10, 538, slamfm, sublimefm, radionl, '
                                      'nporadio1, nporadio2, nporadio4, nporadio5, nporadio6, '
                                      'classicfm, brabant, flevoland, fryslan, '
                                      'rtvdrenthe, rtvnoord')
        self.parser.add_argument('-day_begin', action='store', dest='day_begin', default='',
                                 help='First day for retrieving playlists: DD-MM-YYYY \n'
                                      '[When not specified retrieves playlists from first day possible.]')
        self.parser.add_argument('-day_end', action='store', dest='day_end', required=True,
                                 help='Last day for retrieving playlists [REQUIRED]: DD-MM-YYYY')
        self.parser.add_argument('-repair_opt', action='store_true', dest='repair_opt', default=False,
                                 help='Option to retrieve playlists from missing urls (repair file needed)')

    def parse_input(self):
        # Add arguments to parser
        self.add_arguments()

        # Parse arguments from input
        args = self.parser.parse_args()

        # Get Batches path and check if exists
        self.path = args.path
        self.station = args.station
        self.day_begin = args.day_begin
        self.day_end = args.day_end
        self.repair_opt = args.repair_opt

args_input = True

if args_input:
    # Get input from user and parse arguments
    input_parser = InputParser()
    input_parser.parse_input()

    # Get variables from input
    path = input_parser.path
    day_begin = input_parser.day_begin
    day_end = input_parser.day_end
    repair_opt = input_parser.repair_opt
    station = input_parser.station

    # Get spider name and radio station
    spider_name = ''
    radio_station = ''
    if station == 'veronica' or station == '3fm' or station == 'skyradio' or station == 'qmusic' \
        or station == '100p' or station == 'radio10' or station == '538' or station == 'slamfm' \
            or station == 'sublimefm' or station == 'radionl':
        spider_name = 'Relisten'
        radio_station = station
    elif station == 'nporadio1' or station == 'nporadio2' or station == 'nporadio4' or station == 'nporadio5':
        spider_name = 'NpoRadio'
        if station == 'nporadio1':
            radio_station = 'NpoRadio1'
        elif station == 'nporadio2':
            radio_station = 'NpoRadio2'
        elif station == 'nporadio4':
            radio_station = 'NpoRadio4'
        elif station == 'nporadio5':
            radio_station = 'NpoRadio5'
    elif station == 'nporadio6':
        spider_name = 'NpoRadio6'
        radio_station = spider_name
    elif station == 'classicfm':
        spider_name = 'ClassicFm'
        radio_station = station
    elif station == 'brabant' or station == 'flevoland':
        spider_name = 'Omroep'
        if station == 'brabant':
            radio_station = 'Brabant'
        elif station == 'flevoland':
            radio_station = 'Flevoland'
    elif station == 'fryslan':
        spider_name = 'Fryslan'
        radio_station = 'Fryslan'
    elif station == 'rtvdrenthe' or station == 'rtvnoord':
        spider_name = 'Rtv'
        if station == 'rtvdrenthe':
            radio_station = 'Drenthe'
        elif station == 'rtvnoord':
            radio_station = 'Noord'

    # Get lists from crawler
    nl_crawler = NetherlandsCrawler(spider_name=spider_name, radio_station=radio_station,
                                    day_begin=day_begin, day_end=day_end,
                                    folder_path=path,
                                    repair_opt=repair_opt)
    nl_crawler.get_lists()
else:
    path = ''
    spider_name = 'Rtv'
    station = ['Drenthe', 'Noord']
    radio_station = station[1]

    nl_crawler = NetherlandsCrawler(spider_name=spider_name, radio_station=radio_station,
                                    day_begin='01-01-2017', day_end='31-07-2017',
                                    folder_path=path,
                                    repair_opt=False)
    nl_crawler.get_lists()


# spider_name = 'Relisten'
# station = ['veronica', '3fm', 'skyradio', 'qmusic', '100p',
#            'radio10', '538', 'slamfm', 'sublimefm', 'radionl']
# radio_station = station[1]

# spider_name = 'NpoRadio'
# station = ['NpoRadio1', 'NpoRadio2', 'NpoRadio4', 'NpoRadio5']
# radio_station = station[2]

# spider_name = 'ClassicFm'
# radio_station = 'classicfm'

# spider_name = 'Omroep'
# station = ['Brabant', 'Flevoland']
# radio_station = station[1]

# spider_name = 'Fryslan'
# radio_station = 'Fryslan'

# spider_name = 'NpoRadio6'
# radio_station = 'NpoRadio6'
