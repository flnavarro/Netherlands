import argparse
import os
import xlrd, xlwt

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

    def __init__(self, spider_name, radio_station, day_begin, day_end):
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

    def get_lists(self):
        process = CrawlerProcess({
            'USER_AGENT': settings.USER_AGENT
        })
        process.crawl(self.spider, radio_station=self.radio_station,
                      day_begin=self.day_begin, day_end=self.day_end)
        process.start()

    def count_plays(self):
        file_path = self.radio_station + '.xls'
        xls = xlrd.open_workbook(file_path, formatting_info=True)
        n_rows = xls.sheet_by_index(0).nrows
        sheet_to_read = xls.sheet_by_index(0)
        track_list = []
        for row in range(1, n_rows):
            title = sheet_to_read.cell(row, 0).value
            artist = sheet_to_read.cell(row, 1).value
            if len(track_list) == 0:
                track_list.append([title, artist, 1])
                print(artist + ' - ' + title + ' // (FIRST TRACK on LIST)')
            already_in_list = False
            for track in track_list:
                if track[0] == title and track[1] == artist:
                    already_in_list = True
                    break
            if already_in_list:
                track_list[track_list.index(track)][2] += 1
                print(artist + ' - ' + title + ' // (PLAY COUNT) // Plays -> ' + str(track[2]))
            else:
                track_list.append([title, artist, 1])
                print(artist + ' - ' + title + ' // (NEW TRACK on LIST)')

        list_xls = xlwt.Workbook()
        sheet = list_xls.add_sheet(self.radio_station + ' Playlist')
        sheet.write(0, 0, 'Track Title')
        sheet.write(0, 1, 'Track Artist')
        sheet.write(0, 2, 'Play Count')
        row = 1
        for track in track_list:
            sheet.write(row, 0, track[0])
            sheet.write(row, 1, track[1])
            sheet.write(row, 2, str(track[2]))
            row += 1
        list_xls.save(self.radio_station + '_output.xls')


# TODO: Adapt to new parsing without batches
class InputParser(object):
    def __init__(self):
        self.parser = argparse.ArgumentParser(description='Download youtube audio tracks from DjChokaMusic.')
        self.batches_path = ''
        self.batch_size = 10

    def add_arguments(self):
        # Arguments to parse
        self.parser.add_argument('-batches_path', action='store', dest='batches_path', default='batches/',
                                 help='path for batches folder where the tracks will be downloaded')
        self.parser.add_argument('-batch_size', metavar='batch_size', type=int, default=10,
                                 help='number of tracks per batch')

    def parse_input(self):
        # Add arguments to parser
        self.add_arguments()

        # Parse arguments from input
        args = self.parser.parse_args()

        # Get Batches path and check if exists
        self.batches_path = args.batches_path
        if not self.batches_path == 'batches/':
            if not os.path.exists(self.batches_path):
                print('The path specified as batches_path does not exist.')
                if not os.path.exists('batches/'):
                    self.batches_path = 'batches/'
                    os.makedirs('batches/')
                    print('Batches folder created inside this script folder.')

        # Get batch size and check its value
        self.batch_size = args.batch_size
        if self.batch_size <= 0:
            print('Incorrect batch size. Setting batch size to a minimum of 10.')
            self.batch_size = 10

args_input = False

# TODO: Adapt to new parsing without batches
if args_input:
    # Get input from user and parse arguments
    input_parser = InputParser()
    input_parser.parse_input()
    # Get batch size
    batch_size = input_parser.batch_size
    # Get batches path
    batches_path = input_parser.batches_path
else:
    pass


# spider_name = 'Relisten'
# station = ['veronica', '3fm', 'skyradio', 'qmusic', '100p',
#            'radio10', '538', 'slamfm', 'sublimefm', 'radionl']
# radio_station = station[9]

# spider_name = 'NpoRadio'
# station = ['NpoRadio1', 'NpoRadio2', 'NpoRadio4', 'NpoRadio5']
# radio_station = station[3]

# spider_name = 'ClassicFm'
# radio_station = 'classicfm'

# spider_name = 'Omroep'
# station = ['Brabant', 'Flevoland']
# radio_station = station[1]

# spider_name = 'Fryslan'
# radio_station = 'Fryslan'

# spider_name = 'NpoRadio6'
# radio_station = 'NpoRadio6'

spider_name = 'Rtv'
station = ['Drenthe', 'Noord']
radio_station = station[0]

nl_crawler = NetherlandsCrawler(spider_name=spider_name, radio_station=radio_station,
                                day_begin='01-01-2017', day_end='31-07-2017')
nl_crawler.get_lists()
nl_crawler.count_plays()
