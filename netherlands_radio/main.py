import argparse
import os

from scrapy.crawler import CrawlerProcess

from spiders import settings
from spiders.relisten_spider import RelistenSpider
from spiders.nporadio_spider import NpoRadioSpider


class NetherlandsCrawler(object):

    def __init__(self, spider_name, radio_station, day_begin, day_end):
        self.spider = None
        if spider_name == 'Relisten':
            self.spider = RelistenSpider
        elif spider_name == 'NpoRadio':
            self.spider = NpoRadioSpider

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
# radio_station = station[0]

spider_name = 'NpoRadio'
station = ['NpoRadio1', '2', '4', '5', '6']
radio_station = station[0]

nl_crawler = NetherlandsCrawler(spider_name=spider_name, radio_station=radio_station,
                                day_begin='01-01-2017', day_end='31-07-2017')
nl_crawler.get_lists()
