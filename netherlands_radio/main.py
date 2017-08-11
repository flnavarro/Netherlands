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

    def __init__(self, spider_name, radio_station, day_begin, day_end, repair_opt=False):
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
        self.repair_opt = repair_opt

    def get_lists(self):
        process = CrawlerProcess({
            'USER_AGENT': settings.USER_AGENT
        })
        process.crawl(self.spider, radio_station=self.radio_station,
                      day_begin=self.day_begin, day_end=self.day_end,
                      repair_opt=self.repair_opt)
        process.start()

    def count_plays(self):
        file_path = self.radio_station + '_raw.xls'
        if os.path.exists(file_path):
            xls = xlrd.open_workbook(file_path, formatting_info=True)
            n_rows = xls.sheet_by_index(0).nrows
            sheet_to_read = xls.sheet_by_index(0)
            track_list = []
            print('Counting plays...')
            for row in range(1, n_rows):
                title = sheet_to_read.cell(row, 0).value
                artist = sheet_to_read.cell(row, 1).value
                if len(track_list) == 0:
                    track_list.append([title, artist, 1])
                    # print(artist + ' - ' + title + ' // (FIRST TRACK on LIST)')
                already_in_list = False
                for track in track_list:
                    if track[0] == title and track[1] == artist:
                        already_in_list = True
                        break
                if already_in_list:
                    track_list[track_list.index(track)][2] += 1
                    # print(artist + ' - ' + title + ' // (PLAY COUNT) // Plays -> ' + str(track[2]))
                else:
                    track_list.append([title, artist, 1])
                    # print(artist + ' - ' + title + ' // (NEW TRACK on LIST)')

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
            list_xls.save(self.radio_station + '.xls')

        else:
            print('Could not find any file named: ' + self.radio_station + '_raw.xls ...')
            print('...so could not count plays!')

    def join_two_lists(self, folder_path, list_name, radio_station):
        self.radio_station = radio_station

        file_path_1_r = folder_path + list_name + '_raw_1.xls'
        file_path_2_r = folder_path + list_name + '_raw_2.xls'
        xls_1_r = xlrd.open_workbook(file_path_1_r)
        xls_2_r = xlrd.open_workbook(file_path_2_r)
        n_rows_1 = xls_1_r.sheet_by_index(0).nrows
        n_rows_2 = xls_2_r.sheet_by_index(0).nrows
        s_read_1 = xls_1_r.sheet_by_index(0)
        s_read_2 = xls_2_r.sheet_by_index(0)

        xls = xlwt.Workbook()
        sheet = xls.add_sheet(self.radio_station + 'Playlist')
        sheet.write(0, 0, 'Track Title')
        sheet.write(0, 1, 'Track Artist')
        sheet.write(0, 2, 'Play Count')
        sheet.write(0, 3, 'URL From')
        xls_row = 1

        for row in range(1, n_rows_1):
            title = s_read_1.cell(row, 0).value
            artist = s_read_1.cell(row, 1).value
            count = ''
            url = s_read_1.cell(row, 3).value

            sheet.write(xls_row, 0, title)
            sheet.write(xls_row, 1, artist)
            sheet.write(xls_row, 2, count)
            sheet.write(xls_row, 3, url)
            print('Writing data from RAW 1  //  row: ' + str(row) + '  //  from total rows: ' + str(n_rows_1))
            xls_row += 1

        for row in range(1, n_rows_2):
            title = s_read_2.cell(row, 0).value
            artist = s_read_2.cell(row, 1).value
            count = ''
            url = s_read_2.cell(row, 3).value

            sheet.write(xls_row, 0, title)
            sheet.write(xls_row, 1, artist)
            sheet.write(xls_row, 2, count)
            sheet.write(xls_row, 3, url)
            print('Writing data from RAW 2  //  row: ' + str(row) + '  //  from total rows: ' + str(n_rows_2))
            xls_row += 1

        xls.save(self.radio_station + '_raw.xls')
        self.count_plays()

    def add_count_from_raw(self, folder_path, list_name, radio_station):
        self.radio_station = radio_station

        file_count = folder_path + list_name + '_1.xls'
        file_raw = folder_path + list_name + '_raw_2.xls'
        xls_count = xlrd.open_workbook(file_count)
        xls_raw = xlrd.open_workbook(file_raw)
        n_rows_c = xls_count.sheet_by_index(0).nrows
        n_rows_r = xls_raw.sheet_by_index(0).nrows
        s_read_c = xls_count.sheet_by_index(0)
        s_read_r = xls_raw.sheet_by_index(0)

        xls = xlwt.Workbook()
        sheet = xls.add_sheet(self.radio_station + 'Playlist')
        sheet.write(0, 0, 'Track Title')
        sheet.write(0, 1, 'Track Artist')
        sheet.write(0, 2, 'Play Count')
        xls_row = 1

        tracks = []
        for row_c in range(1, n_rows_c):
            title = s_read_c.cell(row_c, 0).value
            artist = s_read_c.cell(row_c, 1).value
            count = int(s_read_c.cell(row_c, 2).value)
            tracks.append([title, artist, count])

        for row_r in range(1, n_rows_r):
            title = s_read_r.cell(row_r, 0).value
            artist = s_read_r.cell(row_r, 1).value
            already_in_list = False
            for track in tracks:
                if track[0] == title and track[1] == artist:
                    already_in_list = True
                    break
            if already_in_list:
                tracks[tracks.index(track)][2] += 1
                print(artist + ' - ' + title + ' // (PLAY COUNT) // Plays -> ' + str(tracks[tracks.index(track)][2]))
            else:
                tracks.append([title, artist, 1])
                print(artist + ' - ' + title + ' // (NEW TRACK on LIST)')

        for track in tracks:
            sheet.write(xls_row, 0, track[0])
            sheet.write(xls_row, 1, track[1])
            sheet.write(xls_row, 2, str(track[2]))
            xls_row += 1
        xls.save(self.radio_station + '.xls')

    def join_all_lists(self):
        path = 'DONE_'
        # stations = [
        #     '3fm',
        #     '100p',
        #     '538',
        #     'ClassicFm',
        #     'Drenthe',
        #     'Noord',
        #     'NpoRadio1',
        #     'NpoRadio2',
        #     'NpoRadio4',
        #     'NpoRadio5',
        #     'NpoRadio6',
        #     'qmusic',
        #     'radio10',
        #     'radionl',
        #     'skyradio',
        #     'slamfm',
        #     'sublimefm',
        #     'veronica'
        # ]
        stations = [
            'Fryslan'
        ]
        for station in stations:
            print('Starting list for radio -> ' + station)
            xls_final = xlwt.Workbook()
            sheet = xls_final.add_sheet(station + ' Playlist All Time')
            sheet.write(0, 0, 'Track Title')
            sheet.write(0, 1, 'Track Artist')
            sheet.write(0, 2, 'Play Count')
            xls_row = 1

            tracks = []

            year_int = 2017
            year = str(year_int)

            while True:
                folder_path = path + year + '/'
                file_path = folder_path + station + '_' + year + '.xls'
                if os.path.exists(file_path):
                    print('Getting list for radio -> ' + station + ' // For year -> ' + year)
                    xls = xlrd.open_workbook(file_path)
                    n_rows = xls.sheet_by_index(0).nrows
                    s_read = xls.sheet_by_index(0)

                    for row in range(1, n_rows):
                        title = s_read.cell(row, 0).value
                        artist = s_read.cell(row, 1).value
                        count = int(s_read.cell(row, 2).value)
                        already_in_list = False
                        for track in tracks:
                            if track[0] == title and track[1] == artist:
                                already_in_list = True
                                break
                        if already_in_list:
                            tracks[tracks.index(track)][2] += count
                        else:
                            tracks.append([title, artist, count])
                else:
                    break

                year_int -= 1
                year = str(year_int)

            list_index = 1
            for track in tracks:
                sheet.write(xls_row, 0, track[0])
                sheet.write(xls_row, 1, track[1])
                sheet.write(xls_row, 2, str(track[2]))
                xls_row += 1
                if xls_row == 65536:
                    print('Saving PART of all-time list for radio -> ' + station)
                    xls_final.save(station + '_ALL-TIME_' + str(list_index) + '.xls')
                    print('Starting new PART for all-time list for radio -> ' + station)
                    list_index += 1
                    xls_final = xlwt.Workbook()
                    sheet = xls_final.add_sheet(station + ' Playlist All Time (' + str(list_index) + ')')
                    sheet.write(0, 0, 'Track Title')
                    sheet.write(0, 1, 'Track Artist')
                    sheet.write(0, 2, 'Play Count')
                    xls_row = 1
            print('Saving all-time list for radio -> ' + station)
            if list_index == 1:
                xls_final.save(station + '_ALL-TIME.xls')
            else:
                xls_final.save(station + '_ALL-TIME_' + str(list_index) + '.xls')

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
# radio_station = station[7]

# spider_name = 'NpoRadio'
# station = ['NpoRadio1', 'NpoRadio2', 'NpoRadio4', 'NpoRadio5']
# radio_station = station[2]

# spider_name = 'ClassicFm'
# radio_station = 'classicfm'

# spider_name = 'Omroep'
# station = ['Brabant', 'Flevoland']
# radio_station = station[1]

spider_name = 'Fryslan'
radio_station = 'Fryslan'

# spider_name = 'NpoRadio6'
# radio_station = 'NpoRadio6'

# spider_name = 'Rtv'
# station = ['Drenthe', 'Noord']
# radio_station = station[1]

nl_crawler = NetherlandsCrawler(spider_name=spider_name, radio_station=radio_station,
                                day_begin='01-01-2015', day_end='30-06-2015',
                                repair_opt=False)
crawl = False
if crawl:
    nl_crawler.get_lists()
    nl_crawler.count_plays()
else:
    nl_crawler.join_all_lists()
    #nnl_crawler.add_count_from_raw('', 'Fryslan_2015', 'Fryslan_2015')

    # nl_crawler.add_count_from_raw('DONE_2010_JOIN/', '3fm_2010', '3fm_2010')
    # nl_crawler.add_count_from_raw('DONE_2010_JOIN/', '100p_2010', '100p_2010')
    # nl_crawler.add_count_from_raw('DONE_2010_JOIN/', '538_2010', '538_2010')
    # nl_crawler.add_count_from_raw('DONE_2010_JOIN/', 'qmusic_2010', 'qmusic_2010')
    # nl_crawler.add_count_from_raw('DONE_2010_JOIN/', 'slamfm_2010', 'slamfm_2010')



