import scrapy
import xlwt


class OmroepSpider(scrapy.Spider):
    name = "Omroep"

    def __init__(self, radio_station, day_begin, day_end, folder_path, repair_opt):
        self.radio_station = radio_station
        self.day_begin = day_begin
        self.day_end = day_end
        self.folder_path = folder_path

        if self.radio_station == 'Brabant':
            self.root_url = 'http://www.omroepbrabant.nl/Playlist.aspx?'
        elif self.radio_station == 'Flevoland':
            self.root_url = 'http://onlineradiobox.com/nl/omroepflev/playlist/'

        self.urls = []
        self.build_urls()

        self.all_tracks = []
        self.list_xls = None
        self.sheet = None
        self.row = 0

    def build_urls(self):
        print('Building urls...')
        if self.radio_station == 'Brabant':
            for day in range(1, 8):
                for part in range(1,5):
                    url = self.root_url + 'day=' + str(day) + '&part=' + str(part)
                    self.urls.append(url)
        elif self.radio_station == 'Flevoland':
            for day in range(0,7):
                url = self.root_url + str(day) + '?cs=nl.omroepflev'
                self.urls.append(url)

    def start_requests(self):
        for url in self.urls:
            yield scrapy.Request(url=url, callback=self.parse, errback=self.parse_error)

    def parse(self, response):
        print('Getting tracks for url...' + response.url)

        tracks = []
        artists = []
        if self.radio_station == 'Brabant':
            titles_and_times = response.css('li').css('div').css('div::text').extract()
            for title_or_time in titles_and_times:
                if title_or_time[:2] == ' -':
                    tracks.append(title_or_time[3:])
            artists = response.css('li').css('strong::text').extract()

        elif self.radio_station == 'Flevoland':
            if '0' in response.url:
                tracks_and_radios = response.css('td').css('a::text').extract()[1:]
            else:
                tracks_and_radios = response.css('td').css('a::text').extract()
            for track_or_radio in tracks_and_radios:
                if ' - ' in track_or_radio:
                    track = track_or_radio.split(' - ')
                    tracks.append(track[1].rstrip().lstrip())
                    artists.append(track[0].rstrip().lstrip())

        for track in tracks:
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

        print('...tracks of url [ ' + response.url + ' ] finished.')

    def parse_error(self, response):
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
        self.list_xls.save(self.folder_path + self.radio_station + '_7-Days' + '.xls')