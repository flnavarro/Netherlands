# -*- coding: utf-8 -*-

# Scrapy settings for netherlands_radio project
#
# For simplicity, this file contains only settings considered important or
# commonly used. You can find more settings consulting the documentation:
#
#     http://doc.scrapy.org/en/latest/topics/settings.html
#     http://scrapy.readthedocs.org/en/latest/topics/downloader-middleware.html
#     http://scrapy.readthedocs.org/en/latest/topics/spider-middleware.html

USER_AGENT = 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)'

# All tracks crawled
ALL_TRACKS = 'crawler_files/all_tracks.csv'
# Posts crawled in last execution
LAST_POSTS = 'crawler_files/last_posts.csv'
# New posts
NEW_POSTS = 'crawler_files/new_posts.csv'
# Embed urls for new posts
EMBED_URLS = 'crawler_files/embed_urls.csv'
# Download list of new posts with clean titles
DL_LIST = 'crawler_files/dl_list.csv'

BOT_NAME = 'netherlands_radio'
SPIDER_MODULES = ['netherlands_radio.spiders']
NEWSPIDER_MODULE = 'netherlands_radio.spiders'

# Obey robots.txt rules
ROBOTSTXT_OBEY = True
