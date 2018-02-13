# -*- coding: utf-8 -*-
import scrapy
import json
from scrapy.selector import Selector
from scrapy.utils.response import open_in_browser
from pprint import pprint


class BicyclesScrapeSpider(scrapy.Spider):
    name = 'bicycles_scrape'
    allowed_domains = ['electricbikereview.com/']
    start_urls = ['https://electricbikereview.com/wp-json/awesome-search/search/?object_class=AR_Review&group%5Btaxonomy%5D=review_cat&group%5Bterms%5D=mid-drive&group%5Bfield%5D=slug&args%5Bpost_types%5D%5Breview%5D%5Bslug%5D=review&args%5Bpost_types%5D%5Breview%5D%5Bobject_class%5D=AR_Review&args%5Bpost_types%5D%5Baccessory%5D%5Bslug%5D=accessory&args%5Bpost_types%5D%5Baccessory%5D%5Bobject_class%5D=AR_Accessory&args%5Bposts_per_page%5D=10&pagination%5BresultsPerPage%5D=4']

    def parse(self, response):
        print('!'*100)
        print(response)
        print('!'*100)
        # open_in_browser(response)
        # print(response.xpath('//table[@class="table-wrapper"]').extract())
