from openpyxl import Workbook
import csv
import os.path
import scrapy
import glob
from scrapy.http import Request

# Min price is 10$
MIN_PRICE = 1000

# Max price is 30000$
MAX_PRICE = 3000000

# LIMIT
LIMIT = 100


class ItemsSpider(scrapy.Spider):
    name = 'items'
    allowed_domains = ['dmarket.com']

    def start_requests(self):
        yield Request(
            url='https://api.dmarket.com/exchange/v1/market/items?side=market&orderBy=price&orderDir=desc&title=&priceFrom='
                + str(MIN_PRICE) + '&priceTo='
                + str(MAX_PRICE)
                + '&treeFilters=&gameId=a8db&types=dmarket&cursor=&limit='
                + str(LIMIT)
                + '&currency=USD&platform=browser&isLoggedIn=false',
        )

    def parse(self, response):
        result = response.json()
        items = result['objects']
        if len(items) > 0:
            for item in items:
                title = item['title']
                discount = item['discount']
                price = item['price']['USD']

                yield {
                    'name': title,
                    'discount': discount,
                    'price': price,
                }

            cursor = result['cursor']
            yield Request(
                response.urljoin(
                    'items?side=market&orderBy=price&orderDir=desc&title=&priceFrom={}&priceTo={}&treeFilters=&gameId=a8db&types=dmarket&cursor={}&limit={}&currency=USD&platform=browser&isLoggedIn=false'.format(
                        MIN_PRICE, MAX_PRICE, cursor, LIMIT)),
                callback=self.parse,
            )

        def close(self, reason):
            csv_file = max(glob.iglob('*csv'), key=os.path.getctime)

            wb = Workbook()
            ws = wb.active

            with open(csv_file, 'r', encoding='utf-8') as f:
                for row in csv.reader(f):
                    ws.append(row)

            wb.save(csv_file.replace('.csv', '') + '.xlsx')
