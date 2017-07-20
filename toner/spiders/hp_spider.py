import scrapy
from pathlib import Path
import xlwings as xw
import re

class BrotherSpider(scrapy.Spider):
    name = "hp_toner"
    n = 2
    wb = None
    sh = None

    def start_requests(self):
        last_row = str(xw.Book('HP.xlsx').sheets['HP Models'].range('A1').end('down').row)
        urls = xw.Book('HP.xlsx').sheets['HP Models'].range('B1:B' + last_row).value

        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        model_name = response.url.split("/")[-1]

        # For excel
        self.wb = None
        excelFile = 'HP.xlsx'
        if Path(excelFile).exists():
            self.wb = xw.Book(excelFile)
        else:
            self.wb = xw.Book()

        self.sh = self.wb.sheets["HP Toners"]

        # setting up headers
        self.sh.range('A1').value = 'HP Model'
        self.sh.range('B1').value = 'Toner Title'
        self.sh.range('C1').value = 'Description'
        self.sh.range('D1').value = 'Price'
        self.sh.range('E1').value = 'image link'

        self.sh.range('A'+ str(self.n)).value = model_name

        for toner in response.css('div.ListLeft'):
            yield scrapy.Request(url=toner.css('h2.product-name a::attr(href)').extract_first(), callback=self.parse_details)
        
        self.wb.save(excelFile)
    
    def parse_details(self, response):
        regex = re.compile(r'[\n\r\t]')
        self.sh.range('B' + str(self.n)).value = response.css('div.product-name h1::text').extract_first()
        self.sh.range('C' + str(self.n)).value = response.css('div.long-description div.std::text').extract_first()
        self.sh.range('D' + str(self.n)).value = regex.sub(" ",response.css('div.price-box p.old-price span.price::text').extract_first())
        self.sh.range('E' + str(self.n)).value = response.css('p.product-image img::attr(src)').extract_first()
        self.n += 1
