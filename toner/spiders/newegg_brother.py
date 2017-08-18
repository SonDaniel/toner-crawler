import scrapy
from pathlib import Path
import xlwings as xw
import re

class BrotherSpider(scrapy.Spider):
    name = "brother_toner"
    n = 1
    m = 1
    wb = None
    sh = None

    def start_requests(self):
        urls = [
            'https://www.newegg.com/Product/ProductList.aspx?Submit=ENE&N=50001816&IsNodeId=1&Description=brother%20toner&bop=And&Page=2&PageSize=96&order=BESTSELLING'
        ]

        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        model_name = response.url.split("/")[-1]

        # For excel
        self.wb = None
        excelFile = 'Brother_Newegg.xlsx'
        if Path(excelFile).exists():
            self.wb = xw.Book(excelFile)
        else:
            self.wb = xw.Book()
            
        self.sh = self.wb.sheets["BrotherToners"]

        
    #     for toner in response.css('div.ListLeft'):
    #         self.sh.range('A'+ str(self.n)).value = model_name
    #         yield scrapy.Request(url=toner.css('h2.product-name a::attr(href)').extract_first(), callback=self.parse_details)
        
    #     self.wb.save(excelFile)
    
    # def parse_details(self, response):
    #     regex = re.compile(r'[\n\r\t]')
    #     self.sh.range('B' + str(self.n)).value = response.css('div.product-name h1::text').extract_first()
    #     self.sh.range('C' + str(self.n)).value = response.css('div.long-description div.std::text').extract_first()
    #     self.sh.range('D' + str(self.n)).value = regex.sub(" ",response.css('div.price-box p.old-price span.price::text').extract_first())
    #     self.n += 1
