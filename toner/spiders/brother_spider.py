import scrapy
from pathlib import Path
import xlwt

class BrotherSpider(scrapy.Spider):
    name = "brother_toner"

    def start_requests(self):
        f = open('brother_models.txt', 'r')
        for line in f:
            yield scrapy.Request(url=line.replace('\n', ''), callback=self.parse)

    def parse(self, response):
        page = response.url.split("/")[-1]

        # For excel
        wb = None
        excelFile = 'Brother.xlsx'
        if Path(excelFile).exists():
            wb = xw.Book(excelFile)
        else:
            wb = xw.Book()
            
        sht = wb.sheets.add(name="BrotherModels")

        data = []
        # for toner in response.css('div.ListLeft'):