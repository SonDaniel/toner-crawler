import scrapy
from pathlib import Path
import xlwings as xw

class BrotherModelSpider(scrapy.Spider):
    name ="brother_model"

    def start_requests(self):
        urls = [
            'http://tonerprice.com/ink-toner/brother-ink-cartridges-toner'
        ]

        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        page = response.url.split("/")[-2]
        filename = 'brother_models.txt'
        hrefList = response.css('div.category_one ul li a::attr(href)').extract()

        # For excel
        wb = None
        sht = None
        excelFile = 'Brother.xlsx'
        if Path(excelFile).exists():
            wb = xw.Book(excelFile)
        else:
            wb = xw.Book()

        sht = wb.sheets["BrotherModels"]


        n = 1

        with open(filename, 'w') as f:
            for href in hrefList:
                f.write(str(href) + '\n')
                sht.range('A' + str(n)).value = str(href).split("/")[-1]
                sht.range('B' + str(n)).value = str(href)
                n+=1
        self.log('saved file %s' % filename)

        #Save Excel
        wb.save(excelFile)