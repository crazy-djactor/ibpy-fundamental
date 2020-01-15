# -*- coding: utf-8 -*-
import scrapy

class HkgetsymbolSpider(scrapy.Spider):
    name = 'hkgetsymbol'
    allowed_domains = ['www.interactivebrokers.com.hk/']
    start_urls = ['http://www.interactivebrokers.com.hk/']
    base_url = 'https://www.interactivebrokers.com.hk/en/index.php?f=2222&exch=sehk&showcategories=STK&p=&cc=&limit=100&page='

    def __init__(self, que):
        self.que = que
        # self.app = ReferenceApp1()
        # self.app.eConnect()

    def start_requests(self):
        page_num = 27
        for i in range(page_num):
            page_url = self.base_url + str(i + 1)
            yield scrapy.Request(page_url, callback=self.parse, meta={'index': i + 1})

    def parse(self, response):
        tr_list = response.xpath('//*[@id="exchange-products"]/div/div/div[3]/div/div/div/table/tbody/tr')
        for tr in tr_list:
            ibsymbol = tr.xpath('.//td[1]/text()').extract_first()
            company = tr.xpath('.//td[2]/a/text()').extract_first()
            currency = tr.xpath('.//td[4]/text()').extract_first()
            self.que.put({'ibsymbol': ibsymbol,
                          'company': company,
                          'currency': currency,
                          'index': response.meta['index']})

            # contract = Contract()
            #
            # contract.m_symbol = str(ibsymbol)
            # contract.m_currency = str(currency)
            # contract.m_secType = 'STK'
            # contract.m_exchange = 'SEHK'
            # self.app.reqFundamentalData('8001', contract, "RESC")
            # sleep(10)
            # fundamentaldata = self.app.wrapper.fundamental_Data_data
            # if fundamentaldata == '':
            #     continue
            # str_src = "{0}\\xml\\{1}.xml".format(self.base_path, str(ibsymbol))
            # f = open(str_src, 'w+')
            # f.write(fundamentaldata)
            # f.close()
            # str_dst = "{0}\\xlsx\\{1}.xlsx".format(self.base_path, str(ibsymbol))
            # copyfile(self.base_path + '\\Book2.xlsx',
            #          str_dst)
            # wb = self.excel.Workbooks.Open(str_dst)
            # hwnd = win32gui.FindWindow(None, "Microsoft Excel")  # Class or title
            # win32gui.ShowWindow(hwnd, win32con.SW_HIDE)  # Hide via Win32Api
            #
            # wb.XmlImport(str_src)
            # wb.SaveAs(str_dst)
            # wb.Close()
        pass
