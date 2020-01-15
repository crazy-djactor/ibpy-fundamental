import os
import threading
from time import sleep

import pythoncom
import win32con
import win32gui
from scrapy.crawler import CrawlerProcess

from getsymbol.spiders.hkgetsymbol import HkgetsymbolSpider
import win32com.client as win32
import Queue
from reference_python import ReferenceApp
from ib.ext.Contract import Contract
from shutil import copyfile
import mysql.connector
import xml.etree.ElementTree as ET


class PutSQLWorker(threading.Thread):
    def __init__(self, qu, path):
        threading.Thread.__init__(self)
        self.qu = qu
        self.base_path = path

    def run(self):
        pythoncom.CoInitialize()
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        cnx = mysql.connector.connect(host='localhost', user='root', password='', database='ib')

        while True:
            try:
                data = self.qu.get()
            except Queue.Empty:
                continue
            str_src = "{0}\\xml\\{1}.xml".format(self.base_path, str(data["ibsymbol"]))
            f = open(str_src, 'w+')
            f.write(data["data"])
            f.close()

            tree = ET.fromstring(data["data"])
            company = tree.findall('./Company/CoName/Name')
            field0 = company[0].text
            field1 = company[0].attrib['type']
            security = tree.findall('./Company/SecurityInfo/Security')
            field4 = security[0].attrib['code']
            exchange = tree.findall('./Company/SecurityInfo/Security/Exchange')
            field5 = exchange[0].text
            field6 = exchange[0].attrib['code']
            country = tree.findall('./Company/SecurityInfo/Security/Country')
            field7 = country[0].text
            field8 = country[0].attrib['code']
            field9 = country[0].attrib['set']

            str_dst = "{0}\\xlsx\\{1}.xlsx".format(self.base_path, str(data["ibsymbol"]))
            copyfile(self.base_path + '\\Book2.xlsx',
                     str_dst)
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            wb = excel.Workbooks.Open(str_dst)
            hwnd = win32gui.FindWindow(None, "Microsoft Excel")  # Class or title
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)  # Hide via Win32Api
            wb.XmlImport(str_src)
            ws = wb.Worksheets("Sheet1")

            table_name = "ib_fundamentaldata{}".format(str(data["index"]))

            for v in ws.UsedRange.Value[1:]:
                params2 = [str(item) for item in v[2:]]
                params2.insert(0, field0)
                params2.insert(1, field1)
                params2[4] = field4
                params2[5] = field5
                params2[6] = field6
                params2[7] = field7
                params2[8] = field8
                params2[9] = field9

                cursor = cnx.cursor()
                sql = "INSERT INTO " + table_name + " VALUES ('" + "','".join(params2) + "');"
                cursor.execute(sql)

            cnx.commit()
            wb.SaveAs(str_dst)
            excel.ScreenUpdating = True
            excel.DisplayAlerts = True
            excel.EnableEvents = True
            wb.Close()
        pass
        excel.Application.Quit()


class EveryWorker(threading.Thread):

    def __init__(self, t_id, qu, sqlque):
        threading.Thread.__init__(self)
        self.t_id = t_id
        self.qu = qu
        self.sqlque = sqlque

    def run(self):
        fundamentaldata = ""
        app = ReferenceApp()
        app.eConnect()
        while True:
            # Get the work from the queue and expand the tuple
            try:
                company_info = self.qu.get()
            except Queue.Empty:
                break

            print(self.t_id, company_info["ibsymbol"], company_info["company"], company_info["currency"])

            contract = Contract()
            contract.m_symbol = str(company_info["ibsymbol"])
            contract.m_currency = str(company_info["currency"])
            contract.m_secType = 'STK'
            contract.m_exchange = 'SEHK'
            app.reqFundamentalData('8001', contract, "RESC")
            sleep(10)

            print("MYERROR-CODE-{} {}".format(app.wrapper.error_code, app.wrapper.error_msg))
            if app.wrapper.error_code == 430:
                # if fundamentaldata == app.wrapper.fundamental_Data_data:
                #     self.qu.put(company_info)
                self.qu.task_done()
                continue
            else:
                fundamentaldata = app.wrapper.fundamental_Data_data
                data = {
                    "ibsymbol": company_info["ibsymbol"],
                    "index": company_info["index"],
                    "data": fundamentaldata,
                }
                self.sqlque.put(data)
                self.qu.task_done()
        self.qu.task_done()
        app.eDisconnect()

def main():
    print("Great, you successfully entered an integer!")
    base_path = os.path.dirname(os.path.abspath(__file__))

    que = Queue.Queue()
    sqlQue = Queue.Queue()
    process = CrawlerProcess()
    # for n_num in range(n_start, n_end):
    #     company_num = n_num

    process.crawl(HkgetsymbolSpider, que=que)
    num_procs = 1
    sqlworker = PutSQLWorker(qu=sqlQue, path=base_path)
    sqlworker.daemon = True
    sqlworker.start()

    for x in range(num_procs):
        worker = EveryWorker(qu=que, sqlque=sqlQue, t_id=x)
        # Setting daemon to True will let the main thread exit even though the workers are blocking
        worker.daemon = True
        worker.start()
    process.start()
    que.join()


if __name__ == '__main__':
    main()
