#!/usr/bin/env python
# -*- coding: utf-8 -*-

##
# This script is an exmple of using the generated code within IbPy in
# the same manner as the Java code.  We subclass EWrapper and give an
# instance of the wrapper to an EClientSocket.
##
import time
from sys import argv
from time import sleep, strftime

from ib.ext.Contract import Contract
from ib.ext.EWrapper import EWrapper
from ib.ext.EClientSocket import EClientSocket
from ib.ext.ExecutionFilter import ExecutionFilter


def showmessage(message, mapping):
    try:
        del(mapping['self'])
    except (KeyError, ):
        pass
    items = list(mapping.items())
    items.sort()
    print(('### %s' % (message, )))
    for k, v in items:
        print(('    %s:%s' % (k, v)))


class ReferenceWrapper(EWrapper):
    def __init__(self):
        self.fundamental_Data_data = ""
        self.fundamental_Data_reqId = ""
        self.error_code = 0
        self.error_msg = ""

    def tickPrice(self, tickerId, field, price, canAutoExecute):
        showmessage('tickPrice', vars())

    def tickSize(self, tickerId, field, size):
        showmessage('tickSize', vars())

    def tickOptionComputation(self, tickerId, field, impliedVol, delta, optPrice, pvDividend, gamma, vega, theta, undPrice):
        showmessage('tickOptionComputation', vars())

    def tickGeneric(self, tickerId, tickType, value):
        showmessage('tickGeneric', vars())

    def tickString(self, tickerId, tickType, value):
        showmessage('tickString', vars())

    def tickEFP(self, tickerId, tickType, basisPoints, formattedBasisPoints, impliedFuture, holdDays, futureExpiry, dividendImpact, dividendsToExpiry):
        showmessage('tickEFP', vars())

    def orderStatus(self, orderId, status, filled, remaining, avgFillPrice, permId, parentId, lastFillPrice, clientId, whyHeId):
        showmessage('orderStatus', vars())

    def openOrder(self, orderId, contract, order, state):
        showmessage('openOrder', vars())

    def openOrderEnd(self):
        showmessage('openOrderEnd', vars())

    def updateAccountValue(self, key, value, currency, accountName):
        showmessage('updateAccountValue', vars())

    def updatePortfolio(self, contract, position, marketPrice, marketValue, averageCost, unrealizedPNL, realizedPNL, accountName):
        showmessage('updatePortfolio', vars())

    def updateAccountTime(self, timeStamp):
        showmessage('updateAccountTime', vars())

    def accountDownloadEnd(self, accountName):
        showmessage('accountDownloadEnd', vars())

    def nextValidId(self, orderId):
        showmessage('nextValidId', vars())

    def contractDetails(self, reqId, contractDetails):
        showmessage('contractDetails', vars())

    def contractDetailsEnd(self, reqId):
        showmessage('contractDetailsEnd', vars())

    def bondContractDetails(self, reqId, contractDetails):
        showmessage('bondContractDetails', vars())

    def execDetails(self, reqId, contract, execution):
        showmessage('execDetails', vars())

    def execDetailsEnd(self, reqId):
        showmessage('execDetailsEnd', vars())

    def connectionClosed(self):
        showmessage('connectionClosed', {})

    def error(self, id=None, errorCode=None, errorMsg=None):
        showmessage('error', vars())
        self.error_code = errorCode
        self.error_msg = errorMsg

    def error_0(self, strvalue=None):
        showmessage('error_0', vars())

    def error_1(self, id=None, errorCode=None, errorMsg=None):
        showmessage('error_1', vars())

    def updateMktDepth(self, tickerId, position, operation, side, price, size):
        showmessage('updateMktDepth', vars())

    def updateMktDepthL2(self, tickerId, position, marketMaker, operation, side, price, size):
        showmessage('updateMktDepthL2', vars())

    def updateNewsBulletin(self, msgId, msgType, message, origExchange):
        showmessage('updateNewsBulletin', vars())

    def managedAccounts(self, accountsList):
        showmessage('managedAccounts', vars())

    def receiveFA(self, faDataType, xml):
        showmessage('receiveFA', vars())

    def historicalData(self, reqId, date, open, high, low, close, volume, count, WAP, hasGaps):
        showmessage('historicalData', vars())

    def scannerParameters(self, xml):
        showmessage('scannerParameters', vars())

    def scannerData(self, reqId, rank, contractDetails, distance, benchmark, projection, legsStr):
        showmessage('scannerData', vars())

    def accountDownloadEnd(self, accountName):
        showmessage('acountDownloadEnd', vars())

    def commissionReport(self, commissionReport):
        showmessage('commissionReport', vars())

    def contractDetailsEnd(self, reqId):
        showmessage('contractDetailsEnd', vars())

    def currentTime(self, time):
        showmessage('currentTime', vars())

    def deltaNeutralValidation(self, reqId, underComp):
        showmessage('deltaNeutralValidation', vars())

    def execDetailsEnd(self, reqId):
        showmessage('execDetailsEnd', vars())

    def fundamentalData(self, reqId, data):
        self.fundamental_Data_data = data
        self.fundamental_Data_reqId = reqId
        self.error_code = 0
        self.error_msg = "OK"
        showmessage('fundamentalData', vars())

    def marketDataType(self, reqId, marketDataType):
        showmessage('marketDataType', vars())

    def openOrderEnd(self):
        showmessage('openOrderEnd', vars())

    def realtimeBar(self, reqId, time, open, high, low, close, volume, wap, count):
        showmessage('realtimeBar', vars())

    def scannerDataEnd(self, reqId):
        showmessage('scannerDataEnd', vars())

    def tickEFP(self, tickerId, tickType, basisPoints, formattedBasisPoints, impliedFuture, holdDays, futureExpiry, dividendImpact, dividendsToExpiry):
        showmessage('tickEFP', vars())

    def tickGeneric(self, tickerId, tickType, value):
        showmessage('tickGeneric', vars())

    def tickSnapshotEnd(self, reqId):
        showmessage('tickSnapshotEnd', vars())

    def error_0(self, strval):
        showmessage('error_0', vars())

    def error_1(self, id, errorCode, errorMsg):
        showmessage('error_1', vars())

    def position(self, account, contract, pos, avgCost):
        showmessage('position', vars())

    def positionEnd(self):
        showmessage('positionEnd', vars())

    def accountSummary(self, reqId, account, tag, value, currency):
        showmessage('accountSummary', vars())

    def accountSummaryEnd(self, reqId):
        showmessage('accountSummaryEnd', vars())

allMethods = []
def ref(method):
    allMethods.append(method.__name__)
    return method


class ReferenceApp:
    def __init__(self, host='localhost', port=7496, clientId=999):
        self.host = host
        self.port = port
        self.clientId = clientId
        self.wrapper = ReferenceWrapper()
        self.connection = EClientSocket(self.wrapper)

    @ref
    def eConnect(self):
        self.connection.eConnect(self.host, self.port, self.clientId)

    @ref
    def reqAccountUpdates(self):
        self.connection.reqAccountUpdates(1, '')

    @ref
    def reqOpenOrders(self):
        self.connection.reqOpenOrders()

    @ref
    def reqExecutions(self):
        filt = ExecutionFilter()
        self.connection.reqExecutions(0, filt)

    @ref
    def reqIds(self):
        self.connection.reqIds(10)

    @ref
    def reqNewsBulletins(self):
        self.connection.reqNewsBulletins(1)

    @ref
    def cancelNewsBulletins(self):
        self.connection.cancelNewsBulletins()

    @ref
    def setServerLogLevel(self):
        self.connection.setServerLogLevel(3)

    @ref
    def reqAutoOpenOrders(self):
        self.connection.reqAutoOpenOrders(1)

    @ref
    def reqAllOpenOrders(self):
        self.connection.reqAllOpenOrders()

    @ref
    def reqManagedAccts(self):
        self.connection.reqManagedAccts()

    @ref
    def requestFA(self):
        self.connection.requestFA(1)

    @ref
    def reqMktData(self):
        contract = Contract() #
        contract.m_symbol = 'AUD'
        contract.m_currency = 'USD'
        contract.m_secType = 'CASH'
        contract.m_exchange = 'IDEALPRO'
        self.connection.reqMktData(1, contract, '', False)

    @ref
    def reqFundamentalData(self, tickerid, contract_details, report_type):
        self.connection.reqFundamentalData(tickerid, contract_details, report_type)

    @ref
    def reqHistoricalData(self):
        contract = Contract()
        contract.m_symbol = 'QQQQ'
        contract.m_secType = 'STK'
        contract.m_exchange = 'SMART'
        endtime = strftime('%Y%m%d %H:%M:%S')
        self.connection.reqHistoricalData(
            tickerId=1,
            contract=contract,
            endDateTime=endtime,
            durationStr='1 D',
            barSizeSetting='1 min',
            whatToShow='TRADES',
            useRTH=0,
            formatDate=1)

    @ref
    def eDisconnect(self):
        sleep(5)
        self.connection.eDisconnect()

    # def create_contract(self, symbol, secType, exchange, currency,
    #                     right = None, strike = None, expiry = None,
    #                     multiplier = None, tradingClass = None,
    #                     localSymbol = None, includeExpired=None):

# if __name__ == '__main__':
#     app = ReferenceApp()
#     app.eConnect()
#     contract = Contract()  #
#     contract.m_symbol = '700'
#     contract.m_currency = 'HKD'
#     contract.m_secType = 'STK'
#     contract.m_exchange = 'SEHK'
#     # f = open('1.xml', 'w')
#     app.reqFundamentalData('8001', contract, "RESC")
#     # f.close(str)
#     time.sleep(10)
#     fundamentaldata = app.wrapper.fundamental_Data_data
#     print(fundamentaldata)
#
#     f = open('/getsymbol/700.xml', 'w+')
#     f.write(fundamentaldata)
#     f.close()
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open('C:\\Users\\Perman\\Documents\\WORK\\IbPy-fundamental\\getsymbol\\Book2.xlsx')
#     wb.XmlImport('C:\\Users\\Perman\\Documents\\WORK\\IbPy-fundamental\\getsymbol\\700.xml')
#     wb.SaveAs("C:\\Users\\Perman\\Documents\\WORK\\IbPy-fundamental\\getsymbol\\Book3.xlsx")
#     wb.Close()
#
#     app.eDisconnect()
    
    # methods = argv[1:]
    #
    # if not methods:
    #     methods = ['eConnect', 'eDisconnect', ]
    # elif methods == ['all']:
    #     methods = allMethods
    # if 'eConnect' not in methods:
    #     methods.insert(0, 'eConnect')
    # if 'eDisconnect' not in methods:
    #     methods.append('eDisconnect')
    #
    # print(('### calling functions:', str.join(', ', methods)))
    # for mname in methods:
    #     call = getattr(app, mname, None)
    #     if call is None:
    #         print(('### warning: no call %s' % (mname, )))
    #     else:
    #         print(('## calling', call.__func__.__name__))
    #         call()
    #         print(('## called', call.__func__.__name__))
