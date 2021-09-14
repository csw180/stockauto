# pip install pywin32
import win32com.client
import pandas as pd
import ctypes
from datetime import datetime
import requests
import time
import stocklist as sl

# 크레온 플러스 공통 OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpStockjpbid = win32com.client.Dispatch('DsCbo1.StockJpBid2')
cpMiOrder     = win32com.client.Dispatch("CpTrade.CpTd5339")       # 미결제주문조회
cpCancelOrder = win32com.client.Dispatch("CpTrade.CpTd0314")       # 취소
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  

def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    myToken = 'xoxb-2218612794309-2221859854627-0A7IXYYuBRwhFDOrKMzaOtv4'
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    post_message(myToken,"#stock",datetime.now().strftime('[%m/%d %H:%M:%S] ') +message)

def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text}
    )
    printlog(f'SLACK msg send respose.status_code={response.status_code}')

def printlog(message, *args):
    """인자로 받은 문자열을 파이썬 셸에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)

def check_cybos_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_cybos_system() : admin user -> FAILED')
        return False
 
    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        printlog('check_cybos_system() : connect to server -> FAILED')
        return False
 
    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        printlog('check_cybos_system() : init trade -> FAILED')
        return False
    return True

def get_stock_balance():
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)         # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)          # 요청 건수(최대 50)
    ret = cpBalance.BlockRequest()     
    # if code == 'ALL':
    #     dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
    #     dbgout('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
    #     dbgout('평가금액: ' + str(cpBalance.GetHeaderValue(3)))
    #     dbgout('평가손익: ' + str(cpBalance.GetHeaderValue(4)))
    #     dbgout('종목수: ' + str(cpBalance.GetHeaderValue(7)))
    if ret == 4:
        remain_time = cpStatus.LimitRequestRemainTime
        printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
        time.sleep(remain_time/1000) 
        return False

    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)   # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        stocks.append({'code': stock_code, 'name': stock_name,'qty': stock_qty})
    return stocks

def make_sellist() :
    lst_old_sell = sl.StockList(buyorsel='sel')
    lst_buy = sl.StockList(buyorsel='buy')
    lst_stocks = get_stock_balance() 
    
    lst_new_stocks = []
    if  str(type(lst_stocks)) == "<class 'list'>" :
        for dct_idx1  in lst_stocks :
            dct_old_sell  = lst_old_sell.hascode(dct_idx1['code'])
            dct_buy  = lst_buy.hascode(dct_idx1['code'])
            dct_new = {}
            if  dct_old_sell is not None :
                dct_new = dct_old_sell.copy()
                lst_new_stocks.append(dct_new)
            elif dct_buy is not None :
                dct_new = dct_buy.copy()
                lst_new_stocks.append(dct_new)
            else :
                dct_new['code'] =  dct_idx1['code']
                dct_new['name'] =  dct_idx1['name']
                dct_new['qty']  =  dct_idx1['qty']
                lst_new_stocks.append(dct_new)
        print(f'Sync Sell List:  TOT Count={len(lst_new_stocks)}')
        for dct_idx3 in lst_new_stocks :
            print(f'Sync Sell List: {dct_idx3}')
        stocklist_sell =  sl.StockList(buyorsel='sel')
        stocklist_sell.set_sellist(lst_new_stocks)
        stocklist_sell.dump()

if __name__ == '__main__':
    printlog('check_cybos_system() :', check_cybos_system())  # 크레온 접속 점검
    printlog('Sync Sell List')                               # 주식보유잔고에 맞게 sell list를 sync 맞춘다.
    make_sellist()