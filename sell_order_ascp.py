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

def sell(code,name,sel_qty):
    try:
        # time_now = datetime.now()
        ret = sel_as_marketprice(code,sel_qty)
        printlog(f'Sell submitted !! ({code}){name} ({sel_qty}) STATUS={ret}')
        if ret == 4:
            remain_time = cpStatus.LimitRequestRemainTime
            printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
            time.sleep(remain_time/1000) 
            return False
        time.sleep(30)
        dbgout("`Sell("+ str(name) + ' : ' + str(code) + ") -> " + str(sel_qty) + "EA submitted!" + "`")
    except Exception as ex:
        dbgout("`Sell("+ str(code) + ") -> exception! " + str(ex) + "`")

def sel_as_marketprice(code,sel_qty) :
    print(f'sel_as_marketprice: code={code}, sel_qty={sel_qty}')
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                

    cpOrder.SetInputValue(0, "1")        # 1: 매도
    cpOrder.SetInputValue(1, acc)        # 계좌번호
    cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
    cpOrder.SetInputValue(3, code) # 종목코드
    cpOrder.SetInputValue(4, sel_qty)    # 매수할 수량
    # cpOrder.SetInputValue(5, sel_price)  # 주문단가
    cpOrder.SetInputValue(7, "0")        # 주문조건 0:기본, 1:IOC, 2:FOK
    # cpOrder.SetInputValue(8, "01")       # 주문호가 1:보통, 3:시장가
    cpOrder.SetInputValue(8, "03")       # 주문호가 1:보통, 3:시장가
                                         # 5:조건부, 12:최유리, 13:최우선 
        # 매수 주문 요청
    ret = cpOrder.BlockRequest() 
    return ret

def get_current_price(code):
    """인자로 받은 종목의 현재가, 매수호가, 매도호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # 현재가
    item['ask'] =  cpStock.GetHeaderValue(16)        # 매수호가
    item['bid'] =  cpStock.GetHeaderValue(17)        # 매도호가    
    return item['cur_price'], item['ask'], item['bid']

if __name__ == '__main__':
    printlog('check_cybos_system() :', check_cybos_system())  # 크레온 접속 점검
    list_sel = sl.StockList(buyorsel='sel').tosellist
    # print(list_sel)
    # time_now = datetime.now()
    # str_today = time_now.strftime('%Y%m%d')
    for d1 in list_sel :
        current_price, ask_price, bid_price = get_current_price(d1['code']) 
        if  d1['cp'] * 0.99 > current_price  :
            print(f"sell as market price!!! code={d1['code']}, name={d1['name']}sel_qty={d1['qty']}")
            sell(d1['code'],d1['name'],d1['qty'])
