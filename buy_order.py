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

def buy(code,name,qty,bp):
    """ input 으로 받은 종목을 매수처리한다"""
    try:
        time_now = datetime.now()
        if bp is None :
            current_price, ask_price, bid_price = get_current_price(code) 
            bidlist, offerlist = get_bidoffer_price(code)
            for i in bidlist :
                if  current_price * 0.99 > i :
                    bp = i
                    break
            printlog(f'({code}){name} buy price assumed currnt price({current_price})*0.99')
     
        buy_qty = qty        # 매수할 수량 초기화
        ret = buy_as_fixedprice(code,buy_qty,bp)
        printlog(f'Buy submitted !! ({code}){name} {bp}({buy_qty}) STATUS={ret}')
        if ret == 4:
            remain_time = cpStatus.LimitRequestRemainTime
            printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
            time.sleep(remain_time/1000) 
            return False
        time.sleep(2)
        dbgout("`Buy("+ str(name) + ' : ' + str(code) + ") -> " + str(buy_qty) + "EA submitted!" + "`")
    except Exception as ex:
        dbgout("`Buy("+ str(code) + ") -> exception! " + str(ex) + "`")

def buy_as_fixedprice(code,buy_qty,buy_price) :
    print(f'buy_as_fixedprice: code={code}, buy_price={buy_price}, buy_qty={buy_qty}')
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                

    cpOrder.SetInputValue(0, "2")        # 2: 매수
    cpOrder.SetInputValue(1, acc)        # 계좌번호
    cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
    cpOrder.SetInputValue(3, code) # 종목코드
    cpOrder.SetInputValue(4, buy_qty)    # 매수할 수량
    cpOrder.SetInputValue(5, buy_price)    # 주문단가
    cpOrder.SetInputValue(7, "0")        # 주문조건 0:기본, 1:IOC, 2:FOK
    cpOrder.SetInputValue(8, "01")        # 주문호가 1:보통, 3:시장가
    # cpOrder.SetInputValue(8, "12")       # 주문호가 1:보통, 3:시장가
                                            # 5:조건부, 12:최유리, 13:최우선 
        # 매수 주문 요청
    ret = cpOrder.BlockRequest() 
    return ret

def get_bidoffer_price(code):
        # 10차 호가 통신
    cpStockjpbid.SetInputValue(0, code)
    cpStockjpbid.BlockRequest()
    print("통신상태", cpStockjpbid.GetDibStatus(), cpStockjpbid.GetDibMsg1())
    bidlist = []
    offerlist = []
    if  cpStockjpbid.GetDibStatus() != 0:
        printlog('10호가 데이터 입수불가 status='+cpStockjpbid.GetDibStatus())
        return bidlist,offerlist
    # 10차호가
    for i in range(10):
        offerlist.append(cpStockjpbid.GetDataValue(0, i))  # 매도호가
        bidlist.append(cpStockjpbid.GetDataValue(1, i) )   # 매수호가
    # for debug
    for i in range(10):
        print(i+1, "차 매도/매수 호가: ", offerlist[i], bidlist[i])
    return bidlist,offerlist

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
    lst_buy  = sl.StockList(buyorsel='buy').tobuylist
    print(lst_buy)
    time_now = datetime.now()
    str_today = time_now.strftime('%Y%m%d')
    for dct_idx1 in lst_buy :
        pass
        buy(dct_idx1['code'],dct_idx1['name'],dct_idx1.get('qty'),dct_idx1.get('bp'))