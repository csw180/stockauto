# pip install pywin32
import win32com.client
import pandas as pd
import ctypes
from datetime import datetime
import requests
import time
import stocklist as sl
import sync_selllist

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

def sell(code,name,sel_qty,tp):
    try:
        # time_now = datetime.now()
        ret = sel_as_fixedprice(code,sel_qty,tp)
        printlog(f'Sell submitted !! ({code}){name} ({sel_qty}) ({tp}) STATUS={ret}')
        if ret == 4:
            remain_time = cpStatus.LimitRequestRemainTime
            printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
            time.sleep(remain_time/1000) 
            return False
        time.sleep(30)
        dbgout("`Sell("+ str(name) + ' : ' + str(code) + ") -> " + str(sel_qty) + str(tp) + "EA submitted!" + "`")
    except Exception as ex:
        dbgout("`Sell("+ str(code) + ") -> exception! " + str(ex) + "`")

def sel_as_fixedprice(code,sel_qty,tp) :
    print(f'sel_as_marketprice: code={code}, sel_qty={sel_qty}, price-{tp}')
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                

    cpOrder.SetInputValue(0, "1")        # 1: 매도
    cpOrder.SetInputValue(1, acc)        # 계좌번호
    cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
    cpOrder.SetInputValue(3, code) # 종목코드
    cpOrder.SetInputValue(4, sel_qty)    # 매수/매도할 수량
    cpOrder.SetInputValue(5, tp)         # 주문단가
    cpOrder.SetInputValue(7, "0")        # 주문조건 0:기본, 1:IOC, 2:FOK
    cpOrder.SetInputValue(8, "01")       # 주문호가 1:보통, 3:시장가
    # cpOrder.SetInputValue(8, "12")       # 주문호가 1:보통, 3:시장가
    #                                      # 5:조건부, 12:최유리, 13:최우선 
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


def get_yet_sell() :
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                

    cpMiOrder.SetInputValue(0, acc)
    cpMiOrder.SetInputValue(1, accFlag[0])
    cpMiOrder.SetInputValue(4, "0") # 전체
    cpMiOrder.SetInputValue(5, "1") # 정렬 기준 - 역순
    cpMiOrder.SetInputValue(6, "0") # 전체
    cpMiOrder.SetInputValue(7, 20) # 요청 개수 - 최대 20개
    cpMiOrder.SetInputValue(13, "1") # 1:매도, 2:매수
    ret = cpMiOrder.BlockRequest() 
    if ret == 4:
        remain_time = cpStatus.LimitRequestRemainTime
        printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
        time.sleep(remain_time/1000) 
        return False
    
    # 수신 개수
    cnt = cpMiOrder.GetHeaderValue(5)
    orders = None
    if cnt != 0 :
        orders = []
        for i in range(cnt):
            dict_order = {}
            dict_order['orderno']= cpMiOrder.GetDataValue(1, i)  # 주문번호
            dict_order['code']  = cpMiOrder.GetDataValue(3, i)  # 종목코드
            dict_order['name']  = cpMiOrder.GetDataValue(4, i)  # 종목명
            dict_order['qty']   = cpMiOrder.GetDataValue(11, i)  # 정정취소 가능수량
            orders.append(dict_order)

            # objRq.GetDataValue(1, i)
            # objRq.GetDataValue(2, i)
            # objRq.GetDataValue(3, i)  # 종목코드
            # objRq.GetDataValue(4, i)  # 종목명
            # objRq.GetDataValue(5, i)  # 주문구분내용
            # objRq.GetDataValue(6, i)  # 주문수량
            # objRq.GetDataValue(7, i)  # 주문단가
            # objRq.GetDataValue(8, i)  # 체결수량
            # objRq.GetDataValue(9, i)  # 신용구분
            # objRq.GetDataValue(11, i)  # 정정취소 가능수량
            # objRq.GetDataValue(13, i)  # 매매구분코드
            # objRq.GetDataValue(17, i)  # 대출일
            # objRq.GetDataValue(19, i)  # 주문호가구분코드내용
            # objRq.GetDataValue(21, i)  # 주문호가구분코드
    return orders

if  __name__ == '__main__':
    printlog('check_cybos_system() :', check_cybos_system())  # 크레온 접속 점검
    lst_stock_balance = get_stock_balance() 
    # print('lst_stock_balance=',lst_stock_balance)

    lst_sel = sl.StockList(buyorsel='sel').tosellist
    lst_buy = sl.StockList(buyorsel='buy').tobuylist

    lst_need_order = []

    if  str(type(lst_stock_balance)) == "<class 'list'>" :
        for dct_idx1 in lst_stock_balance :
            if  dct_idx1['qty'] > 0 :     # 매도주문 안나간 물량만큼 qty 로 남아있다.
                dct_need_order = dct_idx1.copy()
                for  dct_idx3 in lst_sel :                   # 어제까지 보유하던 물량은 sellist 에 목표가, 손절가 정보가 들어 있다.
                    if  dct_need_order['code'] == dct_idx3['code']  :
                        dct_need_order['bp'] = dct_idx3['bp']
                        dct_need_order['tp'] = dct_idx3['tp']
                        dct_need_order['cp'] = dct_idx3['cp']
                        break

                for  dct_idx4 in lst_buy :                    # 오늘 사서 보유하게된 물량은 buylist 에 목표가, 손절가 정보가 들어 있다.
                    if  dct_need_order['code'] == dct_idx4['code']  :
                        dct_need_order['bp'] = dct_idx4['bp']
                        dct_need_order['tp'] = dct_idx4['tp']
                        dct_need_order['cp'] = dct_idx4['cp']
                        break

                lst_need_order.append(dct_need_order)      
        print('lst_need_order=',lst_need_order)

    for dct_idx4 in  lst_need_order :
        pass
        print(dct_idx4) 
        sell(dct_idx4['code'],dct_idx4['name'],dct_idx4['qty'],dct_idx4['tp'])