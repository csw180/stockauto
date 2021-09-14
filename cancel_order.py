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
cpMiOrder     = win32com.client.Dispatch("CpTrade.CpTd5339")       # 미결제주문조회
cpCancelOrder = win32com.client.Dispatch("CpTrade.CpTd0314")       # 취소
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

def get_incomplete_order() :
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                

    cpMiOrder.SetInputValue(0, acc)
    cpMiOrder.SetInputValue(1, accFlag[0])
    cpMiOrder.SetInputValue(4, "0") # 전체
    cpMiOrder.SetInputValue(5, "1") # 정렬 기준 - 역순
    cpMiOrder.SetInputValue(6, "0") # 전체
    cpMiOrder.SetInputValue(7, 20) # 요청 개수 - 최대 20개

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

# 취소 주문 - BloockReqeust 를 이용해서 취소 주문
def cancel(ordernum, code, name, amount):
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션       

    cpCancelOrder.SetInputValue(1, ordernum)  # 원주문 번호 - 정정을 하려는 주문 번호
    cpCancelOrder.SetInputValue(2, acc)  # 상품구분 - 주식 상품 중 첫번째
    cpCancelOrder.SetInputValue(3, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpCancelOrder.SetInputValue(4, code)  # 종목코드
    cpCancelOrder.SetInputValue(5, 0)  # 정정 수량, 0 이면 잔량 취소임

    ret = cpCancelOrder.BlockRequest()
    if ret == 4:
        remain_time = cpStatus.LimitRequestRemainTime
        printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
        time.sleep(remain_time/1000) 
        return False

    time.sleep(2)
    dbgout("`cancel("+ str(name) + ' : ' + str(code) + ") -> all cancel submitted!" + "`")
    if cpCancelOrder.GetDibStatus() != 0:
        return False
    return True
 
if __name__ == '__main__': 
    printlog('check_cybos_system() :', check_cybos_system())  # 크레온 접속 점검
    lst_incomplete_orders = get_incomplete_order()
    if lst_incomplete_orders is not None :
        for dct_idx1 in  lst_incomplete_orders :
            cancel(dct_idx1['orderno'],dct_idx1['code'],dct_idx1['name'],dct_idx1['qty'])