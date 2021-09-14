import json
import stocklist   as sl

if __name__ == '__main__': 
    initList = [ 
        { 'code' : 'A001570', 'name' : '금양','bp':5770,'tp':6180,'cp':5310} ]


    initList1 = [ 
        { 'date' : '20210707','code' : 'A017000', 'name' : '신원종합개발','bp':7940} ,
        { 'date' : '20210707','code' : 'A036540', 'name' : 'SFA반도체','bp':7630}  ,
        { 'date' : '20210707','code' : 'A052690', 'name' : '한전기술','bp':55000}  ,
        { 'date' : '20210707','code' : 'A103590', 'name' : '일진전기','bp':5750}  ,
        { 'date' : '20210707','code' : 'A105550', 'name' : '트루윈','bp':7720}  ,
        { 'date' : '20210707','code' : 'A052600', 'name' : '한네트','bp':10100 }]

    initList2 = [ 
        {"code": "A017000", "name": "신원종합개발", "qty": 1,"tp": 8300, "cp": 7580},
        {"code": "A036540", "name": "SFA반도체", "qty": 1,"tp": 7990, "cp": 7430},
        {"code": "A052690", "name": "한전기술", "qty": 1,"tp": 58400, "cp": 51200},
        {"code": "A103590", "name": "일진전기", "qty": 1,"tp": 6060, "cp": 5620},
        {"code": "A105550", "name": "트루윈", "qty": 1,"tp": 8190, "cp": 7500}]

    buylist = sl.StockList(buyorsel='buy')
    buylist.set_buylist(initList)
    buylist.dump()
    print(buylist.tobuylist)

    # sellist = sl.StockList(buyorsel='sel')
    # sellist.set_sellist(initList2)
    # sellist.dump()
    # print(sellist.tosellist)


