import stocklist as sl

def make_buylist() :
    initList1 = [ 
        { 'code' : 'A015750', 'name' : '성우하이텍',"qty":1,'bp':6840,'tp':7210,'cp':6700},
        { 'code' : 'A106240', 'name' : '파인테크닉스',"qty":1,'bp':9020,'tp':9460,'cp':8800}
        ]
    buylist = sl.StockList(buyorsel='buy')
    buylist.set_buylist(initList1)
    buylist.dump()
    print(buylist.tobuylist)

if   __name__ == '__main__': 
    make_buylist()
