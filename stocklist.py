import json

class StockList :
    def __init__(self,buyorsel) :
        self.buyorsel = buyorsel
        if self.buyorsel == 'buy' :
            with open('tobuylist.json','r',encoding='utf-8') as json_file_r:
                self.tobuylist = json.load(json_file_r)
        else :
            with open('tosellist.json','r',encoding='utf-8') as json_file_r:
                self.tosellist = json.load(json_file_r)

    def set_buylist(self,list) :
        self.tobuylist = list

    def set_sellist(self,list) :
        self.tosellist = list

    def dump(self) :
        if self.buyorsel == 'buy' :
            with open('tobuylist.json','w',encoding='utf-8') as json_file_w :
                json.dump(self.tobuylist, json_file_w, ensure_ascii=False)
        else :
            with open('tosellist.json','w',encoding='utf-8') as json_file_w :
                json.dump(self.tosellist, json_file_w, ensure_ascii=False)

    def hascode(self, code) :
        rt = None
        if self.buyorsel == 'buy' :
            for d1 in self.tobuylist :
                if d1['code'] == code :
                    rt =d1
        else :
            for d1 in self.tosellist :
                if d1['code'] == code :
                    rt =d1
        return rt

    # def setclosed(self,code) :
    #     for l in self.tobuylist :
    #         if l['code'] == code :
    #             l['closed'] = 'Y'

    # def get_notclosed(self) :
    #     temp =[]
    #     for l in self.tobuylist :
    #         if l['closed'] == 'N' : temp.append(l)
    #     return temp

    # def delete(self,code) :
    #     for l in self.tobuylist :
    #         if l['code'] == code :
    #             self.tobuylist.remove(l)