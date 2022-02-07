#pip install xlwings; install MS-office

import xlwings as xw
from datetime import datetime

class TaxCalPL():
    lastrow=16
    app=None
    wb=None
    sh=None
    mapSize={'bu':10}
    def insertRow(self,row):
        lot=row["lot"]
        for i in range(lot):
            self.sh.api.Rows(self.lastrow).Insert()
            self.sh.range("A"+str(self.lastrow)).value=row["symbol"]
            self.sh.range("B"+str(self.lastrow)).value=row["descr"]
            self.sh.range("C"+str(self.lastrow)).value=row["size"]
            self.sh.range("D"+str(self.lastrow)).value=1
            self.sh.range("E"+str(self.lastrow)).value=row["date"]
            self.sh.range("F"+str(self.lastrow)).value=row["price"]
            self.sh.range("G"+str(self.lastrow)).value=row["fee"]/lot
            self.lastrow+=1

    def initHistory(self):
        rowno=7
        while self.sh.range("A"+str(rowno)).value!=None:
            row={
                "symbol":self.sh.range("A"+str(rowno)).value,
                "descr":self.sh.range("B"+str(rowno)).value,
                "size":self.sh.range("C"+str(rowno)).value,
                "lot":int(self.sh.range("D"+str(rowno)).value),
                "direct":self.sh.range("E"+str(rowno)).value,
                "date":self.sh.range("F"+str(rowno)).value,
                "price":self.sh.range("G"+str(rowno)).value,
                "fee":self.sh.range("H"+str(rowno)).value
            }
            self.insertRow(row)
            rowno+=1

    def getSize(self,symbol):
        preSymbol=''.join([i for i in symbol if not i.isdigit()]) #cut digit from symbol
        if self.mapSize.get(preSymbol)==None:
            print("Can't find size for "+symbol)
            return(None)
        return(self.mapSize.get(preSymbol))
    
    def readTradeFile(self):
        state=0
        rowNo=0
        trades=[]
        closes=[]
        file=open("./trade_sample.txt","r")
        while True:
            line=file.readline()
            if not line:
                break
            
            if(line.find("Transaction Record")!=-1):
                state=1
                rowNo=0
            if(line.find("Position Closed")!=-1):
                state=2
                rowNo=0
            rowNo+=1
            if(state==1): #trade record
                if(rowNo>=6):
                    if line[0:4]=="----":
                        state=0
                    else:
                        fields=line.split("|")
                        row={
                            "date":datetime.strptime(fields[1].strip(), '%Y%m%d'),
                            "descr":fields[3].strip(),
                            "symbol":fields[4].strip(),
                            "direct":fields[5].strip(),
                            "price":float(fields[7].strip()),
                            "lot":int(fields[8].strip()),
                            "oc":fields[10].strip(),
                            "fee":float(fields[11].strip()),
                            "size":self.getSize(fields[4].strip())
                        }
                        trades.append(row)

            if(state==2): #close record
                if(rowNo>=6):
                    if line[0:4]=="----":
                        state=0
                    else:
                        fields=line.split("|")
                        row={
                            "closedate":datetime.strptime(fields[1].strip(), '%Y%m%d'),
                            "descr":fields[3].strip(),
                            "symbol":fields[4].strip(),
                            "opendate":datetime.strptime(fields[5].strip(), '%Y%m%d'),
                            "direct":fields[6].strip(),
                            "lot":int(fields[7].strip()),
                            "openprice":float(fields[8].strip()),
                            "closeprice":float(fields[10].strip())
                        }
                        closes.append(row)
        file.close()
        ret={"trades":trades, "closes":closes}
        return(ret)
            
    def closeTrade(self,trade,closeRow,closeLot):
        for lotNo in range(closeLot):
            #search in sheet
            rowNo=16
            closeNo=-1
            while(self.sh.range("A"+str(rowNo)).value!=None):
                symbol=self.sh.range("A"+str(rowNo)).value
                opendate=self.sh.range("E"+str(rowNo)).value
                openprice=self.sh.range("F"+str(rowNo)).value
                closedate=self.sh.range("H"+str(rowNo)).value
                if(closedate==None and closeRow['symbol']==symbol and closeRow['opendate']==opendate and closeRow['openprice']==openprice):
                    closeNo=rowNo
                    break
                rowNo+=1
            if(closeNo==-1):
                print("Can't find close row in excel sheet")
                return(-1)
            closeRowNo=str(closeNo)
            self.sh.range("H"+closeRowNo).value = trade['date']
            self.sh.range("I"+closeRowNo).value = trade['price']
            self.sh.range("J"+closeRowNo).value = trade['fee']/trade['lot']
            self.sh.range("K"+closeRowNo).formula = "=(I"+closeRowNo+"-F"+closeRowNo+")*C"+closeRowNo+"*D"+closeRowNo+"-G"+closeRowNo+"-J"+closeRowNo
        return(0)

    def processTrade(self,ret):
        trades=ret["trades"]
        closes=ret["closes"]
        for i in range(len(trades)):
            trade=trades[i]
            if trade["oc"]=="O":
                self.insertRow(trade)
            else:
                remainLot=trade['lot']
                while remainLot>0:
                    #find closeRow first
                    closeRow=None
                    for j in range(len(closes)):
                        closeRow=closes[j]
                        if(trade['symbol']==closeRow['symbol'] and trade['date']==closeRow['closedate'] and trade['price']==closeRow['closeprice'] and closeRow['lot']>0):
                            break
                    if(closeRow==None):
                        print("Can't find closeRow")
                        return(-1)
                    else:
                        closeLot=closeRow['lot']
                        if(remainLot>=closeRow['lot']):
                            remainLot-=closeRow['lot']
                            closeRow['lot']=0
                        else:
                            closeRow['lot']-=remainLot
                            closeLot=remainLot
                            remainLot=0
                        flag=self.closeTrade(trade,closeRow,closeLot)
                        if(flag==-1):
                            return(-1)
        return(0)

    def cal(self):
        self.app = xw.App(visible=False)
        self.wb = self.app.books.open("./Account_Sample.xls")
        self.sh = self.wb.sheets["sheet1"]
        self.initHistory()
        ret=self.readTradeFile()
        flag=self.processTrade(ret)
        print("flag="+str(flag))
        if(self.lastrow>16):
            self.sh.range("K"+str(self.lastrow+1)).formula = "=SUM(K16:K"+str(self.lastrow-1)+")"
        self.wb.save()
        self.wb.close()
        self.app.quit()


taxCalPL=TaxCalPL()
taxCalPL.cal()