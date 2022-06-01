import sys
from numpy import transpose
import requests
import json
import pandas as pd
import time

sym = 'NIFTY'
exp_date = '02-Jun-2022'

requestTimeout = 10

#for xlwt
'''
import xlwt
from xlwt import Workbook
wb = Workbook()
sheet1 = wb.add_sheet('Sheet3', cell_overwrite_ok=True)
rownum = 0
colnum = 0
'''

#for openpyxl
'''
import openpyxl
my_wb = openpyxl.Workbook()
my_sheet = my_wb.active
rownum = 1
colnum = 1
'''

import xlwings as xw
from xlwings.constants import AutoFillType
import xlsxwriter as wr 
#from xlwings import range as xlrange
file = xw.Book('book4.xlsx')
sh1 = file.sheets("Sheet1")
sh2 = file.sheets("Sheet2")
rownum = 1
colnum = 0

a=1

cookie = '_ga=GA1.2.176340240.1650143938; nsit=wmVwpoMJDGbHg9r_TfmoNPbO; AKA_A2=A; ak_bmsc=9D4C981183F0ABF6EDBCEAA18B5F54A4~000000000000000000000000000000~YAAQDj/LFyyukvSAAQAAwsCWIA8V0op9J8DY9AxBryLkAW18mRLiBESyLEWyHDBpVsE3xZXg+KqqLd50VOntO8EZjI3hJst0o67XjjuxMvMlrEKKrlO6DOCGODe91NEf7iCPTIXOLqHwGP+Fi3TYWcIUqIgqwPF761GKDGuDhsjXCGr4SJXDgaP4fiakqoMkJ/pKRnQnvO29Bb6znk0TzQJX9SZT/zvjgNzxI+rX79tPUJFtPYVeD8EmBLxZglDVTxeoKSFFxEC6NoIneEkaaGxM3wrHp2jfH+yABd51Q34PoFWj2c2uXDHx3nhtQX/xzNq5BJcbm+zoQPguAPbI87+JCm9LsBPXaCI+rXaoooki8iWDzm9U4JZCYobWrKkgAbFcebEXeyaH0rqnG3VvT0RpMJk5X5j/6WmQCzBc8lj6DNTtiOtZg+eovccHJOFWHPyJ4maISg64h0np5zrefUzeMvXLVT1jS/l8mwUQL0gVP8oJRkBYBTwaHkGZDw==; nseappid=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJhcGkubnNlIiwiYXVkIjoiYXBpLm5zZSIsImlhdCI6MTY1NDEwOTE2OCwiZXhwIjoxNjU0MTEyNzY4fQ.1Q2l0nQw04NYmtEGuI6eFHq4ZJkjBn9v708WmvsCKQU; RT="z=1&dm=nseindia.com&si=ca6b0f71-5a98-4dbb-aa81-c35b6aeea104&ss=l3vxvugb&sl=1&tt=2iz&bcn=%2F%2F684d0d44.akstat.io%2F"; bm_sv=AEF532EB2C24D9C71D76F70D6834E08F~YAAQDj/LF4aukvSAAQAAS+WWIA9rBvgZYcwalxRvYv2DpFiLbrAA1gaO94bsGob2+reXQzYrD+h80NRh61+JeSOOxnNRfgRKt9BbIdXYJ9v10guKFFUmzd6rnyEWzL5fZJWSabCinGFIXXWrWvWXUJ/2ytJt9HConQlhGjdV3Xfm1MbcXlGp3Woh7T6ReU5xOXPibBI8ZS1Mxn4q0bHKXdE9QVeapmcBQEKcYzSmcM7M7z6DimUHahS5BOprAhQHjcI=~1' 
    
def timeCountdown(t):
    while t:
        mins, secs = divmod(t, 60)
        timeformat = '{:02d}:{:02d}'.format(mins, secs)
        print("Time Remaining : " + timeformat, end='\r')
        time.sleep(1)
        t -= 1
    print('\n')

def pullOptionChain(symbol):
    try:
        url = 'https://www.nseindia.com/api/option-chain-indices?symbol=' + symbol 
        headers = {
          "accept" : "*/*",
          "accept-encoding" : "gzip, deflate, br",
          "accept-language" : "en-US,en;q=0.9",
          "cookie" : cookie,
          "referer" : "https://www.nseindia.com/option-chain", 
          "user-agent" : "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.41 Safari/537.36"
        }
        response = requests.get(url, headers = headers, timeout = requestTimeout).text
        data = json.loads(response)
        return data
    except Exception as e:
        print(e)
        print("\nUpdate the cookies from NSE website".upper())
        sys.exit(1)

def writeOIDataToExcel(oi, colnum, underlying_price):
    try:
        global rownum
        if rownum == 1:
            colname = wr.utility.xl_col_to_name(colnum)
            cell = colname + str(rownum)
            sh1.autofit('columns')
            #sh1.range(cell).api.autofill(sh1.range(cell), AutoFillType.xlFillDefault)
            sh1.range(cell).color = (10, 220, 10)
            sh1.range(cell).font.bold = True
            sh1.range(cell).font.size = 16
            sh1.range(cell).value=underlying_price
            sh1.range(cell).row_height = 30
            rownum = rownum + 1
            
        if rownum == 2 and colnum == 0:
            for x in oi.values():
                print(str(x['strikeprice']))
                colname = wr.utility.xl_col_to_name(colnum)
                #print(colname)
                #print(colname + str(rownum))
                cell = colname + str(rownum)
                sh1.autofit('columns')
                sh1.range(cell).color = (255, 100, 55)
                sh1.range(cell).font.bold = True
                sh1.range(cell).font.size = 13
                #sh1.range(cell).api.Borders(9).LineStyle = 1
                #sh1.range(cell).api.Borders(9).Weight = 3
                #sh1.range(cell).api.HorizontalAlignment = -4108
                sh1.range(cell).value=str(x['strikeprice'])
                colnum = colnum + 1 
        else: 
            print("Writing oi data to excel")
            colnum = 0
            for x in oi.values():
                #print(str(x['openInterest']))
                colname = wr.utility.xl_col_to_name(colnum)
                cell = colname + str(rownum)
                sh1.autofit('columns')
                sh1.range(cell).font.size = 13
                #print(colname)
                #print(colname + str(rownum)) 
                #sh1.range(cell).api.Borders(9).LineStyle = 1
                #sh1.range(cell).api.Borders(9).Weight = 3
                #sh1.range(cell).api.HorizontalAlignment = -4108
                sh1.range(cell).value=str(x['openInterest'])
                colnum = colnum + 1
    except Exception as e:
        print("Writing to Excel Failed with Exception : " + str(e))
        sys.exit(1)

#def underlyingPriceRoundedToNearestFifty

while True:
    try: 
        data = pullOptionChain(sym)
        if len(data) == 0:
            print("No data pulled from NSE site, check the Expiry date")
            break
        else:
            print("Got Response from NSE")
            #print(json.dumps(data, indent= 4))
            exp_list = data['records']['expiryDates']
            underlying_value = data['records']['data'][0]['PE']['underlyingValue']
            print(underlying_value)
            print(data['records']['strikePrices'])
            ce = {}
            pe = {}
            n = 0
            m = 0
            oi = {}
            x = 0
            """
            for i in data['records']['data'] :
                if i['expiryDate'] == exp_date:
                    try: 
                        ce[n] = i['CE']
                        n = n+1
                    except:
                        pass
                    try:
                        pe[m] = i['PE']
                        m = m+1
                    except:
                        pass
            """

            """
            n = 0 
            for i in data['records']['data']:
                if i['expiryDate'] == exp_date:
                    try:
                        oi[n] = { 
                            "changeinOpenInterest_ce" : i["CE"]["changeinOpenInterest"],
                            "pchangeinOpenInterest_ce" : i["CE"]["pchangeinOpenInterest"],
                            "impliedVolatility" : i["CE"]["impliedVolatility"],
                            "change_ce" : i["CE"]["change"],
                            "lastPrice_ce" : i["CE"]["lastPrice"],
                            "bidprice_ce" : i["CE"]["bidprice"],
                            "askPrice_ce" : i["CE"]["askPrice"],
                            "oi_ce" : i['CE']["openInterest"] / 2,

                            "strikeprice" : i['strikePrice'], 

                            "oi_pe" : i["PE"]["openInterest"] / 2,
                            "bidprice_pe" : i["PE"]["bidprice"],
                            "askPrice_pe" : i["PE"]["askPrice"],
                            "lastPrice_pe" : i["PE"]["lastPrice"],
                            "change_pe" : i["PE"]["change"],
                            "impliedVolatility" : i["PE"]["impliedVolatility"],
                            "changeinOpenInterest_pe" : i["PE"]["changeinOpenInterest"],
                            "pchangeinOpenInterest_pe" : i["PE"]["pchangeinOpenInterest"],
                        }
                        n = n+1
                    except:
                        pass
            """
            n = 0
            for i in data['records']['data']:
                if i['expiryDate'] == exp_date:
                    try:
                        oi[n] = { 
                            "strikeprice" : i['strikePrice'], 
                            "openInterest" : "|".join([str(i['CE']["openInterest"] / 2), str(i['PE']["openInterest"] / 2)])
                        }
                        n = n+1
                    except:
                        pass  
            
            if len(oi) == 0:
                print("OpenInterest data not available, check the Expiry date. \nAvailable Expiry Dates : ")
                print(*exp_list, sep='\n')
                sys.exit(1)
            #print("OPENINTEREST") 
            #print(json.dumps(oi, indent = 4))
            #print(type(oi[0]['openInterest']))

            #xlwings ( option - 1 )
            #TO-DO
            #1.get ATM price  2.format columns    
            writeOIDataToExcel(oi, colnum, underlying_value)
            rownum = rownum + 1  
            
            
            #xlwings ( options - 2)
            '''
            #using dataframe  - working
            df = pd.DataFrame.from_dict(oi).transpose()
            print(df)
            if a == 1 :
                oi_df = df[['strikeprice', 'oi']]
            else:
                oi_df = df[['oi']]

            print(oi_df.transpose())
            data = oi_df.transpose()
            sh1.range("A"+str(a)).value = data

            if a == 1:
                a = a + 3
            else: a = a + 2
            '''
    
        
            '''
            #xlwt - only works with .xls(old format excel)
            # This is working but excel values are not getting refreshed . Have to close and open again to see new writes
            if rownum == 0 and colnum == 0:
                for x in oi.values():
                    sheet1.write(rownum, colnum, x['strikeprice'])
                    colnum = colnum + 1
            else: 
                colnum = 0
                for x in oi.values():
                    sheet1.write(rownum, colnum, str(x['oi']))
                    colnum = colnum + 1

            wb.save('book1.xls')
            rownum = rownum + 1
            '''
            '''
            #openpyxl
            # This is working but excel values are not getting refreshed . Have to close and open again to see new writes
            if rownum == 1 and colnum == 1:
                for x in oi.values():
                    my_sheet.cell(row=rownum, column=colnum).value=x['strikeprice']
                    colnum = colnum + 1
            else: 
                colnum = 1
                for x in oi.values():
                    my_sheet.cell(row=rownum, column=colnum).value=str(x['oi'])
                    colnum = colnum + 1
            my_wb.save("book3.xlsx") 
            rownum = rownum + 1       
            '''        
            #print(oi.values())

            timeCountdown(10)
            #time.sleep(180)
    
    
            '''
            oi_df = pd.DataFrame.from_dict(oi).transpose()

            print(oi_df)

            if a == 1:
                df_oi = oi_df[['strikeprice',  'oi']]
            else:
                df_oi = oi_df[['oi']]

            print(df_oi.transpose())
            data = df_oi.transpose() 
            data.to_excel('book1.xlsx', sheet_name='Sheet1", index=False')   
            '''
    except Exception as e:
        print(e)
        print("End of Script")
        #print("\nUpdate the cookies from NSE website".upper())
        sys.exit(1)