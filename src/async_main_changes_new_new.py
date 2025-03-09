import time
import os
import asyncio
import pandas as pd
from datetime import datetime,timedelta
from ib_broker import *
import credentials
from openpyxl import load_workbook
import nest_asyncio

nest_asyncio.apply()

def resource_path(relative_path):
    """
    Ensures the path works both in development and after PyInstaller bundling.
    """
    base_path = os.path.abspath(".")  # Change as needed
    return os.path.join(base_path, relative_path)

class IBRKExcel:
    def __init__(self):
        self.symbol          = 'N225M'
        self.exchange        = 'OSE.JPN'
        self.path            = credentials.xlsx_path
        self.current_time    = datetime.now()
        # self.excel_data      = pd.read_excel(self.path, sheet_name=credentials.sheet_name) 
        self.excel_data      = pd.read_excel(resource_path(r"storage\new_order_sheet.xlsx"),sheet_name=credentials.sheet_name)
        self.length          = len(self.excel_data)
        self.orderbook       = []
        self.failed_orders   = []
        # self.database_path   = credentials.database_path
        self.upper_trigger   = 100000000
        self.lower_trigger   = -100000000
        self.current_time    = datetime.now().time()

    async def check_excel_changes(self):
        # new_data   = pd.read_excel(self.path, sheet_name=credentials.sheet_name)
        new_data   = pd.read_excel(resource_path(r"storage\new_order_sheet.xlsx"),sheet_name=credentials.sheet_name)
        new_length = len(new_data)

        if new_length  != self.length:
            self.length     = new_length
            self.excel_data = new_data
            return True
        return False

    async def connection_show(self) -> bool:
        host, port = credentials.host, credentials.port
        self.client = IB()
        self.ib = self.client
        connection_print = self.client.connect(host=host,port=port,clientId=13,account='DU9727656',timeout=60)
        print(connection_print)

    async def check_for_new_positions(self): # put this in async
        if await self.check_excel_changes():
            print("a change on the excel has been made")
            # length   = len(pd.read_excel(self.path, sheet_name=credentials.sheet_name))
            length   = len(pd.read_excel(resource_path(r"storage\new_order_sheet.xlsx"),sheet_name=credentials.sheet_name))
            for i in range(length):
                if self.excel_data.loc[i,'Activation'] == 1: # a new order detected
                    row           = self.excel_data.iloc[i]
                    self.symbol        = 'N225M'
                    self.exchange      = 'OSE.JPN' 
                    self.trigger_level = row['Trigger_Level_High_Low']
                    self.entry_type    = row['Entry_Type']
                    self.entry_strike  = row['Entry_Strike']
                    self.strike_type   = row['Strike_Type']
                    self.expiry        = str(row['Expiry'])
                    self.target        = row['Target']
                    self.stop_loss     = row['Stop_Loss']
                    self.qty           = row['Qty']
                    self.slicing       = row['Slicing']
                    self.time_interval = row['Time_Interval']
                    self.activation    = row['Activation']
                    # if self.strike_type == 'CE': # CHANGED THIS LINE
                    if self.strike_type == "SELL":
                        self.side = 'SELL'
                    else:
                        self.side = 'BUY' 

                    datevar = self.expiry
                    # date,timep = datevar.split(' ') # changed
                    # date = datevar
                    # print(f"date is on line 73 {date}")
                    # print(date)
                    # year,day,month = date.split('-')
                    datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                    year,day,month = datevar.split('-')
                    print(f"year is {year}")
                    print(f"month is {month}")
                    print(f"day is {day}")
                    # day,month,year = date.split('-')
                    # formatted_date = f"{year}{month.zfill(2)}" # changed
                    formatted_date = f"{year}{month.zfill(2)}{day}"
                    
                    print(f"formatted_date is {formatted_date}")
                    # exit(0)
                    contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                    print(self.trigger_level)
                    print(await self.get_current_market_price_futures(contract)) 
                    print(self.entry_type)
                    print(self.strike_type)
                    # if self.strike_type == "PE" and self.trigger_level < await self.get_current_market_price_futures(contract):
                    if self.strike_type == "BUY" and self.trigger_level <= await self.get_current_market_price_futures(contract): # current_price breaks through trigger_level
                        if self.entry_type == "LIMIT":
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                # date,timep = datevar.split(' ')
                                # date = datevar 
                                # print(f"date is {date}")
                                # year,day,month = date.split('-')
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar.split('-')
                                # formatted_date = f"{year}{month.zfill(2)}"
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                # exit(0)
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                attempt = 0
                                # while attempt<3:
                                # successful_orders = 0
                                # failed_orders = 0
                                while attempt<int(credentials.attempts):
                                    bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                    if credentials.trade_type_default == 0:
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(self.entry_strike)) 
                                    else:
                                        print(f"using trade_type default {credentials.trade_type_default}")
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(int((bid + (2**attempt - 1)*ask)/2**attempt))) 
                                    self.order.account = 'DU9727656'
                                    self.order.transmit = True
                                    print(f"Placing limit order,attempt {attempt+1}")
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print(self.order_details)
                                    # await asyncio.sleep(3) # keep this part same
                                    await asyncio.sleep(credentials.pause_time)
                                    print(self.order_details.isDone())

                                    if not self.order_details.isDone():
                                        print("The cancelled order is :\n")
                                        self.canceled_order_details = self.client.cancelOrder(order=self.order_details.orderStatus)
                                        print(self.canceled_order_details)
                                        print("Order failed")
                                        # failed_orders = failed_orders+1
                                    else:
                                        print("Limit order placed successfully")
                                        print(self.order_details)
                                        # successful_orders = successful_orders+1
                                        break
                                    
                                    attempt = attempt+1
                                
                                # if attempt == 3:
                                if attempt == credentials.attempts:
                                    print(f"Limit order failed {credentials.attempts} times placing market order")
                                    # self.contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                                    self.order = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing)))
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print("Market order placed")
                                    print(self.order_details)

                            self.excel_data.loc[i,'Activation'] = -1 
                            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)

                        else:
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                # date,timep = datevar.split(' ')
                                # date = datevar
                                # year,day,month = date.split('-')
                                # print(f"date is on line 161 {date}")
                                # print(date)
                                # day,month,year = date.split('-')
                                # year,day,month = date.split('-')
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar.split('-')
                                # formatted_date = f"{year}{month.zfill(2)}"
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                # exit(0)
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                # self.contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                                self.order          = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) 
                                self.order.account = 'DU9727656'
                                self.order.transmit = True
                                self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                # await asyncio.sleep(3)
                                if credentials.user_time_default == 0:
                                    await asyncio.sleep(self.time_interval)
                                else:
                                    print(f"sleep for {credentials.user_time}")
                                    await asyncio.sleep(credentials.user_time)
                                print("The order has been placed")
                            self.excel_data.loc[i,'Activation'] = -1 
                            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)

                    elif self.strike_type == "SELL" and self.trigger_level >= await self.get_current_market_price_futures(contract):
                        if self.entry_type == "LIMIT":
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                # date,timep = datevar.split(' ')
                                # date = datevar
                                # year,day,month = date.split('-')
                                # print(f"date is on line 197 {date}")
                                # print(date)
                                # day,month,year = date.split('-')
                                # year,day,month = date.split('-')
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar.split('-')
                                # formatted_date = f"{year}{month.zfill(2)}"
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                # exit(0)
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                attempt = 0
                                while attempt<int(credentials.attempts):
                                    bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                    if credentials.trade_type_default == 0:
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(self.entry_strike))
                                    else:
                                        print(f"using trade_type default {credentials.trade_type_default}")
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(int((bid + (2**attempt - 1)*ask)/2**attempt)))  
                                    self.order.account = 'DU9727656'
                                    self.order.transmit = True
                                    print(f"Placing limit order,attempt {attempt+1}")
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print(self.order_details)
                                    # await asyncio.sleep(3) # keep this same
                                    await asyncio.sleep(credentials.pause_time)
                                    print(self.order_details.isDone())

                                    if not self.order_details.isDone():
                                        print("The cancelled order is :\n")
                                        self.canceled_order_details = self.client.cancelOrder(order=self.order_details.orderStatus)
                                        print(self.canceled_order_details)
                                        print("Order failed")
                                    else:
                                        print("Limit order placed successfully")
                                        print(self.order_details)
                                        break
                                    attempt = attempt+1
                                
                                # if attempt == 3:
                                if attempt == credentials.attempts:
                                    print(f"Limit order failed {credentials.attempts} times placing market order")
                                    # self.contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                                    self.order = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing)))
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print("Market order placed")
                                    print(self.order_details)

                            self.excel_data.loc[i,'Activation'] = -1 # now use excelwriter fn
                            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)

                        else:
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                # date,timep = datevar.split(' ')
                                # date = datevar
                                # print(date)
                                # print(f"date is on line 256 {date}")
                                # year,day,month = date.split('-')
                                # day,month,year = date.split('-')
                                # year,day,month = date.split('-')
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar.split('-')
                                # formatted_date = f"{year}{month.zfill(2)}"
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                # exit(0)
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                # self.contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                                self.order          = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) 
                                self.order.account = 'DU9727656'
                                self.order.transmit = True
                                self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                await asyncio.sleep(3) # keep this same
                                print(self.order_details)
                                print("The order has been placed")
                            self.excel_data.loc[i,'Activation'] = -1 # now use excelwriter fn
                            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)
                    else:
                        print("The trigger price has not being triggered")

                elif self.excel_data.loc[i,'Activation'] == 3: # a new order detected
                    row           = self.excel_data.iloc[i]
                    self.symbol        = 'N225M'
                    self.exchange      = 'OSE.JPN' 
                    self.trigger_level = row['Trigger_Level_High_Low']
                    self.entry_type    = row['Entry_Type']
                    self.entry_strike  = row['Entry_Strike']
                    self.strike_type   = row['Strike_Type']
                    self.expiry        = str(row['Expiry'])
                    self.target        = row['Target']
                    self.stop_loss     = row['Stop_Loss']
                    self.qty           = row['Qty']
                    self.slicing       = row['Slicing']
                    self.time_interval = row['Time_Interval']
                    self.activation    = row['Activation']
                    # if self.strike_type == 'CE': # CHANGED THIS LINE
                    if self.strike_type == "SELL":
                        self.side = 'SELL'
                    else:
                        self.side = 'BUY' 

                    datevar = self.expiry
                    # date,timep = datevar.split(' ') # changed
                    # date = datevar
                    # print(date)
                    # print(f"date is on line 308 {date}")
                    # year,day,month = date.split('-')
                    datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                    year,day,month = datevar.split('-')
                    print(f"year is {year}")
                    print(f"month is {month}")
                    print(f"day is {day}")
                    # day,month,year = date.split('-')
                    # formatted_date = f"{year}{month.zfill(2)}" # changed
                    formatted_date = f"{year}{month.zfill(2)}{day}"
                    
                    print(f"formatted_date is {formatted_date}")
                    # exit(0)
                    contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                    print(self.trigger_level)
                    print(await self.get_current_market_price_futures(contract)) 
                    print(self.entry_type)
                    print(self.strike_type)
                    # if self.strike_type == "PE" and self.trigger_level < await self.get_current_market_price_futures(contract):
                    if self.strike_type == "BUY" and self.trigger_level <= await self.get_current_market_price_futures(contract): # current_price breaks through trigger_level
                        if self.entry_type == "LIMIT":
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                # date,timep = datevar.split(' ')
                                # date = datevar 
                                # print(f"date is {date}")
                                # year,day,month = date.split('-')
                                # formatted_date = f"{year}{month.zfill(2)}"
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,month,day = datevar.split('-')
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                # exit(0)
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                attempt = 0
                                # while attempt<3:
                                # successful_orders = 0
                                # failed_orders = 0
                                while attempt<int(credentials.attempts):
                                    bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                    if credentials.trade_type_default == 0:
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(self.entry_strike),tif='GTD') 
                                        expiry_date = (datetime.now() + timedelta(days=1)).strftime("%Y%m%d 23:59:59")
                                        self.order.goodTillDate = expiry_date
                                        self.order.tif = "GTD"
                                        
                                    else:
                                        print(f"using trade_type default {credentials.trade_type_default}")
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(int((bid + (2**attempt - 1)*ask)/2**attempt)),tif="GTD") 
                                        expiry_date = (datetime.now() + timedelta(days=1)).strftime("%Y%m%d 23:59:59")
                                        self.order.goodTillDate = expiry_date
                                        self.order.tif = "GTD"

                                    self.order.account = 'DU9727656'
                                    self.order.transmit = True
                                    print(f"Placing limit order,attempt {attempt+1}")
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print(self.order_details)
                                    # await asyncio.sleep(3) # keep this part same
                                    await asyncio.sleep(credentials.pause_time)
                                    print(self.order_details.isDone())

                                    if not self.order_details.isDone():
                                        print("The cancelled order is :\n")
                                        self.canceled_order_details = self.client.cancelOrder(order=self.order_details.orderStatus)
                                        print(self.canceled_order_details)
                                        print("Order failed")
                                        # failed_orders = failed_orders+1
                                    else:
                                        print("Limit order placed successfully")
                                        print(self.order_details)
                                        # successful_orders = successful_orders+1
                                        break
                                    
                                    attempt = attempt+1
                                
                                # if attempt == 3:
                                if attempt == credentials.attempts:
                                    print(f"Limit order failed {credentials.attempts} times placing market order")
                                    # self.contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                                    self.order = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing)),tif="GTD")
                                    expiry_date = (datetime.now() + timedelta(days=1)).strftime('%Y%m%d 23:59:59')
                                    self.order.goodTillDate = expiry_date
                                    self.order.tif = 'GTD'
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print("Market order placed")
                                    print(self.order_details)

                            self.excel_data.loc[i,'Activation'] = -3 
                            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)

                        else:
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                # date,timep = datevar.split(' ')
                                # date = datevar
                                # year,day,month = date.split('-')
                                # print(date)
                                # print(f"date is on line 402 {date}")
                                # day,month,year = date.split('-')
                                # year,day,month = date.split('-')
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar.split('-')
                                # formatted_date = f"{year}{month.zfill(2)}"
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                # exit(0)
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                # self.contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                                self.order          = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) 
                                expiry_date = (datetime.now() + timedelta(days=1)).strftime('%Y%m%d 23:59:59')
                                self.order.goodTillDate = expiry_date
                                self.order.tif = "GTD"
                                self.order.account = 'DU9727656'
                                self.order.transmit = True
                                self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                # await asyncio.sleep(3)
                                if credentials.user_time_default == 0:
                                    await asyncio.sleep(self.time_interval)
                                else:
                                    print(f"sleep for {credentials.user_time}")
                                    await asyncio.sleep(credentials.user_time)
                                print("The order has been placed")
                            self.excel_data.loc[i,'Activation'] = -3
                            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)

                    elif self.strike_type == "SELL" and self.trigger_level >= await self.get_current_market_price_futures(contract):
                        if self.entry_type == "LIMIT":
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                # date,timep = datevar.split(' ')
                                # date = datevar
                                # year,day,month = date.split('-')
                                # print(date)
                                # print(f"date is on line 440 {date}")
                                # day,month,year = date.split('-')
                                # year,day,month = date.split('-')
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar.split('-')

                                # formatted_date = f"{year}{month.zfill(2)}"
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                # exit(0)
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                attempt = 0
                                while attempt<int(credentials.attempts):
                                    bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                    if credentials.trade_type_default == 0:
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(self.entry_strike))
                                        expiry_date = (datetime.now() + timedelta(days=1)).strftime("%Y%m%d 23:59:59")
                                        self.order.goodTillDate = expiry_date
                                        self.order.tif = "GTD"
                                    else:
                                        print(f"using trade_type default {credentials.trade_type_default}")
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(int((bid + (2**attempt - 1)*ask)/2**attempt)))  
                                        expiry_date = (datetime.now() + timedelta(days=1)).strftime("%Y%m%d 23:59:59")
                                        self.order.goodTillDate = expiry_date
                                        self.order.tif = "GTD"

                                    self.order.account = 'DU9727656'
                                    self.order.transmit = True
                                    print(f"Placing limit order,attempt {attempt+1}")
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print(self.order_details)
                                    # await asyncio.sleep(3) # keep this same
                                    await asyncio.sleep(credentials.pause_time)
                                    print(self.order_details.isDone())

                                    if not self.order_details.isDone():
                                        print("The cancelled order is :\n")
                                        self.canceled_order_details = self.client.cancelOrder(order=self.order_details.orderStatus)
                                        print(self.canceled_order_details)
                                        print("Order failed")
                                    else:
                                        print("Limit order placed successfully")
                                        print(self.order_details)
                                        break
                                    attempt = attempt+1
                                
                                # if attempt == 3:
                                if attempt == credentials.attempts:
                                    print(f"Limit order failed {credentials.attempts} times placing market order")
                                    # self.contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                                    self.order = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing)))
                                    expiry_date = (datetime.now() + timedelta(days=1)).strftime('%Y%m%d 23:59:59')
                                    self.order.goodTillDate = expiry_date
                                    self.order.tif = "GTD"
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print("Market order placed")
                                    print(self.order_details)

                            self.excel_data.loc[i,'Activation'] = -3 # now use excelwriter fn
                            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)

                        else:
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                # date,timep = datevar.split(' ')
                                # date = datevar
                                # print(date)
                                # print(f"date is on line 504 {date}")
                                # year,day,month = date.split('-')
                                # day,month,year = date.split('-')
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar.split('-')

                                # formatted_date = f"{year}{month.zfill(2)}"
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                # exit(0)
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                # self.contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                                self.order          = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) 
                                expiry_date = (datetime.now() + timedelta(days=1)).strftime('%Y%m%d 23:59:59')
                                self.order.goodTillDate = expiry_date
                                self.order.tif = "GTD"
                                self.order.account = 'DU9727656'
                                self.order.transmit = True
                                self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                await asyncio.sleep(3) # keep this same
                                print(self.order_details)
                                print("The order has been placed")
                            self.excel_data.loc[i,'Activation'] = -3 # now use excelwriter fn
                            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)
                    else:
                        print("The trigger price has not being triggered")
        else:
            print("No changes in excel")
                
    async def close_empty_trigger_fn_upper(self):
        print("fn in close_all_if_trigger")
        # self.df = pd.read_excel(self.path, sheet_name=credentials.sheet_name) 
        self.df   = pd.read_excel(resource_path(r"storage\new_order_sheet.xlsx"),sheet_name=credentials.sheet_name)
        df = self.df
        self.expiryvar = None
        for i in range(len(self.df)):
            if(self.df.loc[i,'Target'] == "-" and self.df.loc[i,'Stop_Loss'] == "-" and self.df.loc[i,'Strike_Type'] == "BUY" and self.df.loc[i,'Activation'] == 2):
                print(f"upper trigger limit at index : {i} at {self.df.loc[i,'Entry_Strike']}")
                self.upper_trigger = self.df.loc[i,'Entry_Strike']
                self.expiryvar       = self.df.loc[i,'Expiry']
                break # close one kind of expiry at a time.

        if self.expiryvar is not None:
            datevar = self.expiryvar
            datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
            year,day,month = datevar_str.split('-')
            formatted_date = f"{year}{month.zfill(2)}{day}"
            # print(f"formatted_date in outer loop is {formatted_date}")
            contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))

            price = await self.get_current_market_price_futures(contract)
            positions = self.client.positions()
            print(f"current price is {price}")
            if price>self.upper_trigger:
                if positions:
                    for i in range(len(df)):
                        if self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Expiry'] == self.expiryvar:
                            contract = Future(symbol='N225M', exchange='OSE.JPN', lastTradeDateOrContractMonth=str(formatted_date))
                            if self.df.loc[i, 'Strike_Type'] == 'SELL':
                                current_action = 'BUY'
                            else:
                                current_action = 'SELL'
                            # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                            order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                            order.account = 'DU9727656'
                            order.transmit = True
                            result = self.client.placeOrder(contract, order)
                            self.df.loc[i, 'Activation'] = 0
                            
                        elif self.df.loc[i,'Activation'] == -3 and self.df.loc[i,'Expiry'] == self.expiryvar:
                            contract = Future(symbol='N225M', exchange='OSE.JPN', lastTradeDateOrContractMonth=str(formatted_date))
                            if self.df.loc[i, 'Strike_Type'] == 'SELL':
                                current_action = 'BUY'
                            else:
                                current_action = 'SELL'
                            # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                            order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                            order.account = 'DU9727656'
                            order.transmit = True
                            result = self.client.placeOrder(contract, order)
                            self.df.loc[i, 'Activation'] = 0
                else:
                    print("Positions are empty")

            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                self.df.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
        else:
            print("No condition for closing till yet")

    async def close_empty_trigger_fn_lower(self):
        print("fn in close_all_if_trigger")
        # self.df = pd.read_excel(self.path, sheet_name=credentials.sheet_name) 
        self.df   = pd.read_excel(resource_path(r"storage\new_order_sheet.xlsx"),sheet_name=credentials.sheet_name)
        df = self.df
        self.expiryvar = None
        for i in range(len(self.df)):
            if(self.df.loc[i,'Target'] == "-" and self.df.loc[i,'Stop_Loss'] == "-" and self.df.loc[i,'Strike_Type'] == "SELL" and self.df.loc[i,'Activation'] == 2): # have the closing only if the activation is one 
                print(f"lower trigger limit at index : {i} at {self.df.loc[i,'Entry_Strike']}")
                self.lower_trigger = self.df.loc[i,'Entry_Strike']
                self.expiryvar       = self.df.loc[i,'Expiry']
                break # close one kind of expiry at a time.

        if self.expiryvar is not None:
            datevar = self.expiryvar
            datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
            year,day,month = datevar_str.split('-')
            formatted_date = f"{year}{month.zfill(2)}{day}"
            # print(f"formatted_date in outer loop is {formatted_date}")
            contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))

            price = await self.get_current_market_price_futures(contract)
            positions = self.client.positions()
            print(f"current price is {price}")
            if price<self.lower_trigger:
                if positions:
                    for i in range(len(df)):
                        if self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Expiry'] == self.expiryvar:
                            contract = Future(symbol='N225M', exchange='OSE.JPN', lastTradeDateOrContractMonth=str(formatted_date))
                            if self.df.loc[i, 'Strike_Type'] == 'SELL':
                                current_action = 'BUY'
                            else:
                                current_action = 'SELL'
                            # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                            order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                            order.account = 'DU9727656'
                            order.transmit = True
                            result = self.client.placeOrder(contract, order)
                            self.df.loc[i, 'Activation'] = 0

                        elif self.df.loc[i,'Activation'] == -3 and self.df.loc[i,'Expiry'] == self.expiryvar:
                            contract = Future(symbol='N225M', exchange='OSE.JPN', lastTradeDateOrContractMonth=str(formatted_date))
                            if self.df.loc[i, 'Strike_Type'] == 'SELL':
                                current_action = 'BUY'
                            else:
                                current_action = 'SELL'
                            # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                            order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                            order.account = 'DU9727656'
                            order.transmit = True
                            result = self.client.placeOrder(contract, order)
                            self.df.loc[i, 'Activation'] = 0
                else:
                    print("Positions are empty")

            # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                self.df.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
        else:
            print("No condition for closing till yet")
    

    async def get_current_market_price_futures(self, contract): 
        """
        Fetches the current market price of the given futures contract.
        """
        self.client.reqMarketDataType(3)
        ticker = self.client.reqMktData(contract, '', False, False)
        self.client.sleep(1)  
        if ticker.last is not None: 
            return ticker.last
        if ticker.close is not None: 
            return ticker.close
        
        print(ticker.last)
        return None
        # return 0.0
    
    async def show_details(self):
        result = self.ib.reqOpenOrders()
        return result
    
    async def get_bid_and_ask(self,contractmonth):
        self.client.reqMarketDataType(3)
        contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=contractmonth)
        test = self.ib.reqTickers(contract)
        for _,r in enumerate(test):
            bid,ask = r.bid,r.ask
            print(bid)
            print(ask)
        
        return bid,ask
    
    async def check_for_tp_sl(self, current_price, target_price,stop_loss,action): 
        if action == 'BUY':
            if current_price >= target_price:  # Target Price hit
                return "SELL"
            elif current_price <= stop_loss:  # Stop Loss hit
                return "SELL"

        elif action == 'SELL':
            if current_price <= target_price:  # corrected
                return "BUY"
            elif current_price >= stop_loss:  # corrected
                return "BUY"
        return None

    async def monitor_tp_sl(self): 
        # self.df = pd.read_excel(self.path, sheet_name=credentials.sheet_name)  
        self.df   = pd.read_excel(resource_path(r"storage\new_order_sheet.xlsx"),sheet_name=credentials.sheet_name)
        for i in range(len(self.df)):
            if self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Strike_Type'] == 'BUY':
                datevar = self.df.loc[i, 'Expiry']
                # Ensure datevar is a string in 'YYYY-MM-DD' format
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                # formatted_date = f"{year}{month.zfill(2)}"
                formatted_date = f"{year}{month.zfill(2)}{day}"
                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date))
                current_price = await self.get_current_market_price_futures(contract)
                if current_price:
                    action = await self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        print(f"An action of sell has been triggered in row {i}")
                        # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                        order = MarketOrder(action='SELL', totalQuantity=str(self.df['Qty'].iloc[i]))
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0
                        print("One position is being closed")  
                    else:
                        print("No profit/loss is triggered")

            elif self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Strike_Type'] == 'SELL':
                datevar = self.df.loc[i, 'Expiry']
                print(datevar)
                # Ensure datevar is a string in 'YYYY-MM-DD' format
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                # formatted_date = f"{year}{month.zfill(2)}" 
                formatted_date = f"{year}{month.zfill(2)}{day}"

                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date)) # change this line
                current_price = await self.get_current_market_price_futures(contract)
                if current_price:
                    action = await self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        print(f"An action of buy has been triggered in row {i}")
                        # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                        order = MarketOrder(action='BUY', totalQuantity=str(self.df['Qty'].iloc[i]))
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0  
                    else:
                        print("No profit/loss is triggered")
            
            elif self.df.loc[i,'Activation'] == -3 and self.df.loc[i,'Strike_Type'] == 'BUY':
                datevar = self.df.loc[i, 'Expiry']
                # Ensure datevar is a string in 'YYYY-MM-DD' format
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                # formatted_date = f"{year}{month.zfill(2)}"
                formatted_date = f"{year}{month.zfill(2)}{day}"
                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date))
                current_price = await self.get_current_market_price_futures(contract)
                if current_price:
                    action = await self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        print(f"An action of sell has been triggered in row {i}")
                        # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                        order = MarketOrder(action='SELL', totalQuantity=str(self.df['Qty'].iloc[i]))
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0
                        print("One position is being closed")  
                    else:
                        print("No profit/loss is triggered")

            elif self.df.loc[i,'Activation'] == -3 and self.df.loc[i,'Strike_Type'] == 'SELL':
                datevar = self.df.loc[i, 'Expiry']
                print(datevar)
                # Ensure datevar is a string in 'YYYY-MM-DD' format
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                # formatted_date = f"{year}{month.zfill(2)}" 
                formatted_date = f"{year}{month.zfill(2)}{day}"

                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date)) # change this line
                current_price = await self.get_current_market_price_futures(contract)
                if current_price:
                    action = await self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        print(f"An action of buy has been triggered in row {i}")
                        # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                        order = MarketOrder(action='BUY', totalQuantity=str(self.df['Qty'].iloc[i]))
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0  
                    else:
                        print("No profit/loss is triggered")

        # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.df.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
    
    async def new_auto_square_off(self): # put this in async
        # self.df = pd.read_excel(self.path, sheet_name=credentials.sheet_name)
        self.df   = pd.read_excel(resource_path(r"storage\new_order_sheet.xlsx"),sheet_name=credentials.sheet_name)
        df = self.df
        current_time = datetime.now().strftime("%H:%M")
        positions = self.client.positions()
        # if current_time > "9:10":
        print(f"current time is {current_time} and closing_time is {credentials.current_time}")
        if current_time >= str(credentials.current_time):
            if positions:
                for i in range(len(df)):
                    if self.df.loc[i,'Activation'] == -1:
                        datevar = self.df.loc[i, 'Expiry']
                        # Ensure datevar is a string in 'YYYY-MM-DD' format
                        datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                        year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                        # formatted_date = f"{year}{month.zfill(2)}" # changed
                        formatted_date = f"{year}{month.zfill(2)}{day}"
                        contract = Future(symbol='N225M', exchange='OSE.JPN', lastTradeDateOrContractMonth=str(formatted_date))
                        if self.df.loc[i, 'Strike_Type'] == 'SELL':
                            current_action = 'BUY'
                        else:
                            current_action = 'SELL'
                        # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                        order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                        
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        self.df.loc[i, 'Activation'] = 0
                    
                    # elif self.df.loc[i,'Activation'] == -3:
                    #     datevar = self.df.loc[i, 'Expiry']
                    #     # Ensure datevar is a string in 'YYYY-MM-DD' format
                    #     datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                    #     year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                    #     formatted_date = f"{year}{month.zfill(2)}" # changed
                    #     # formatted_date = f"{year}{month.zfill(2)}{day}"
                    #     contract = Future(symbol='N225M', exchange='OSE.JPN', lastTradeDateOrContractMonth=str(formatted_date))
                    #     if self.df.loc[i, 'Strike_Type'] == 'SELL':
                    #         current_action = 'BUY'
                    #     else:
                    #         current_action = 'SELL'
                    #     # contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),currency="JPY")
                    #     order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                        
                    #     order.account = 'DU9727656'
                    #     order.transmit = True
                    #     result = self.client.placeOrder(contract, order)
                    #     self.df.loc[i, 'Activation'] = 0
                    
            else:
                print("Positions are empty")
        else:
            print("The time is not for closing the market is not yet")

        # with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        with pd.ExcelWriter(resource_path(r"storage\new_order_sheet.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.df.to_excel(writer, sheet_name=credentials.sheet_name, index=False)

    async def run(self):
        print("The process has started")
        await self.connection_show()
        while True:
            await asyncio.gather(self.check_for_new_positions(),self.new_auto_square_off(),self.monitor_tp_sl(),self.close_empty_trigger_fn_lower(),self.close_empty_trigger_fn_upper())
            await asyncio.sleep(10) # keep this same

if __name__ == "__main__":
    if credentials.master is not False:
        session = IBRKExcel()
        asyncio.run(session.run())
    else:
        print("The bot is currently off make changes in the master.")
