import time
import asyncio
import pandas as pd
from datetime import datetime,timedelta
from ib_broker import *
import credentials
from openpyxl import load_workbook
import nest_asyncio

nest_asyncio.apply()

class IBRKExcel:
    def __init__(self):
        self.symbol          = credentials.symbol
        self.exchange        = credentials.exchange
        self.path            = credentials.xlsx_path_1
        self.current_time    = datetime.now()
        self.excel_data      = pd.read_excel(self.path, sheet_name=credentials.sheet_name) 
        self.length          = len(self.excel_data)
        self.upper_trigger   = 100000000
        self.lower_trigger   = -100000000
        self.current_time    = datetime.now().time()

    async def check_excel_changes(self):
        new_data   = pd.read_excel(self.path, sheet_name=credentials.sheet_name)
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
            length   = len(pd.read_excel(self.path, sheet_name=credentials.sheet_name))
            for i in range(length):
                if self.excel_data.loc[i,'Activation'] == 1: # a new order detected
                    row           = self.excel_data.iloc[i]
                    self.symbol             = 'N225M'
                    self.exchange           = 'OSE.JPN' 
                    self.trigger_level      = row['Trigger_Level_High_Low']
                    self.entry_type         = row['Entry_Type']
                    self.entry_strike       = row['Entry_Strike']
                    self.option_strike_type = row['Option_Type']
                    self.strike_type        = row['Strike_Type']
                    self.expiry             = str(row['Expiry'])
                    self.target             = row['Target']
                    self.stop_loss          = row['Stop_Loss']
                    self.qty                = row['Qty']
                    self.slicing            = row['Slicing']
                    self.time_interval      = row['Time_Interval']
                    self.activation_type    = row['Activation_Type']
                    self.activation         = row['Activation']

                    if self.strike_type == "SELL":
                        self.side = 'SELL'
                    else:
                        self.side = 'BUY' 

                    datevar = self.expiry
                    datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                    date,timep = datevar.split(' ') 
                    year,day,month = date.split('-')
                    print(f"year is {year}")
                    print(f"month is {month}")
                    print(f"day is {day}")
                    formatted_date = f"{year}{month.zfill(2)}{day}"
                    
                    print(f"formatted_date is {formatted_date}")
                    contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                    print(self.trigger_level)
                    print(await self.get_current_market_price_futures(contract)) 
                    print(self.entry_type)
                    print(self.strike_type)
                    if self.strike_type == "BUY" and self.trigger_level <= await self.get_current_market_price_futures(contract): # current_price breaks through trigger_level
                        if self.entry_type == "LIMIT":
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                date,timep = datevar.split(' ')
                                print(f"date is {date}")
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                self.contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                                bid,ask = await self.get_bid_and_ask_options(contractmonth=formatted_date,entry_strike=self.entry_strike,option_strike_type=self.option_strike_type)
                                attempt = 0
                                while attempt<int(credentials.attempts):
                                    print(f"the current bid is {bid} and the current ask is {ask} and the order is being placed at {float(str(int((bid + (2**attempt - 1)*ask)/2**attempt)))}.")
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
                                if attempt == credentials.attempts:
                                    print(f"Limit order failed {credentials.attempts} times placing market order")
                                    self.order = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing)))
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print("Market order placed")
                                    print(self.order_details)

                            self.excel_data.loc[i,'Activation'] = -1 
                            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)

                        else:
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                date,timep = datevar.split(' ')
                                year,day,month = date.split('-')
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                self.contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                                self.order          = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) 
                                self.order.account = 'DU9727656'
                                self.order.transmit = True
                                self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                if credentials.user_time_default == 0:
                                    await asyncio.sleep(self.time_interval)
                                else:
                                    print(f"sleep for {credentials.user_time}")
                                    await asyncio.sleep(credentials.user_time)
                                print("The order has been placed")
                            self.excel_data.loc[i,'Activation'] = -1 
                            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
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
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                date,timep = datevar.split(' ')
                                year,day,month = date.split('-')
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                self.contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                                bid,ask = await self.get_bid_and_ask_options(contractmonth=formatted_date,entry_strike=self.entry_strike,option_strike_type=self.option_strike_type)
                                attempt = 0
                                while attempt<int(credentials.attempts):
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

                                if attempt == credentials.attempts:
                                    print(f"Limit order failed {credentials.attempts} times placing market order")
                                    self.order = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing)))
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print("Market order placed")
                                    print(self.order_details)

                            self.excel_data.loc[i,'Activation'] = -1 # now use excelwriter fn
                            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)

                        else:
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                datevar = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                date,timep = datevar.split(' ')
                                year,day,month = date.split('-')
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                print(f"formatted_date is {formatted_date}")
                                self.contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                                self.order          = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) 
                                self.order.account = 'DU9727656'
                                self.order.transmit = True
                                self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                await asyncio.sleep(3) 
                                print(self.order_details)
                                print("The order has been placed")
                            self.excel_data.loc[i,'Activation'] = -1 
                            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                print(f"sleeping for {credentials.user_time}")
                                await asyncio.sleep(credentials.user_time)
                    else:
                        print("The trigger price has not being triggered")
                
    async def close_empty_trigger_fn_upper(self):
        print("fn in close_all_if_trigger upper")
        self.df = pd.read_excel(self.path, sheet_name=credentials.sheet_name) 
        df = self.df
        self.expiryvar = None
        for i in range(len(self.df)):
            if(self.df.loc[i,'Target'] == "-" and self.df.loc[i,'Stop_Loss'] == "-" and self.df.loc[i,'Strike_Type'] == "BUY" and self.df.loc[i,'Activation'] == 1 and self.df.loc[i,'Activation_Type'] == 2):
                print(f"upper trigger limit at index : {i} at {self.df.loc[i,'Entry_Strike']}")
                datevar = self.df.loc[i,'Expiry']
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')
                formatted_date = f"{year}{month.zfill(2)}{day}"
                contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))

                price = await self.get_current_market_price_futures(contract)
                positions = self.client.positions()
                print(f"current price is {price}")
                if price>self.upper_trigger:
                    if positions:
                        for i in range(len(df)):
                            if self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Expiry'] == self.expiryvar:
                                datevar = self.df.loc[i,'Expiry']
                                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar_str.split('-')
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                self.option_strike_type = self.df.loc[i,'Option_Type']
                                self.entry_strike = self.df.loc[i,'Entry_Strike']
                                contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                                if self.df.loc[i, 'Strike_Type'] == 'SELL':
                                    current_action = 'BUY'
                                else:
                                    current_action = 'SELL'
                                order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                                order.account = 'DU9727656'
                                order.transmit = True
                                result = self.client.placeOrder(contract, order)
                                self.df.loc[i, 'Activation'] = 0
                                
                            elif self.df.loc[i,'Activation'] == 3 and self.df.loc[i,'Expiry'] == self.expiryvar:
                                datevar = self.df.loc[i,'Expiry']
                                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar_str.split('-')
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                self.option_strike_type = self.df.loc[i,'Option_Type']
                                self.entry_strike = self.df.loc[i,'Entry_Strike']
                                contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                                if self.df.loc[i, 'Strike_Type'] == 'SELL':
                                    current_action = 'BUY'
                                else:
                                    current_action = 'SELL'
                                order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                                order.account = 'DU9727656'
                                order.transmit = True
                                result = self.client.placeOrder(contract, order)
                                self.df.loc[i, 'Activation'] = 0
                else:
                    print("Positions are empty")

            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                self.df.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
        else:
            print("No condition for closing till yet")

    async def close_empty_trigger_fn_lower(self):
        print("fn in close_all_if_trigger lower")
        self.df = pd.read_excel(self.path, sheet_name=credentials.sheet_name) 
        df = self.df
        self.expiryvar = None
        for i in range(len(self.df)):
            if(self.df.loc[i,'Target'] == "-" and self.df.loc[i,'Stop_Loss'] == "-" and self.df.loc[i,'Strike_Type'] == "SELL" and self.df.loc[i,'Activation_Type'] == 2 and self.df.loc[i,'Activation'] == 1): # have the closing only if the activation is one 
                print(f"lower trigger limit at index : {i} at {self.df.loc[i,'Entry_Strike']}")
                datevar = self.df.loc[i,'Expiry']
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')
                formatted_date = f"{year}{month.zfill(2)}{day}"
                contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))

                price = await self.get_current_market_price_futures(contract)
                positions = self.client.positions()
                print(f"current price is {price}")
                if price<self.lower_trigger:
                    if positions:
                        for i in range(len(df)):
                            if self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Expiry'] == self.expiryvar:
                                datevar = self.df.loc[i,'Expiry']
                                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar_str.split('-')
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                self.option_strike_type = self.df.loc[i,'Option_Type']
                                self.entry_strike = self.df.loc[i,'Entry_Strike']
                                contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                                if self.df.loc[i, 'Strike_Type'] == 'SELL':
                                    current_action = 'BUY'
                                else:
                                    current_action = 'SELL'
                                order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                                order.account = 'DU9727656'
                                order.transmit = True
                                result = self.client.placeOrder(contract, order)
                                self.df.loc[i, 'Activation'] = 0

                            elif self.df.loc[i,'Activation'] == 3 and self.df.loc[i,'Expiry'] == self.expiryvar:
                                datevar = self.df.loc[i,'Expiry']
                                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                                year,day,month = datevar_str.split('-')
                                formatted_date = f"{year}{month.zfill(2)}{day}"
                                self.option_strike_type = self.df.loc[i,'Option_Type']
                                self.entry_strike = self.df.loc[i,'Entry_Strike']
                                contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                                if self.df.loc[i, 'Strike_Type'] == 'SELL':
                                    current_action = 'BUY'
                                else:
                                    current_action = 'SELL'
                                order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                                order.account = 'DU9727656'
                                order.transmit = True
                                result = self.client.placeOrder(contract, order)
                                self.df.loc[i, 'Activation'] = 0
                else:
                    print("Positions are empty")

            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
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
    
    async def get_bid_and_ask_options(self,contractmonth,entry_strike,option_strike_type):
        self.client.reqMarketDataType(3)
        contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=contractmonth,strike=float(entry_strike),right=option_strike_type)
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
        self.df = pd.read_excel(self.path, sheet_name=credentials.sheet_name)  
        for i in range(len(self.df)):
            if self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Strike_Type'] == 'BUY' and self.df.loc[i,'Activation_Type'] == 1:
                datevar = self.df.loc[i, 'Expiry']
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-') 
                formatted_date = f"{year}{month.zfill(2)}{day}"
                self.entry_strike = self.df.loc[i,'Entry_Strike']
                self.option_strike_type = self.df.loc[i,'Option_Type']
                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date))
                current_price = await self.get_current_market_price_futures(contract)
                if current_price:
                    action = await self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        print(f"An action of sell has been triggered in row {i}")
                        contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                        order = MarketOrder(action='SELL', totalQuantity=str(self.df['Qty'].iloc[i]))
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0
                        print("One position is being closed")  
                    else:
                        print("No profit/loss is triggered")

            elif self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Strike_Type'] == 'SELL' and self.df.loc[i,'Activation_Type'] == 1:
                datevar = self.df.loc[i, 'Expiry']
                print(datevar)
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')
                formatted_date = f"{year}{month.zfill(2)}{day}"

                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date))
                current_price = await self.get_current_market_price_futures(contract)
                contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                if current_price:
                    action = await self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        print(f"An action of buy has been triggered in row {i}")
                        order = MarketOrder(action='BUY', totalQuantity=str(self.df['Qty'].iloc[i]))
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0  
                    else:
                        print("No profit/loss is triggered")

        with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.df.to_excel(writer, sheet_name=credentials.sheet_name, index=False)
    
    async def new_auto_square_off(self): # put this in async
        self.df = pd.read_excel(self.path, sheet_name=credentials.sheet_name)
        df = self.df
        current_time = datetime.now().strftime("%H:%M")
        positions = self.client.positions()
        print(f"current time is {current_time} and closing_time is {credentials.current_time}")
        if current_time >= str(credentials.current_time):
            if positions:
                for i in range(len(df)):
                    if self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Activation_Type'] == 1:
                        datevar = self.df.loc[i, 'Expiry']
                        datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                        year,day,month = datevar_str.split('-')
                        formatted_date = f"{year}{month.zfill(2)}{day}"
                        self.entry_strike = str(self.df.loc[i,'Entry_Strike'])
                        self.option_strike_type = str(self.df.loc[i,'Option_Type'])
                        contract = Option(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date),strike=float(self.entry_strike),right=self.option_strike_type)
                        if self.df.loc[i, 'Strike_Type'] == 'SELL':
                            current_action = 'BUY'
                        else:
                            current_action = 'SELL'
                        order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                        
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        self.df.loc[i, 'Activation'] = 0
            else:
                print("Positions are empty")
        else:
            print("The time is not for closing the market is not yet")

        with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.df.to_excel(writer, sheet_name=credentials.sheet_name, index=False)

    async def run(self):
        print("The process has started")
        await self.connection_show()
        while True:
            await asyncio.gather(self.check_for_new_positions(),self.new_auto_square_off(),self.monitor_tp_sl(),self.close_empty_trigger_fn_lower(),self.close_empty_trigger_fn_upper())
            await asyncio.sleep(7) 

if __name__ == "__main__":
    session = IBRKExcel()
    asyncio.run(session.run())
