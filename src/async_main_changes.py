import time
import asyncio
import pandas as pd
from datetime import datetime
from ib_broker import *
import credentials
from openpyxl import load_workbook
import nest_asyncio

nest_asyncio.apply()

class IBRKExcel:
    def __init__(self):
        self.symbol          = 'N225M'
        self.exchange        = 'OSE.JPN'
        self.path            = credentials.xlsx_path
        self.current_time    = datetime.now()
        self.excel_data      = pd.read_excel(self.path, sheet_name='Sheet6') 
        self.length          = len(self.excel_data)
        self.orderbook       = []
        self.failed_orders   = []
        self.database_path   = credentials.database_path
        self.current_time    = datetime.now().time()

    async def check_excel_changes(self):
        new_data   = pd.read_excel(self.path, sheet_name='Sheet6')
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

    async def format_date_ddmmyyyy(self,var):
        date,timep = var.split(" ")
        year,day,month = date.split('-')
        formatted_date = f"{year}{month.zfill(2)}"
        return str(formatted_date)

    async def check_for_new_positions(self): # put this in async
        if await self.check_excel_changes():
            print("a change on the excel has been made")
            length   = len(pd.read_excel(self.path, sheet_name='Sheet6'))
            for i in range(length):
                if self.excel_data.loc[i,'Activation'] == 1:
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
                    if self.strike_type == 'CE':
                        self.side = 'SELL'
                    else:
                        self.side = 'BUY' 

                    datevar = self.expiry
                    date,timep = datevar.split(' ')
                    year,day,month = date.split('-')
                    formatted_date = f"{year}{month.zfill(2)}"
                    contract = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                    print(self.trigger_level)
                    print(await self.get_current_market_price_futures(contract)) # need to be fix the price is not giving - fix with paper_trading_account.
                    print(self.entry_type)
                    print(self.strike_type)
                    # if self.strike_type == "PE" and self.trigger_level < await self.get_current_market_price_futures(contract):
                    if self.strike_type == "PE" and self.trigger_level <= await self.get_current_market_price_futures(contract):
                        if self.entry_type == "LIMIT":
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                date,timep = datevar.split(' ')
                                year,day,month = date.split('-')
                                formatted_date = f"{year}{month.zfill(2)}"
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                attempt = 0
                                # while attempt<3:
                                while attempt<int(credentials.attempts):
                                    if credentials.trade_type_default == 0:
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(self.entry_strike)) 
                                    else:
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(int((bid + (2**attempt - 1)*ask)/2**attempt))) 
                                    self.order.account = 'DU9727656'
                                    self.order.transmit = True
                                    print(f"Placing limit order,attempt {attempt+1}")
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print(self.order_details)
                                    await asyncio.sleep(2) # keep this part same
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
                                
                                if attempt == 3:
                                    print("Limit order failed thrice placing market order")
                                    self.order = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing)))
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print("Market order placed")
                                    print(self.order_details)

                            self.excel_data.loc[i,'Activation'] = -1 
                            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name="Sheet6", index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                await asyncio.sleep(credentials.user_time)

                        else:
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                date,timep = datevar.split(' ')
                                year,day,month = date.split('-')
                                formatted_date = f"{year}{month.zfill(2)}"
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                self.order          = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) 
                                self.order.account = 'DU9727656'
                                self.order.transmit = True
                                self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                # await asyncio.sleep(3)
                                if credentials.user_time_default == 0:
                                    await asyncio.sleep(self.time_interval)
                                else:
                                    await asyncio.sleep(credentials.user_time)
                                print("The order has been placed")
                            self.excel_data.loc[i,'Activation'] = -1 
                            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name="Sheet6", index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                await asyncio.sleep(credentials.user_time)

                    elif self.strike_type == "CE" and self.trigger_level >= await self.get_current_market_price_futures(contract):
                        if self.entry_type == "LIMIT":
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                date,timep = datevar.split(' ')
                                year,day,month = date.split('-')
                                formatted_date = f"{year}{month.zfill(2)}"
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                bid,ask = await self.get_bid_and_ask(contractmonth=formatted_date)
                                attempt = 0
                                while attempt<int(credentials.attempts):
                                    if credentials.trade_type_default == 0:
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(self.entry_strike))
                                    else:
                                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(int((bid + (2**attempt - 1)*ask)/2**attempt)))  
                                    self.order.account = 'DU9727656'
                                    self.order.transmit = True
                                    print(f"Placing limit order,attempt {attempt+1}")
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print(self.order_details)
                                    await asyncio.sleep(2) # keep this same
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
                                
                                if attempt == 3:
                                    print("Limit order failed thrice placing market order")
                                    self.order = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing)))
                                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                    print("Market order placed")
                                    print(self.order_details)

                            self.excel_data.loc[i,'Activation'] = -1 # now use excelwriter fn
                            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name="Sheet6", index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                await asyncio.sleep(credentials.user_time)

                        else:
                            for _ in range(0,int(self.qty/self.slicing),1):
                                datevar = self.expiry
                                date,timep = datevar.split(' ')
                                year,day,month = date.split('-')
                                formatted_date = f"{year}{month.zfill(2)}"
                                self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                                self.order          = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) 
                                self.order.account = 'DU9727656'
                                self.order.transmit = True
                                self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                                await asyncio.sleep(3) # keep this same
                                print(self.order_details)
                                print("The order has been placed")
                            self.excel_data.loc[i,'Activation'] = -1 # now use excelwriter fn
                            with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                self.excel_data.to_excel(writer, sheet_name="Sheet6", index=False)
                            if credentials.user_time_default == 0:
                                await asyncio.sleep(self.time_interval)
                            else:
                                await asyncio.sleep(credentials.user_time)
                    else:
                        print("The trigger price has not being triggered")
        else:
            print("No changes in excel")

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
        
        return bid,ask
    
    async def check_for_tp_sl(self, current_price, target_price,stop_loss,action): 
        if action == 'PE':
            if current_price >= target_price:  # Target Price hit
                return "SELL"
            elif current_price <= stop_loss:  # Stop Loss hit
                return "SELL"

        elif action == 'CE':
            if current_price <= target_price:  # corrected
                return "BUY"
            elif current_price >= stop_loss:  # corrected
                return "BUY"
        return None

    async def monitor_tp_sl(self): 
        self.df = pd.read_excel(self.path, sheet_name="Sheet6")  
        for i in range(len(self.df)):
            if self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Strike_Type'] == 'PE':
                datevar = self.df.loc[i, 'Expiry']
                # Ensure datevar is a string in 'YYYY-MM-DD' format
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                formatted_date = f"{year}{month.zfill(2)}"
                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date))
                current_price = await self.get_current_market_price_futures(contract)
                if current_price:
                    action = await self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        print(f"An action of sell has been triggered in row {i}")
                        order = MarketOrder(action='SELL', totalQuantity=str(self.df['Qty'].iloc[i]))
                        order.account = 'DU9727656'
                        order.transmit = True
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0
                        print("One position is being closed")  
                    else:
                        print("No profit/loss is triggered")

            elif self.df.loc[i,'Activation'] == -1 and self.df.loc[i,'Strike_Type'] == 'CE':
                datevar = self.df.loc[i, 'Expiry']
                print(datevar)
                # Ensure datevar is a string in 'YYYY-MM-DD' format
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                formatted_date = f"{year}{month.zfill(2)}"

                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date)) # change this line
                current_price = await self.get_current_market_price_futures(contract)
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
            self.df.to_excel(writer, sheet_name="Sheet6", index=False)
    
    async def new_auto_square_off(self): # put this in async
        self.df = pd.read_excel(self.path, sheet_name="Sheet6")
        df = self.df
        current_time = datetime.now().strftime("%H:%M")
        positions = self.client.positions()
        # if current_time > "9:10":
        if current_time >= str(credentials.current_time):
            if positions:
                for i in range(len(df)):
                    if self.df.loc[i,'Activation'] == -1:
                        datevar = self.df.loc[i, 'Expiry']
                        # Ensure datevar is a string in 'YYYY-MM-DD' format
                        datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                        year,day,month = datevar_str.split('-')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                        formatted_date = f"{year}{month.zfill(2)}"
                        contract = Future(symbol='N225M', exchange='OSE.JPN', lastTradeDateOrContractMonth=str(formatted_date))
                        if self.df.loc[i, 'Strike_Type'] == 'CE':
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
            self.df.to_excel(writer, sheet_name="Sheet6", index=False)

    async def run(self):
        print("The process has started")
        await self.connection_show()
        while True:
            await asyncio.gather(self.check_for_new_positions(),self.new_auto_square_off(),self.monitor_tp_sl())
            await asyncio.sleep(10) # keep this same
    
    async def test(self):
        print("The process has started")
        await self.connection_show()
        while True:
            await asyncio.gather(self.get_bid_and_ask('202503'))
            exit(0)

if __name__ == "__main__":
    if credentials.master is not False:
        session = IBRKExcel()
        # asyncio.run(session.run())
        asyncio.run(session.test())
        # asyncio.run(session.get_bid_and_ask('202503'))
        # asyncio.run(session.get_bid_and_ask('202503'))
        exit(0)
    else:
        print("The bot is currently off.")

