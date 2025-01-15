import time
import asyncio
import pandas as pd
from datetime import datetime,time
from ib_broker import *
import credentials
from openpyxl import load_workbook

class IBRKExcel:
    def __init__(self):
        self.path            = credentials.xlsx_path
        self.current_time    = datetime.now()
        self.excel_url = "https://docs.google.com/spreadsheets/d/1gKSi1A1gidZHTm9Vrzm0dEUF8QzcCkCMiSnMMzZtKAA/export?format=csv"
        self.excel_data      = pd.read_csv(self.excel_url)
        self.length          = len(self.excel_data)
        self.orderbook       = []
        self.failed_orders   = []
        self.database_path   = credentials.database_path
        self.current_time    = datetime.now().time()

    def check_excel_changes(self):
        new_data = pd.read_csv(self.excel_url)
        new_length = len(new_data)

        if new_length  != self.length:
            self.length     = new_length
            self.excel_data = new_data
            return True
        return False

    def connection_show(self) -> bool:
        host, port = credentials.host, credentials.port
        self.client = IB()
        self.ib = self.client
        connection_print = self.client.connect(host=host, port=port, clientId=13, timeout=60)
        print(connection_print)

    def format_date_ddmmyyyy(self,var):
        date,time = var.split(" ")
        year,day,month = date.split('-')
        formatted_date = f"{year}{month.zfill(2)}"
        return str(formatted_date)
    
    def check_for_new_positions(self): # put this in async
        if self.check_excel_changes():
            last_row           = self.excel_data.iloc[-1]
            self.symbol        = 'N225M'
            self.exchange      = 'OSE.JPN' 
            self.trigger_level = last_row['Trigger_Level_High_Low']
            self.entry_type    = last_row['Entry_Type']
            self.entry_strike  = last_row['Entry_Strike']
            self.strike_type   = last_row['Strike_Type']
            self.expiry        = str(last_row['Expiry'])
            self.target        = last_row['Target']
            self.stop_loss     = last_row['Stop_Loss']
            self.qty           = last_row['Qty']
            self.slicing       = last_row['Slicing']
            self.time_interval = last_row['Time_Interval']
            self.activation    = last_row['Activation']
            if self.strike_type == 'CE':
                self.side = 'SELL'
            else:
                self.side = 'BUY' 
            self.i = 0
            if self.activation == 1:
                if self.entry_type == "LIMIT":
                    # for _ in range(0,int(self.qty/self.slicing),1):
                    for i in range(0,int(self.qty/self.slicing)):
                        datevar = self.expiry
                        datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                        day,month,year = datevar_str.split('/')  
                        formatted_date = f"{year}{month.zfill(2)}"
                        self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                        # self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.slicing)),lmtPrice=str(self.entry_strike)) # maybe change to self.qty/self.slicing
                        self.order         = LimitOrder(action=self.side,totalQuantity=str(int(self.qty)),lmtPrice=str(self.entry_strike)) 
                        self.order.transmit = True
                        self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                        print(self.order_details)
                        print("The order has been placed")
                        time.sleep(self.time_interval*2)
                        
                else:
                    # for _ in range(0,int(self.qty/self.slicing),1):
                    for i in range(0,int(self.qty/self.slicing)):
                        datevar = self.expiry
                        datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                        day,month,year = datevar_str.split('/')  
                        formatted_date = f"{year}{month.zfill(2)}"
                        self.contract       = Future(symbol=self.symbol,exchange=self.exchange,lastTradeDateOrContractMonth=str(formatted_date))
                        # self.order         = MarketOrder(action=self.side,totalQuantity=str(int(self.slicing))) # maybe change to self.qty/self.slicing
                        self.order         = MarketOrder(action=self.side,totalQuantity=str(int(self.qty))) 
                        self.order.transmit = True
                        self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                        print(self.order_details)
                        print("The order has been placed")
                        time.sleep(self.time_interval*2)


    def get_current_market_price_futures(self, contract): 
        """
        Fetches the current market price of the given futures contract.
        """
        self.client.reqMarketDataType(3)
        ticker = self.client.reqMktData(contract, '', False, False)
        self.client.sleep(1)  
        if ticker.last is not None:
            return ticker.last
        elif ticker.close is not None:
            return ticker.close
        return None
    
    def show_details(self):
        result = self.ib.reqOpenOrders()
        return result
    
    def check_for_tp_sl(self, current_price, target_price,stop_loss,action):
        if action == 'PE':
            if current_price >= target_price:  # Target Price hit
                return "SELL"
            elif current_price <= stop_loss:  # Stop Loss hit
                return "SELL"

        elif action == 'CE':
            if current_price >= target_price:  # Target Price hit
                return "BUY"
            elif current_price <= stop_loss:  # Stop Loss hit
                return "BUY"
        return None

    def monitor_tp_sl(self): # put this in async
        self.df = pd.read_csv(self.excel_url)
        
        for i in range(len(self.df)):
            if self.df.loc[i,'Activation'] == 1 and self.df.loc[i,'Strike_Type'] == 'PE':

                datevar = self.df.loc[i, 'Expiry']
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('/')  
                formatted_date = f"{year}{month.zfill(2)}"
                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date))
                current_price = self.get_current_market_price_futures(contract)
                if current_price:
                    action = self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        order = MarketOrder(action='SELL', totalQuantity=str(self.df['Qty'].iloc[i]))
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0  
                    else:
                        print("No profit/loss is triggered")

            elif self.df.loc[i,'Activation'] == 1 and self.df.loc[i,'Strike_Type'] == 'CE':

                datevar = self.df.loc[i, 'Expiry']
                print(datevar)
                datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                year,day,month = datevar_str.split('/')  
                formatted_date = f"{year}{month.zfill(2)}"

                contract      = Future(symbol='N225M',exchange='OSE.JPN',lastTradeDateOrContractMonth=str(formatted_date)) # change this line
                current_price = self.get_current_market_price_futures(contract)
                if current_price:
                    action = self.check_for_tp_sl(current_price, self.df['Target'].iloc[i],self.df['Stop_Loss'].iloc[i],self.df.loc[i,'Strike_Type'])
                    if action is not None:  
                        order = MarketOrder(action='BUY', totalQuantity=str(self.df['Qty'].iloc[i]))
                        result = self.client.placeOrder(contract, order)
                        print(result)
                        self.df.loc[i, 'Activation'] = 0  
                    else:
                        print("No profit/loss is triggered")

        with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.df.to_excel(writer, sheet_name="Sheet6", index=False)
    
    def new_auto_square_off(self): # put this in async
        self.df = pd.read_csv(self.excel_url)
        df = self.df
        current_time = datetime.now().strftime("%H:%M")
        positions = self.client.positions()
        if current_time == "10:20":
            if positions:
                for i in range(len(df)):
                    if self.df.loc[i,'Activation'] == 1:
                        datevar = self.df.loc[i, 'Expiry']
                        print(datevar)
                        datevar_str = datevar.strftime('%Y-%m-%d') if isinstance(datevar, pd.Timestamp) else str(datevar)
                        day,month,year = datevar_str.split('/')  # Ensure the date is in 'YYYY-MM-DD HH:MM:SS' format
                        formatted_date = f"{year}{month.zfill(2)}"
                        contract = Future(symbol='N225M', exchange='OSE.JPN', lastTradeDateOrContractMonth=str(formatted_date))
                        if self.df.loc[i, 'Strike_Type'] == 'CE':
                            current_action = 'BUY'
                        else:
                            current_action = 'SELL'
                        order = MarketOrder(action=current_action, totalQuantity=self.df.loc[i, 'Qty'])
                        result = self.client.placeOrder(contract, order)
                        self.df.loc[i, 'Activation'] = 0
            else:
                print("Positions are empty")
        else:
            print("The time is not for closing the market is not yet")
            print(positions)

        with pd.ExcelWriter(self.path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.df.to_excel(writer, sheet_name="Sheet6", index=False)

    def run(self):
        self.connection_show()
        while True:
            self.check_for_new_positions() # working fine 
            # self.new_auto_square_off() # working fine 
            # result = self.show_details() # working fine
            # print(result) # working
            # self.monitor_tp_sl() # working fine
            time.sleep(5)

def main():
    session = IBRKExcel()
    session.run()

if __name__ == "__main__":
    main()

# fix async the rest is good 

