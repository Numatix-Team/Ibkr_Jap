import time
import asyncio
import pandas as pd
from datetime import datetime,time
from ib_broker import *
import credentials

class IBRKExcel:
    def __init__(self):
        self.path            = credentials.xlsx_path
        self.current_time    = datetime.now()
        self.excel_data      = pd.read_excel(credentials.xlsx_path, sheet_name='Sheet6')
        self.length          = len(self.excel_data)
        self.orderbook       = []
        self.failed_orders   = []
        self.database_path   = credentials.database_path
        # self.square_off_time = datetime.time(11, 45)  # Auto square-off time (11:45)
        self.current_time    = datetime.now().time()

    def check_excel_changes(self):
        new_data   = pd.read_excel(self.path, sheet_name='Sheet6')
        new_length = len(new_data)

        if new_length  != self.length:
            self.length     = new_length
            self.excel_data = new_data
            return True
        return False

    def connect(self) -> bool:
        host, port = credentials.host, credentials.port
        self.client = IB()
        self.ib = self.client
        self.client.connect(host=host, port=port, clientId=13, timeout=60)
        print("Connected")
    
    def check_for_new_positions(self):
        if self.check_excel_changes():
            last_row           = self.excel_data.iloc[-1]
            self.symbol        = 'N225M'
            self.local_symbol  = 'N225M_CONT'
            self.multiplier    = '1'
            self.exchange      = 'OSE.JPN'
            self.currency      = 'JPY'
            self.trigger_level = last_row['Trigger_Level_High_Low']
            self.entry_type    = last_row['Entry_Type']
            self.entry_strike  = last_row['Entry_Strike']
            self.strike_type   = last_row['Strike_Type']
            self.expiry        = last_row['Expiry']
            self.target        = last_row['Target']
            self.stop_loss     = last_row['Stop_Loss']
            self.qty           = last_row['Qty']
            self.slicing       = last_row['Slicing']
            self.time_interval = last_row['Time_Interval']
            self.activation    = last_row['Activation']
            if last_row['Trigger_Level_High_Low'] == 'High':
                self.side = 'BUY'
            else:
                self.side = 'SELL'

            if self.activation == 1:
                for _ in range(0,int(self.qty/self.slicing),1):
                    self.contract      = ContFuture(symbol=self.symbol,exchange=self.exchange,localSymbol=self.local_symbol,multiplier=self.multiplier,currency=self.currency)
                    # self.order         = LimitOrder(action=self.side,totalQuantity=int(self.qty/self.slicing),lmtPrice=self.entry_strike)
                    self.order         = MarketOrder(action=self.side,totalQuantity=int(self.qty/self.slicing))
                    self.order_details = self.client.placeOrder(contract=self.contract,order=self.order)
                    print(self.order_details)
                    print("The order has been placed")
                    time.sleep(self.time_interval)

    # def get_current_market_price_futures(self, contract):
    #     """
    #     Fetches the current market price of the given futures contract.
    #     """
    #     ticker = self.client.reqMktData(contract, '', False, False)
    #     self.client.sleep(1)  # Allow data to fetch
    #     if ticker.last is not None:
    #         return ticker.last
    #     elif ticker.close is not None:
    #         return ticker.close
    #     return None
    
    # def check_for_tp_sl(self, current_price, target_price):
    #     """
    #     Checks if the current price has hit the target price or stop loss.
    #     """
    #     if current_price >= 1.03 * target_price:  # Target Price hit
    #         return "SELL"
    #     elif current_price <= 0.97 * target_price:  # Stop Loss hit
    #         return "SELL"
    #     return None
    
    # def monitor_tp_sl(self):
    #     """
    #     Monitors TP/SL for active positions and places sell orders if hit.
    #     """
    #     df = pd.read_excel("path_to_your_excel_file.xlsx", sheet_name="Sheet6")  # Replace with your file path
    #     for i in range(len(df)):
    #         if df['Activation'].iloc[i] == 1:  # Active position
    #             contract = ContFuture(
    #                 symbol=df['Symbol'].iloc[i],
    #                 exchange=df['Exchange'].iloc[i],
    #                 localSymbol=df['LocalSymbol'].iloc[i],
    #                 multiplier='1',
    #                 currency=df['Currency'].iloc[i]
    #             )
    #             current_price = self.get_current_market_price_futures(contract)
    #             if current_price:
    #                 action = self.check_for_tp_sl(current_price, df['Target'].iloc[i])
    #                 if action:  # Place sell order if TP/SL hit
    #                     order = MarketOrder(action='SELL', totalQuantity=df['Qty'].iloc[i])
    #                     self.client.placeOrder(contract, order)
    #                     df.loc[i, 'Activation'] = 0  # Deactivate the position

    #     df.to_excel("path_to_your_excel_file.xlsx", sheet_name="Sheet6", index=False)  # Save changes

    def auto_square_off(self):
        """
        Square off all open positions at the specified square-off time.
        """
        positions = self.client.positions()
        current_time = datetime.now().strftime("%H:%M")
        if current_time == "11:44": # timing part is working fine 
            if positions:
                for pos in positions:
                    contract = pos.contract
                    order = MarketOrder(action='SELL', totalQuantity=pos.position)
                    self.client.placeOrder(contract, order)
                print("All open positions squared off.")
            else:
                print("Positions are empty")
        else:
            print("The time is not for closing the market is not yet")

    def run(self):
        self.connect()
        while True:
            # run the three of them in async
            # self.check_for_new_positions() # working fine 
            self.auto_square_off() # check at around 8 in the morning
            # self.monitor_tp_sl()
            time.sleep(5)

def main():
    session = IBRKExcel()
    session.run()


if __name__ == "__main__":
    main()

# working code later revision