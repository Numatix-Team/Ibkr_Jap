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
        self.ib_ved = IBTWSAPI(creds=credentials)

    def check_excel_changes(self):
        new_data   = pd.read_excel(self.path, sheet_name='Sheet6')
        new_length = len(new_data)

        if new_length  != self.length:
            self.length     = new_length
            self.excel_data = new_data
            return True
        return False
    
    def check_for_new_positions(self):
        if self.check_excel_changes():
            last_row           = self.excel_data.iloc[-1]
            # 'N225M','N225M_CONT','1','OSE.JPN','JPY'
            self.symbol        = 'N225M'
            self.local_symbol  = 'N225M_CONT'
            self.multiplier    = '100' 
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
                    # self.contract = self.ib_ved._create_contract(contract='futureContracts',symbol=self.symbol,exchange=self.exchange)
                    # self.contract = self.ib_ved._create_contract(contract='future',symbol='N225M',exchange='OSE.JPN',currency='JPY')
                    self.contract = self.ib_ved._create_contract(contract='future',symbol='N225M',ltdocm='202503',exchange='OSE.JPN')
                    self.order_details = self.ib_ved.place_order(contract=self.contract,symbol=self.symbol,side=self.side,quantity=int(self.qty/self.slicing),order_type="MARKET",price=self.entry_strike,exchange=self.exchange)
                    print(self.order_details)
                    print("The order has been placed")
                    time.sleep(self.time_interval)

    def run(self):
        self.ib_ved.connect()
        while True:
            # run the three of them in async
            self.check_for_new_positions() # working fine 
            # self.auto_square_off() # check at around 8 in the morning
            # self.monitor_tp_sl()
            time.sleep(5)

def main():
    session = IBRKExcel()
    session.run()

if __name__ == "__main__":
    main()
