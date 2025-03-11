port = 7497
host = "127.0.0.1"
# xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\Ibkr.xlsx"
# xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\order_sheet.xlsx"
xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\new_order_sheet.xlsx"
xlsx_path_1 = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\options_order_sheet.xlsx"
sheet_name = "Sheet1"
current_time = r"22:13"
pause_time = 5 # set this as default
attempts = 3
master = True
user_time_default = 1
user_time = 5
trade_type_default = 0

"""
user_time_default 0 - use the time_interval from the xlsx file
user_time_default 1 - use user_time from creds file
trade_type_default 0 - set the limitprice to the current price
trade_type_default 1 - set the limitprice based on the bid and ask
attempts - the number of times the bot would place the orders ex - 3
current_time - the time at which the trading bot would stop ex - "22:00" # closes at 10PM
master - used to turn the bot on / off (True/False)
xlsx_path - path of the xlsx file (would also need to specify the sheet like Sheet1,Sheet6)
port,host - used in login
"""

"""
when you set the activation to 1 then normal orders for buy/selling for the futures are made once you save them then order gets placed and the activation becomes -1
when you set the activation to 3 then goodtilldate type orders for buy/selling for the futures are made once you save them then order gets placed and the activation becomes -3
when you don't set the take profit and stop loss and fill them with '-' and '-' respectively then based on 'buy' we will set the upper_trigger to the entry price of that column / if 'sell' then we will set the lower_trigger to the entry price of that column
in case of multiple columns have -,- then the first one is considered this all only works when the activation is manually set to 2 otherwise if the activation is set to 0 then the column is ignored
"""
