"""ports host and account"""
port = 7497
host = "127.0.0.1"
# account_no = "DU9727656"
account_no = "DU7166729"

"""symbol and paths"""
symbol = "N225M"
exchange = "OSE.JPN"
# xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\new_order_sheet.xlsx"
"""sheet path and sheet"""
xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\final_2.xlsx"
sheet_name = "Sheet1"

"""parameters for trading"""
pause_time = 5 # set this as default - time before 
attempts = 3 # attempts 
closing_time = r"23:00"
"""

attempts - the number of times the bot would place the orders ex - 3
current_time - the time at which the trading bot would stop ex - "22:00" # closes at 10PM
xlsx_path - path of the xlsx file (would also need to specify the sheet like Sheet1,Sheet6)
port,host - used in login
"""

"""
when you set the activation to 1 then normal orders for buy/selling for the futures are made once you save them then order gets placed and the activation becomes -1
when you set the activation to 3 then goodtilldate type orders for buy/selling for the futures are made once you save them then order gets placed and the activation becomes -3
when you don't set the take profit and stop loss and fill them with '-' and '-' respectively then based on 'buy' we will set the upper_trigger to the entry price of that column / if 'sell' then we will set the lower_trigger to the entry price of that column
in case of multiple columns have -,- then the first one is considered this all only works when the activation is manually set to 2 otherwise if the activation is set to 0 then the column is ignored
"""
