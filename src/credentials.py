port = 7497
host = "127.0.0.1"
# xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\Ibkr.xlsx"
xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\order_sheet.xlsx"
current_time = r"20:35"
attempts = 3
master = True
user_time_default = 1
user_time = 5
trade_type_default = 1

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