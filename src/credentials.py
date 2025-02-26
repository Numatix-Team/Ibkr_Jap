port = 7497
instrument = "SPX"
exchange = "CBOE"
currency = "USD"
strike = 6000
deviation = 10
date = "20241227"
host = "127.0.0.1"
data_type = 4
number_of_re_entry = 2
hedge_difference = 10
call_sl = 15
put_sl = 15
# xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\Ibkr.xlsx"
xlsx_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\Ibkr.xlsx"
database_path = r"C:\Users\vaibh\OneDrive\Desktop\New folder\Folder Python\Folder Algotrading\ibkr_excel\storage\orderbook.csv"
current_time = r"21:35"
attempts = 3
master = True


"""
user_time_default 0 - use the time_interval from the xlsx file
user_time_default 1 - use user_time from creds file
trade_type_default 0 - set the limitprice to the current price
trade_type_default 1 - set the limitprice based on the bid and ask
"""
user_time_default = 1
user_time = 3
trade_type_default = 0