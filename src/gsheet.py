import pandas as pd
import gspread
import threading
import sys
import numpy as np
import time

from google.oauth2.service_account import Credentials
import creds

import logging
logger = logging.getLogger(__name__)

class GSheet:
    def __init__(self):
        self.SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        self.SheetID = creds.sheet_id
        self.status = "Offline"
        self.AllTokens = None
        self.CSV_FILE = "alltokens.csv"
        self.CREDS_PATH = "credsGoogleSheets.json"
        self.workSheetNum = 0
        self.flag = False
        self.creds = None
        self.client = None
        self.spreadsheet = None
        self.worksheet = None
        self.exchange_mapping = {}
        self.lock = threading.Lock()
        self.ExhangeMappingLock = threading.Lock()
        self.newData = {}
        self.gSheetUpdated = False
        self.previous_data = []
        self.verification = np.zeros(1500, dtype=int)
        self.condition = threading.Condition()
        self.tokenName = [[""] for _ in range(1500)]
        self.NumStocks = 0

    def read_csv_file(self):
        logger.info(f"Reading CSV file ({self.CSV_FILE})...")
        try:
            self.AllTokens = pd.read_csv(self.CSV_FILE, low_memory=False)
            logger.info("CSV loaded successfully.")
        except Exception as e:
            logger.error("Error occurred in reading CSV.")
            logger.error(f"Error details: {e}")
            raise e

    def get_creds(self):
        logger.info(f"Reading {self.CREDS_PATH}...")
        try:
            self.creds = Credentials.from_service_account_file(self.CREDS_PATH, scopes=self.SCOPES)
            self.client = gspread.authorize(self.creds)
            logger.info("Read credentials file successfully.")
        except Exception as e:
            logger.error(f"Error reading credentials: {e}")
            self.terminate()

    def setUpGSheet(self):
        self.get_creds()
        try:
            self.spreadsheet = self.client.open_by_key(self.SheetID)
            self.worksheet = self.spreadsheet.get_worksheet(self.workSheetNum)
            self.changeStatus(code = 2)
        except Exception as e:
            logger.error(f"Error setting up Google Sheet: {e}")
            self.terminate()

    def diffAlgo(self, d1, d2, reverse=False):
        changedData = set(d1) - set(d2)
        for data in changedData:
            if reverse:
                self.verification[d1.index(data)] = 0
            else:
                condition = (self.AllTokens['exch_seg'] == data[1]) & (self.AllTokens['token'] == data[0])
                matched_symbol = self.AllTokens.loc[condition,'symbol'].tolist()
                if condition.any():
                    self.verification[d1.index(data)] = 1
                    self.tokenName[d1.index(data)][0] = matched_symbol[0]
                else:
                    self.verification[d1.index(data)] = 0
                    self.tokenName[d1.index(data)][0] = "Invalid"


    def compare_data(self, current_data):
        if current_data != self.previous_data:
            logger.info("Data has changed in Google Sheet.")
            self.diffAlgo(self.previous_data, current_data, reverse=True)
            self.diffAlgo(current_data, self.previous_data)
            self.previous_data = current_data
            with self.lock:
                self.gSheetUpdated = True
            return True
        return False

    def get_sheet_data(self):
        with self.lock:
            try:
                data = self.worksheet.get_all_values()
                columns_ab = []
                self.newData = {}
                for row in data:
                    columns_ab.append((row[0], row[1]))
                    if self.verification[data.index(row)] == 1:
                        self.newData.setdefault(row[1], []).append(row[0])
                return columns_ab
            except Exception as e:
                logger.error(f"Error fetching sheet data: {e}")
                self.terminate()

    def monitor_changes(self):
        while True:
            try:
                self.current_data = self.get_sheet_data()
                if self.compare_data(self.current_data):
                    logger.info("Change detected in the sheet column.")
                    self.trigger_condition()
                time.sleep(5)
            except Exception as e:
                logger.error(f"Error monitoring changes: {e}")
                break
    
    def google_sheet_changes_monitor(self):
        while True:
            try:
                self.current_data = self.get_sheet_data()
                if self.compare_data(self.current_data):
                    return True
                else:
                    return False
                
            except Exception as e:
                print(f"Error in google_sheet changes")

    def terminate(self):
        logger.critical("Terminating the program.")
        sys.exit("Terminating the program.")

    def updateData(self, ohlcv):

        try:
            cell_range = f"D2:H{len(ohlcv) + 1}"
            self.worksheet.update(cell_range, ohlcv)
            logger.info(f"Updated Google Sheet with {len(ohlcv)} rows.")
        except Exception as e:
            logger.error(f"Error updating Google Sheet: {e}")

    def changeStatus(self,code = 0):
        if code == 0:
            self.status = "Running"
        elif code == 1:
            self.status = "Invalid Token"
        elif code == 2:
            self.status = "Reading Tokens"
        elif code == 3:
            self.status = "Offline"
        elif code == 4:
            self.status = "No Tick"
        try:
            self.worksheet.update("J6", [[self.status]])
        except Exception as e:
            logger.error(f"Error changing sheet status: {e}")

    def signal_function(self):
        with self.condition:
            logger.info("Waiting for the condition to be met...")
            while not self.flag:
                self.condition.wait()
            time.sleep(10)
            
            logger.info("Condition met! Executing signal function.")
    def updateSymbol(self):
        if len(self.tokenName) > 1:
            empty_index = self.tokenName.index([""])
            valid_range = self.tokenName[1:empty_index]
            self.worksheet.update(f"C2:{empty_index}",valid_range)

    def trigger_condition(self):
        with self.condition:
            self.flag = True
            self.condition.notify_all()
            logger.info("Condition triggered.")

    def main(self):
        try:
            logger.info("Starting CSV reading thread...")
            csv_thread = threading.Thread(target=self.read_csv_file)
            csv_thread.start()

            self.setUpGSheet()
            csv_thread.join()
            logger.info("CSV reading thread completed.")

            self.monitor_changes()
        except KeyboardInterrupt:
            logger.info("Closing program due to KeyboardInterrupt.")
            self.terminate()
        except Exception as e:
            logger.error(f"Error in main: {e}")

    def updateByRow(self,token,ohlcv):
        index = next((i for i, (key, _) in enumerate(self.previous_data) if key == token), -1)
        if index == -1:
            logger.error("Not able to get Token")
            return
        
        cell_range = f"D{index+1}:H{index+1}"
        self.worksheet.update(cell_range, [ohlcv])
        logger.info("Data Updated")
# Usage
# sheet = GSheet()
# sheet.main()