import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

class GSheet:
    def __init__(self):
        self.SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        # self.SheetID = "your_google_sheet_id"
        self.SheetID = "16bpyg4FMgbd22_9FwOCrCqmDqg1Y7Ulff38w-xX1kXM"
        self.CREDS_PATH = "credsGoogleSheets.json"
        self.workSheetNum = 0  # Index of the worksheet
        self.client = None
        self.worksheet = None

    def get_creds(self):
        creds = Credentials.from_service_account_file(self.CREDS_PATH, scopes=self.SCOPES)
        self.client = gspread.authorize(creds)
        self.worksheet = self.client.open_by_key(self.SheetID).get_worksheet(self.workSheetNum)

    def update_entire_dataframe(self, df):
        if df.empty:
            print("DataFrame is empty. Nothing to update.")
            return

        # Convert DataFrame to list of lists
        data = [df.columns.tolist()] + df.values.tolist()  # Add headers

        try:
            # Clear the existing sheet before updating
            self.worksheet.clear()
            # Update the sheet with new data
            self.worksheet.update("A1", data)
            print(f"Updated {len(df)} rows successfully.")
        except Exception as e:
            print(f"Error updating Google Sheet: {e}")
