from __future__ import print_function
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import yfinance as yf
​
# defining scope, the spreadsheet to write into, and the range within the spreadsheet
# the spreadsheetID is the long string you find between the /d/ and /edit
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1PE8gBnr4bBujov502OUAfLpMFFZdT-yS5D042tEGRkQ'
SAMPLE_RANGE_NAME = 'Sheet2!B2'
​
def main():
    """This is a fairly simple working example to look up stock prices
       of a list of stocks (stored in ticker_list) and write the ticker
       and its previous closing price into a google spreadsheet.
​
       The previous closing prices are obtained using Yahoo Finance API.
       Writing into google sheets is done using Google Sheets API.
​
       The first time running this, I had more complex code to
       establish authorization.  Code came from Google Sheets API pages.
       Ive since removed said code now that the token
       file (token.json) has been created.
    """
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
​
    service = build('sheets', 'v4', credentials=creds)
​
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
​
​
    ticker_list = ["TSLA", "AAPL", "LB", "EXPE", "PXD", "MCHP", "CRM", "GOOG", "NRG", "ORCL", "NOW"]
​
    data_to_write = []
​
    for t in ticker_list:
        mydict = yf.Ticker(t).info
        print()
        previous_close = mydict.get('previousClose')
        print(t, previous_close)
        data_to_write.append([t, previous_close])
​
    print(data_to_write)
​
    request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME,
                                                     valueInputOption="USER_ENTERED", body={"values":data_to_write})
    response = request.execute()
​
if __name__ == '__main__':
    main()
​
# reference
# https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/update
# https://codingandfun.com/financial-data-from-yahoo-finance/
# https://www.youtube.com/watch?v=4ssigWmExak
