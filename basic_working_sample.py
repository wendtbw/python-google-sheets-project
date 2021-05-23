from __future__ import print_function
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import yfinance as yf
​
# defining scope, the spreadsheet to write into, and the range within the spreadsheet
# the spreadsheetID is the long string you find between the /d/ and /edit
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1PE8gBnr4bBujov502OUAfLpMFFZdT-yS5D042tEGRkQ'
SAMPLE_RANGE_NAME = 'Sheet2!B3'
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
       # https://developers.google.com/sheets/api/quickstart/python/
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
    ticker_list = ["6758.T", "9983.T", "9432.T", "6098.T", "7974.T", "8035.T", "9434.T", "7267.T", "9984.T", "7203.T", "4755.T"]
#    ticker_list = ["TSLA", "AAPL", "LB", "EXPE", "PXD", "CRM", "GOOG", "ORCL", "4755.T", "9984.T", "7203.T"]
​
    metrics = ['previousClose', 'currency', 'shortName', 'twoHundredDayAverage', 'trailingPE',
                       'trailingAnnualDividendYield', 'payoutRatio', 'priceToSalesTrailing12Months']
​
    data_to_write = []
    metric_values = []
​
    for t in ticker_list:                                   # for each ticker in the ticker list....
        mydict = yf.Ticker(t).info                          # get a dictionary of numerous metric values from yahoo finance
        metric_values.append(t)                             # add the ticker as the first column in metric_values for ticker t
        for metric in metrics:                              # for each metric of ticker t we want to write into the google sheet....
            metric_values.append(mydict.get(metric))        # append the value of the metric to metric_values
        print(metric_values)
        data_to_write.append(metric_values)                 # take the list of metric values we obtained for ticker t, and append it
                                                            # to data_to_write which is a list of lists (like a 2 dimensional array)
                                                            # that we will eventually write into the google sheet below
        metric_values = []                                  # reset metric_values for the next ticker
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
