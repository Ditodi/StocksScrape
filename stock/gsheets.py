import time
import gspread
from yahoo_fin import stock_info as si
from oauth2client.service_account import ServiceAccountCredentials

price_col = 4
price_row_init = 2

print("Connecting to Google Server...")

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_service.json', scope)

client = gspread.authorize(creds)

sheet = client.open('Stocks').sheet1
sheet.sort((1, 'asc'))

print('Getting Stock Symbols from Google Sheets...')

ticker = sheet.col_values(1)
ticker.pop(0)

print("Updating Live Price from Market...")
print('Please Do not Edit the Following Sheets (Sheet1)')

for item in ticker:
    # print(item)

    # Obtain Share Outsanding
    try:
        data = si.get_quote_data(item)
        shareout = data["sharesOutstanding"]
        sheet.update_cell(price_row_init, 2, int(shareout))
    except:
        sheet.update_cell(price_row_init, 2, "NaN")
        pass

    # Obtain Retained Earnings
    try:
        balsheet = si.get_balance_sheet(item, yearly=False)
        print(balsheet)
        # retained = balsheet.loc['retainedEarnings', balsheet.index[1]]
        retained = balsheet.loc['retainedEarnings'].iloc[0]
        print(retained)
        sheet.update_cell(price_row_init, 3, int(retained))
    except:
        sheet.update_cell(price_row_init, 3, "NaN")
        pass

    # Obtain stocks live price and PER
    try:
        price = si.get_live_price(item)
        quote = si.get_quote_table(item)
        sheet.update_cell(price_row_init, price_col, price)
        sheet.update_cell(price_row_init, 9, quote["PE Ratio (TTM)"])
    except:
        sheet.update_cell(price_row_init, 9, "NaN")
        pass

    # Obtain PBVR
    try:
        val = si.get_stats_valuation(item)
        print(val)
        pricebook = val.iloc[6, 1]
        sheet.update_cell(price_row_init, 10, pricebook)
    except:
        sheet.update_cell(price_row_init, 10, "NaN")
        pass

    # Obtain ROE
    try:
        stats = si.get_stats(item)
        roe = stats.iloc[34, 1]
        sheet.update_cell(price_row_init, 11, roe)
    except:
        sheet.update_cell(price_row_init, 11, "NaN")
        pass
    time.sleep(0.5)
    price_row_init += 1
print("Data Scraping Completed..")
