import requests
from bs4 import BeautifulSoup
import pandas as pd

from selenium import webdriver
from selenium.webdriver.edge.service import Service
import time

# Set up the Selenium driver (update the path to where you placed msedgedriver.exe)
service = Service("C:\\Users\\rayra\\Code\\Web-Crawler\\msedgedriver.exe")

driver = webdriver.Edge(service=service)

# The rest of your script follows...






# URL and headers
url = "https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=PBR&STOCK_ID=2330&CHT_CAT=WEEK"
headers = {'User-Agent': 'Mozilla/5.0'}

# Fetch and parse the page
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.content, 'html.parser')

# Find all tables and print their count for debugging
tables = soup.find_all('table')
print(f"Number of tables found: {len(tables)}")

# Assuming we need the first table (modify this based on actual needs)
if tables:
    table = tables[0]  # Adjust the index if needed

    # Lists to hold your data
    closing_prices = []
    bps_values = []

    # Assuming the first row is headers and data starts from the second row
    for row in table.find_all('tr')[1:]:
        cells = row.find_all('td')

        # Use try-except to avoid IndexError if a row has unexpected number of cells
        try:
            closing_price = cells[1].text.strip()  # Index 1 for the second column
            bps_value = cells[4].text.strip()      # Index 4 for the fifth column
            closing_prices.append(closing_price)
            bps_values.append(bps_value)
        except IndexError:
            print(f"Skipping a row with unexpected format: {row}")

    # Creating a DataFrame
    df = pd.DataFrame({
        '收盤價(Closing price)': closing_prices,
        '河流圖BPS(元)': bps_values
    })

    # Saving to an Excel file
    df.to_excel('stock_data.xlsx', index=False)

    print('Data extraction complete. Check the stock_data.xlsx file.')
else:
    print("No tables found on the page.")
