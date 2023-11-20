import requests
from bs4 import BeautifulSoup
import pandas as pd

# URL and headers
url = "your_url_here"
headers = {'User-Agent': 'Mozilla/5.0'}

# Fetch and parse the page
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.content, 'html.parser')

# Find the table
table = soup.find('table', class_='b1')

# Lists to hold your data
closing_prices = []
bps_values = []

# Assuming the first row is headers and data starts from the second row
for row in table.find_all('tr')[1:]:
    cells = row.find_all('td')
    # You need to determine the correct indices for your data
    closing_price = cells[your_closing_price_index].text
    bps_value = cells[your_bps_value_index].text

    closing_prices.append(closing_price)
    bps_values.append(bps_value)

# Creating a DataFrame
df = pd.DataFrame({
    '收盤價(Closing price)': closing_prices,
    '河流圖BPS(元)': bps_values
})

# Saving to an Excel file
df.to_excel('stock_data.xlsx', index=False)

print('Data extraction complete. Check the stock_data.xlsx file.')
