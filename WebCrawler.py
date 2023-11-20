import requests
from bs4 import BeautifulSoup
import pandas as pd

# URL of the page
url = "https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=PBR&STOCK_ID=2330&CHT_CAT=WEEK"

# Headers to simulate a browser visit
headers = {'User-Agent': 'Mozilla/5.0'}

# Send a GET request to the URL
response = requests.get(url, headers=headers)
response.raise_for_status()  # Ensure the request was successful

# Parse the HTML content
soup = BeautifulSoup(response.content, 'html.parser')

# Locate the table and rows - you need to adjust this based on the actual page structure
table = soup.find(...)  # Use the correct identifier for the table
rows = table.find_all('tr') if table else []

# Data extraction
data = []
for row in rows:
    cols = row.find_all('td')
    # Extract data from the specific columns - adjust the indices as per the table's structure
    closing_price = cols[index_of_closing_price].get_text(strip=True)
    bps = cols[index_of_bps].get_text(strip=True)
    data.append([closing_price, bps])

# Convert to DataFrame
df = pd.DataFrame(data, columns=['收盤價(Closing price)', '河流圖BPS(元)'])

# Write to Excel
df.to_excel('stock_data.xlsx', index=False)

print('Data extraction complete. Check the stock_data.xlsx file.')
