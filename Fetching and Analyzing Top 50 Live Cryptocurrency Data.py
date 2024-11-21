# 1- Fetch Live Data
import requests
from time import sleep

def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }

    retries = 5  # Number of retries
    for i in range(retries):
        try:
            response = requests.get(url, params=params)
            response.raise_for_status()  # Check for HTTP errors
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data (attempt {i+1}/{retries}): {e}")
            sleep(2)  # Wait 2 seconds before retrying
    return []

crypto_data = fetch_crypto_data()
if crypto_data:
    print(f"Fetched {len(crypto_data)} cryptocurrencies.")
else:
    print("Failed to fetch cryptocurrency data.")

import pandas as pd

def process_crypto_data(crypto_data):
    if not crypto_data:
        return None
    df = pd.DataFrame(crypto_data)
    df = df[[
        "name",
        "symbol",
        "current_price",
        "market_cap",
        "total_volume",
        "price_change_percentage_24h"
    ]]
    df.columns = [
        "Cryptocurrency Name",
        "Symbol",
        "Current Price (USD)",
        "Market Capitalization",
        "24h Trading Volume",
        "24h Price Change (%)"
    ]
    return df

crypto_df = process_crypto_data(crypto_data)
if crypto_df is not None:
    print(crypto_df.head())

#2 Data Analysis

import requests
import pandas as pd
from time import sleep

def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }

    retries = 5  # Number of retries
    for i in range(retries):
        try:
            response = requests.get(url, params=params)
            response.raise_for_status()  # Check for HTTP errors
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data (attempt {i+1}/{retries}): {e}")
            sleep(2)  # Wait 2 seconds before retrying
    return []


def process_crypto_data(crypto_data):
    if not crypto_data:
        return None
    df = pd.DataFrame(crypto_data)
    df = df[[
        "name",
        "symbol",
        "current_price",
        "market_cap",
        "total_volume",
        "price_change_percentage_24h"
    ]]
    df.columns = [
        "Cryptocurrency Name",
        "Symbol",
        "Current Price (USD)",
        "Market Capitalization",
        "24h Trading Volume",
        "24h Price Change (%)"
    ]
    return df

def analyze_data(df):
    if df is not None:
        # Identifying the top 5 cryptocurrencies by market capitalization
        top_5_by_market_cap = df.nlargest(5, 'Market Capitalization')

        # Calculating the average price of the top 50 cryptocurrencies
        average_price = df['Current Price (USD)'].mean()

        # Analyzing the highest and lowest 24-hour price change
        highest_price_change = df.loc[df['24h Price Change (%)'].idxmax()]
        lowest_price_change = df.loc[df['24h Price Change (%)'].idxmin()]

        print("Top 5 Cryptocurrencies by Market Cap:")
        print(top_5_by_market_cap[['Cryptocurrency Name', 'Market Capitalization']])

        print("\nAverage Price of Top 50 Cryptocurrencies: $", round(average_price, 2))

        print("\nHighest 24h Price Change:")
        print(highest_price_change[['Cryptocurrency Name', '24h Price Change (%)']])

        print("\nLowest 24h Price Change:")
        print(lowest_price_change[['Cryptocurrency Name', '24h Price Change (%)']])

# Main execution
crypto_data = fetch_crypto_data()
crypto_df = process_crypto_data(crypto_data)
analyze_data(crypto_df)

# 3 Live Running Excel Sheet
import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from time import sleep

# Function to fetch cryptocurrency data from API
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",            # Fetch data in USD
        "order": "market_cap_desc",      # Order by market cap
        "per_page": 50,                  # Get top 50 cryptocurrencies
        "page": 1,                       # Page 1
        "sparkline": False               # Exclude sparkline data
    }

    retries = 5  # Number of retries
    for i in range(retries):
        try:
            response = requests.get(url, params=params)
            response.raise_for_status()  # Check for HTTP errors
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data (attempt {i+1}/{retries}): {e}")
            sleep(2)  # Wait 2 seconds before retrying
    return []  # Return an empty list if failed after retries

# Function to process the data into a DataFrame
def process_crypto_data(crypto_data):
    if not crypto_data:
        return None
    df = pd.DataFrame(crypto_data)
    df = df[[
        "name",
        "symbol",
        "current_price",
        "market_cap",
        "total_volume",
        "price_change_percentage_24h"
    ]]
    df.columns = [
        "Cryptocurrency Name",
        "Symbol",
        "Current Price (USD)",
        "Market Capitalization",
        "24h Trading Volume",
        "24h Price Change (%)"
    ]
    return df

# Function to create or update Excel file with new data
def update_excel(df):
    wb = Workbook()  # Create a new workbook (you can modify to append to an existing one if needed)
    ws = wb.active
    ws.title = "Cryptocurrency Data"

    # Add headers
    ws.append(df.columns.tolist())

    # Add data rows
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    # Save the workbook
    wb.save("crypto_data.xlsx")

# Function to update the data every 5 minutes
def live_update():
    while True:
        print("Fetching data...")
        crypto_data = fetch_crypto_data()
        crypto_df = process_crypto_data(crypto_data)
        
        if crypto_df is not None:
            print("Updating Excel sheet...")
            update_excel(crypto_df)
            print("Data updated in Excel.")
        else:
            print("Failed to fetch data.")

        print("Waiting for next update...")
        sleep(300)  # Wait for 5 minutes (300 seconds)

# Main function to start the live update
if __name__ == "__main__":
    live_update()
