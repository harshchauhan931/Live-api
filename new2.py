import pandas as pd
import requests
from datetime import datetime
import schedule
import time

from new1 import update_excel

# API constants
API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": "false",
}

# File path in a cloud-synced folder (adjust this to your environment)
FILE_PATH = "D:\onedrive ug.sharda.ac.in\OneDrive - ug.sharda.ac.in\Crypto_Live_Data.xlsx"

# Function to fetch crypto data
def fetch_crypto_data():
    response = requests.get(API_URL, params=PARAMS)
    if response.status_code == 200:
        data = response.json()
        crypto_list = []
        for item in data:
            crypto_list.append({
                "Name": item["name"],
                "Symbol": item["symbol"].upper(),
                "Current Price (USD)": item["current_price"],
                "Market Cap (USD)": item["market_cap"],
                "24h Volume (USD)": item["total_volume"],
                "24h Price Change (%)": item["price_change_percentage_24h"],
            })
        return pd.DataFrame(crypto_list)
    else:
        print(f"Error fetching data: {response.status_code}")
        return pd.DataFrame()

# Function to save data to Excel
def save_to_excel(df):
    # Add a timestamp column
    df["Last Updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Write the DataFrame to an Excel file
    with pd.ExcelWriter(FILE_PATH, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="Crypto Data", index=False)
    print(f"Excel file updated at {datetime.now()}")

# Task to fetch data and save to Excel
def update_excel_file():
    print("Fetching cryptocurrency data...")
    crypto_data = fetch_crypto_data()
    if not crypto_data.empty:
        save_to_excel(crypto_data)
    else:
        print("No data to save. Skipping update.")

# Schedule the update task
schedule.every(0.10).minutes.do(update_excel_file)

# Run the scheduled tasks
if __name__ == "__main__":
    print("Starting live data update...")
    update_excel_file()  # Initial run
    while True:
        schedule.run_pending()
        time.sleep(1)


# Run the script
print("Starting live updates. Press Ctrl+C to stop.")
update_excel()  # Initial run
while True:
    schedule.run_pending()
    time.sleep(1)
def save_to_open_excel(df, file_name):
    try:
        wb = load_workbook(file_name)
        ws = wb.active

        # Clear existing data
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                cell.value = None

        # Write new data
        for i, row in df.iterrows():
            ws.append(row.tolist())

        wb.save(file_name)
        print(f"Data updated in {file_name}")
    except Exception as e:
        print(f"Error updating Excel file: {e}")