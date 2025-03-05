import requests
import pandas as pd
import openpyxl
import time
import schedule

API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False
}


def fetch_crypto_data():
    response = requests.get(API_URL, params=PARAMS)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data:", response.status_code)
        return None


def process_data(data):
    if not data:
        return None
    
    crypto_list = []
    for coin in data:
        crypto_list.append([
            coin['name'],
            coin['symbol'].upper(),
            coin['current_price'],
            coin['market_cap'],
            coin['total_volume'],
            coin['price_change_percentage_24h']
        ])

    df = pd.DataFrame(crypto_list, columns=[
        "Cryptocurrency Name", "Symbol", "Current Price (USD)", 
        "Market Capitalization", "24h Trading Volume", "24h Price Change (%)"
    ])
    
    return df


def analyze_data(df):
    if df is None:
        print("No data available for analysis.")
        return None

    top_5 = df.nlargest(5, "Market Capitalization")
    avg_price = df["Current Price (USD)"].mean()
    highest_change = df.loc[df["24h Price Change (%)"].idxmax()]
    lowest_change = df.loc[df["24h Price Change (%)"].idxmin()]

    print("\n🔹 Top 5 Cryptocurrencies by Market Cap:\n", top_5)
    print("\n💰 Average Price of Top 50 Cryptocurrencies: $", round(avg_price, 2))
    print("\n📈 Highest 24h Price Change:", highest_change["Cryptocurrency Name"], "-", highest_change["24h Price Change (%)"], "%")
    print("\n📉 Lowest 24h Price Change:", lowest_change["Cryptocurrency Name"], "-", lowest_change["24h Price Change (%)"], "%")

    return top_5, avg_price, highest_change, lowest_change


def update_excel(df):
    file_name = "Live_Crypto_Data.xlsx"
    df.to_excel(file_name, index=False, engine='openpyxl')
    print("✅ Excel file updated successfully!")


def run_script():
    print("\n⏳ Fetching latest cryptocurrency data...")
    data = fetch_crypto_data()
    df = process_data(data)

    if df is not None:
        update_excel(df)
        analyze_data(df)


schedule.every(5).minutes.do(run_script)


run_script()


while True:
    schedule.run_pending()
    time.sleep(1)
