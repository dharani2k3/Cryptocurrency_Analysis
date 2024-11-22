import requests
import pandas as pd

# URL to get top 50 cryptocurrencies by market cap
url = 'https://api.coingecko.com/api/v3/coins/markets'
params = {
    'vs_currency': 'usd',  # Data in USD
    'order': 'market_cap_desc',  # Sort by market cap descending
    'per_page': 50,  # Get the top 50 coins
    'page': 1,  # Fetch from the first page
    'sparkline': 'false'  # No sparkline data
}

# Send the request and get the response
response = requests.get(url, params=params)
data = response.json()  # This converts the response into a JSON object

# Extract and organize relevant data into a list of dictionaries
cryptos = []
for coin in data:
    cryptos.append({
        'Name': coin['name'],
        'Symbol': coin['symbol'],
        'Price (USD)': coin['current_price'],
        'Market Cap': coin['market_cap'],
        '24h Trading Volume': coin['total_volume'],
        '24h Price Change (%)': coin['price_change_percentage_24h']
    })

# Create a pandas DataFrame from the list of dictionaries
df = pd.DataFrame(cryptos)

# Show the DataFrame in the console (Optional)
print(df)

# Save the data to an Excel file
df.to_excel("cryptocurrency_data.xlsx", index=False)

top_5 = df.nlargest(5, 'Market Cap')  # Select top 5 by market cap
print("Top 5 Cryptocurrencies by Market Cap:")
print(top_5[['Name', 'Market Cap']])

average_price = df['Price (USD)'].mean()
print(f"Average Price of the Top 50 Cryptocurrencies: ${average_price:.2f}")

max_change = df['24h Price Change (%)'].max()  # Highest 24-hour price change
min_change = df['24h Price Change (%)'].min()  # Lowest 24-hour price change

# Get the corresponding rows with the highest and lowest changes
highest_change = df[df['24h Price Change (%)'] == max_change]
lowest_change = df[df['24h Price Change (%)'] == min_change]

print(f"Highest Price Change: {highest_change[['Name', '24h Price Change (%)']]}")
print(f"Lowest Price Change: {lowest_change[['Name', '24h Price Change (%)']]}")

import time

while True:
    # Re-fetch data from the API
    response = requests.get(url, params=params)
    data = response.json()

    # Process the data and update the DataFrame
    cryptos = []
    for coin in data:
        cryptos.append({
            'Name': coin['name'],
            'Symbol': coin['symbol'],
            'Price (USD)': coin['current_price'],
            'Market Cap': coin['market_cap'],
            '24h Trading Volume': coin['total_volume'],
            '24h Price Change (%)': coin['price_change_percentage_24h']
        })

    df = pd.DataFrame(cryptos)

    # Save the updated data to the Excel file
    df.to_excel("cryptocurrency_data.xlsx", index=False)

    # Wait for 5 minutes before fetching new data
    time.sleep(300)  # Sleep for 300 seconds (5 minutes)