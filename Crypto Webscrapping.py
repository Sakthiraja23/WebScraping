from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = " Top 10 Cryptocurrency"
sheet.append(['Rank','Name','Price','Market Cap','Volume','Circulating Supply'])

try:
    response = requests.get("https://coinmarketcap.com/")
    soup = BeautifulSoup(response.text, 'html.parser')

    table_body = soup.find('table', class_="sc-db1da501-3 ccGPRR cmc-table").find("tbody")
    rows = table_body.find_all('tr')

    count = 0  # Counter to limit to first 10

    for row in rows:
        cols = row.find_all('td')
        if len(cols) < 10:
            continue  # Skip malformed rows

        try:
            # Rank
            rank = cols[1].text.strip()

            # Name
            name_tag = cols[2].find('p', class_='coin-item-name')
            name = name_tag.text.strip() if name_tag else "N/A"

            # Price
            price_tag = cols[3].find('span')
            price = price_tag.text.strip() if price_tag else "N/A"

            # Market Cap
            market_cap_tag = cols[7].find('span', class_='sc-11478e5d-0')
            market_cap = market_cap_tag.text.strip() if market_cap_tag else "N/A"

            # Volume (24h)
            volume_tag = cols[8].find('p')
            volume = volume_tag.text.strip() if volume_tag else "N/A"

            # Circulating Supply
            supply_tag = cols[9].find('div', class_='circulating-supply-value')
            circulating_supply = supply_tag.text.strip() if supply_tag else "N/A"

            # Final Output
            # print(rank, name, price, market_cap, volume, circulating_supply)
            sheet.append([rank, name, price, market_cap, volume, circulating_supply])

            count += 1
            if count == 10:
                break  # Stop after first 10 valid rows

        except Exception as e:
            print(f"Skipping row due to error: {e}")

except Exception as e:
    print("Error:", e)

excel.save("Top 10 Cryptocurrency")
