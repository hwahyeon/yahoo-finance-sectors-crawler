import requests
import re
import os
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import Workbook

# List of sectors to scrape
sectors = [
    "technology",
    "financial-services",
    "healthcare",
    "consumer-cyclical",
    "communication-services",
    "industrials",
    "consumer-defensive",
    "energy",
    "basic-materials",
    "real-estate",
    "utilities"
]

def fetch_data():
    base_url = "https://finance.yahoo.com/sectors/{}"
    all_data = {sector: [] for sector in sectors}
    headers = {
        "Accept-Language": "en-US,en;q=0.9",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36"
    }

    for sector in sectors:
        url = base_url.format(sector)
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
        except requests.RequestException as e:
            print(f"Request error: {e}")
            return

        soup = BeautifulSoup(response.content, "html.parser")
        heatmap_container = soup.find('div', class_='heatMap-container')
        if not heatmap_container:
            print(f"Could not find 'heatMap-container' on the {sector} page.")
            continue

        # Find 'a' tags matching the pattern
        pattern = re.compile(r'none-link.*fin-size-medium|fin-size-medium.*none-link')
        sector_links = heatmap_container.find_all('a', class_=pattern)

        if not sector_links:
            print(f"No data found on the {sector} page.")
            continue

        for link in sector_links:
            sector_name = link.find('div', class_='ticker-div').text.strip()
            percent_change = link.find('div', class_='percent-div').text.strip()
            all_data[sector].append((sector_name, percent_change))

        # Sort data by sector_name
        all_data[sector].sort(key=lambda x: x[0])

    if not any(all_data.values()):
        print("No data was found.")
        return

    # Save data to an Excel file
    wb = Workbook()
    ws = wb.active

    # Add column headers
    headers = ["page"] + [sector for sector in sectors for _ in all_data[sector]]
    ws.append(headers)

    # Add sector row
    sector_row = ["sector"] + [sector_name for sector in sectors for sector_name, _ in all_data[sector]]
    ws.append(sector_row)

    # Add change row
    change_row = ["change"] + [change for sector in sectors for _, change in all_data[sector]]
    ws.append(change_row)

    today_date = datetime.today().strftime('%Y-%m-%d')
    file_name = f"output_{today_date}.xlsx"
    wb.save(file_name)

    print(f"Data has been saved to an Excel file: {file_name}")
    return file_name

def send_email(file_path):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.application import MIMEApplication

    EMAIL_ADDRESS = os.environ['EMAIL_ADDRESS']
    EMAIL_PASSWORD = os.environ['EMAIL_PASSWORD']
    recipient = os.environ['RECIPIENT_EMAIL']

    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient
    msg['Subject'] = 'Crawling Data Result'

    # Attach the Excel file
    with open(file_path, 'rb') as file:
        attachment = MIMEApplication(file.read(), _subtype="xlsx")
        attachment.add_header('Content-Disposition', 'attachment', filename=file_path)
        msg.attach(attachment)

    # Send email
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    server.send_message(msg)
    server.quit()

    print("Email has been sent successfully.")

if __name__ == "__main__":
    excel_file = fetch_data()
    if excel_file:
        send_email(excel_file)
