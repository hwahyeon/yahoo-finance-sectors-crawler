import requests
import re
import os
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import Workbook

# List of sectors to scrape
sectors = [
    "",
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
        heatmap = soup.find('div', class_='heatMap-container')
        if not heatmap:
            print(f"Could not find 'heatMap-container' on the {sector} page.")
            continue

        boxes = heatmap.find_all("div", class_="rect-container")
        if not boxes:
            print(f"No data found on the {sector} page.")
            continue
        
        for box in boxes:
            ticker_div = box.find("div", class_="ticker-div")
            percent_div = box.find("div", class_="percent-div")
            if ticker_div and percent_div:
                sector_name = ticker_div.text.strip()
                percent_change = percent_div.text.strip()
                all_data[sector].append((sector_name, percent_change))

        all_data[sector].sort(key=lambda x: x[0])       

    if not any(all_data.values()):
        print("No data was found.")
        return

    # Save data to an Excel file
    wb = Workbook()
    ws = wb.active

    # Add column headers
    headers = ["page"] + [("all" if sector == "" else sector) for sector in sectors for _ in all_data[sector]]
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
