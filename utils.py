# utils.py
# Utility functions for parsing data, and now, for scraping per diem rates.

import re
import pdfplumber
import calendar
import config
import time
from datetime import datetime
from bs4 import BeautifulSoup
from dateutil.parser import parse as parse_date
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def get_usd_to_inr_rate(report_date):
    """
    Gets the USD to INR conversion rate for a specific date.
    Uses the middle of the month for a stable average.
    """
    url = "https://www.oanda.com/currency-converter/en/?from=USD&to=INR&amount=1"
    rates = {}
    
    # Setup Selenium WebDriver
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'eager' # As requested, for faster interaction
    options.add_argument('--headless') # Run in background without opening a browser window
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 15) # Wait for up to 15 seconds

    try:
        if config.DEBUG_MODE: print(f"Navigating to per diem website for {report_date.strftime('%Y-%m-%d')}...")
        driver.get(url)
        
        # Wait for the country dropdown to be present before interacting with it.
        date_selector = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/main/div[2]/div/div/div[3]/div/div[1]/div[1]/div/div[3]/div[1]/div[2]/div/div/input')))
        date_selector.value = report_date.strftime('%Y-%m-%d')
        time.sleep(5)  # Allow time for the page to update with the new date
        rate_value = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/main/div[2]/div/div/div[3]/div/div[1]/div[1]/div/div[2]/div[3]/div[2]/div[1]/div/input')))
        if config.DEBUG_MODE: print(f"Using USD to INR exchange rate from {report_date.strftime('%Y-%m-%d')}: {rate_value.get_attribute('value')}")
        return float(rate_value.get_attribute('value'))
    except Exception as e:
        print(f"Could not fetch currency conversion rate: {e}")
        # Return a fallback rate if the API fails
        return 85.0 
    finally:
        driver.quit()

def get_mie_breakdown():
    return config.MIE_BREAKDOWN


def get_per_diem_rates_with_selenium(year, month, country_name="India"):
    """
    Uses Selenium to navigate the US State Dept website and scrape per diem rates.
    """
    url = "https://allowances.state.gov/web920/per_diem.asp"
    rates = {}
    
    # Setup Selenium WebDriver
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'eager' # As requested, for faster interaction
    options.add_argument('--headless') # Run in background without opening a browser window
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 15) # Wait for up to 15 seconds

    try:
        if config.DEBUG_MODE: print(f"Navigating to per diem website for {calendar.month_name[month]} {year}...")
        driver.get(url)
        
        # Wait for the country dropdown to be present before interacting with it.
        country_dropdown_element = wait.until(EC.presence_of_element_located((By.NAME, 'CountryCode')))
        country_dropdown = Select(country_dropdown_element)
        # Use the exact case from the website's dropdown
        country_dropdown.select_by_visible_text(country_name.upper())
        
        # Find the "Go" button next to the country selector and click it
        go_button_1 = country_dropdown_element.find_element(By.XPATH, "../following-sibling::td/input")
        go_button_1.click()
        
        # Wait for the month dropdown (PublicationDate) to appear on the next page
        month_dropdown_element = wait.until(EC.presence_of_element_located((By.NAME, 'PublicationDate')))
        month_dropdown = Select(month_dropdown_element)
        # The value format is YYYYMMDD, we just need the month part
        month_value = f"{year}{str(month).zfill(2)}01"
        month_dropdown.select_by_value(month_value)

        # Find the "Go" button next to the date selector and click it
        go_button_2 = month_dropdown_element.find_element(By.XPATH, "../following-sibling::td/input")
        go_button_2.click()
        
        # Now scrape the final table, waiting for it to be present
        wait.until(EC.presence_of_element_located((By.XPATH, "//td[@title='Country Name']/..")))
        rows = driver.find_elements(By.XPATH, "//td[@title='Country Name']/..")
        if not rows:
            print("Error: Per diem rates table not found on the final page.")
            return None
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, 'td')
            if len(cols) >= 6:
                post_name = cols[1].text # "Post Name" is the second column
                lodging = int(cols[4].text)
                mie = int(cols[5].text)
                rates[post_name] = {"lodging": lodging, "total_mie": mie}
        
        if config.DEBUG_MODE: print(f"Successfully scraped per diem rates for {len(rates)} locations in {country_name}.")
        return rates

    except Exception as e:
        print(f"An error occurred during Selenium scraping: {e}")
        return None
    finally:
        driver.quit()

def parse_flight_pdf(pdf_path):
    """
    Parses a flight confirmation PDF to extract travel details for all flight legs.
    """
    flights = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text.replace('\n', ' ')

            pattern = re.compile(r"(\d{1,2}-[A-Za-z]{3}-\d{4})-([A-Za-z\s]+) to ([A-Za-z\s]+)- by Air")
            matches = pattern.finditer(full_text)

            for match in matches:
                date_str = match.group(1)
                from_city = match.group(2).strip()
                to_city = match.group(3).strip()
                travel_date = datetime.strptime(date_str, '%d-%b-%Y').date()
                flights.append({"date": travel_date, "from": from_city, "to": to_city})
                if config.DEBUG_MODE: print(f"  -> Found flight in PDF: {from_city} to {to_city} on {travel_date}")
        return flights
    except Exception as e:
        print(f"Error parsing PDF file {pdf_path}: {e}")
        return []

    
def find_fare_city(address):
    """
    Extracts the city name from an Indian address string as the third comma-separated word from the last.
    If not enough parts, returns the last part or 'N/A'.
    """
    city = "NA"
    if address and isinstance(address, str):
        parts = [p.strip() for p in address.split(',') if p.strip()]
        if len(parts) >= 3:
            city = parts[-3]
        elif parts:
            city = parts[-1]
    if config.DEBUG_MODE: print (f"Extracted city from address '{address}': {city}")
    return city

def parse_uber_receipt_email(email_body):
    """
    Parses the HTML content of an Uber receipt email to extract trip details.
    """
    soup = BeautifulSoup(email_body, 'html.parser')
    details = {"from": "N/A", "to": "N/A", "fare": "N/A", "date": None, "fare-city": "N/A"}

    try:
        total_header_tag = soup.find('td', class_='total_head', string='Total')
        if total_header_tag:
            total_value_tag = total_header_tag.find_next_sibling('td', class_='total_head')
            if total_value_tag:
                fare_match = re.search(r'[\d,]+\.\d{2}', total_value_tag.get_text())
                if fare_match:
                    details["fare"] = fare_match.group(0)

        header_date_tag = soup.find('span', class_='Uber18_text_p1', string=re.compile(r'\w+\s\d{1,2},\s\d{4}'))
        if header_date_tag:
            details["date"] = parse_date(header_date_tag.get_text(strip=True)).date()
        
        time_pattern = re.compile(r'\d{1,2}:\d{2}\s*(?:AM|PM)')
        time_tags = soup.find_all(string=time_pattern)
        
        addresses = []
        for tag in time_tags:
            tr_time = tag.find_parent('tr')
            if tr_time:
                tr_address = tr_time.find_next_sibling('tr')
                if tr_address:
                    address_tag = tr_address.find('td')
                    if address_tag:
                        address = address_tag.get_text(strip=True)
                        if len(address) > 5:
                            addresses.append(address)
        
        if len(addresses) >= 2:
            details["from"] = addresses[0]
            details["to"] = addresses[1]
            details["fare-city"] = find_fare_city(details["from"])  
    except Exception as e:
        print(f"Warning: An error occurred while parsing Uber receipt: {e}")
    
    return details
