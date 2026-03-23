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
import os
import base64
import requests
from datetime import date

def html_to_pdf_chrome(html_path, pdf_path):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless=new")  # headless mode
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--kiosk-printing")

    driver = webdriver.Chrome(options=chrome_options)

    file_url = "file://" + os.path.abspath(html_path)
    driver.get(file_url)

    # Tell Chrome to print to PDF via DevTools
    result = driver.execute_cdp_cmd(
        "Page.printToPDF",
        {
            "printBackground": True,  # keep CSS background colors
            "landscape": False,
            "scale": 1,
            "paperWidth": 8.27,  # A4
            "paperHeight": 11.69,  # A4
        }
    )

    pdf_data = base64.b64decode(result['data'])
    with open(pdf_path, "wb") as f:
        f.write(pdf_data)

    driver.quit()
def get_usd_to_inr_rate(report_date):
    url = f"https://api.frankfurter.app/{report_date.isoformat()}"
    params = {"from": "USD", "to": "INR"}
    return requests.get(url, params=params).json()["rates"]["INR"]

def get_usd_to_inr_rate_old(report_date):
    """
    Gets the USD to INR conversion rate for a specific date.
    Uses the middle of the month for a stable average.
    """
    url = "https://www.oanda.com/currency-converter/en/?from=USD&to=INR&amount=1"
    rates = {}
    
    # Setup Selenium WebDriver
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'normal' # As requested, for faster interaction
    #options.add_argument('--headless') # Run in background without opening a browser window
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 15) # Wait for up to 15 seconds

    try:
        if config.DEBUG_MODE: print(f"Navigating to per diem website for {report_date.strftime('%Y-%m-%d')}...")
        driver.get(url)
        
        # Wait for the country dropdown to be present before interacting with it.
        date_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@aria-label='Date']")))
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
    options.page_load_strategy = 'normal' # As requested, for faster interaction
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

def parse_hotel_reservation_pdf(pdf_path):
    """
    Parses a hotel reservation PDF to extract hotel details.
    Returns dict with hotel_name, address, checkin_date, checkout_date or None if not a hotel PDF.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"

            # Check if this is a hotel reservation (contains "Hotel Name" and "Checkin Date")
            if "Hotel Name" not in full_text or "Checkin Date" not in full_text:
                return None

            hotel_info = {}

            # Extract hotel name
            hotel_match = re.search(r'Hotel Name\s+([^\n]+)', full_text)
            if hotel_match:
                hotel_info["hotel_name"] = hotel_match.group(1).strip()

            # Extract address
            address_match = re.search(r'Address\s+([^\n]+)', full_text)
            if address_match:
                # Clean up address (remove extra commas/spaces)
                address = address_match.group(1).strip()
                address = re.sub(r',\s*,', ',', address)  # Remove double commas
                address = re.sub(r'\s+', ' ', address)  # Normalize spaces
                hotel_info["address"] = address

            # Extract check-in date
            checkin_match = re.search(r'Checkin Date\s+(\d{1,2}\s+\w+\s+\d{4})', full_text)
            if checkin_match:
                try:
                    hotel_info["checkin_date"] = datetime.strptime(
                        checkin_match.group(1), "%d %b %Y"
                    ).date()
                except ValueError:
                    pass

            # Extract check-out date
            checkout_match = re.search(r'CheckOut Date\s+(\d{1,2}\s+\w+\s+\d{4})', full_text)
            if checkout_match:
                try:
                    hotel_info["checkout_date"] = datetime.strptime(
                        checkout_match.group(1), "%d %b %Y"
                    ).date()
                except ValueError:
                    pass

            if hotel_info.get("hotel_name") and hotel_info.get("address"):
                if config.DEBUG_MODE:
                    print(f"  -> Found hotel reservation: {hotel_info['hotel_name']} at {hotel_info['address']}")
                return hotel_info

    except Exception as e:
        if config.DEBUG_MODE:
            print(f"Error parsing hotel PDF {pdf_path}: {e}")

    return None


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

            pattern = re.compile(
                r"(\d{1,2}-[A-Za-z]{3}-\d{4})-([A-Za-z\s]+) to ([A-Za-z\s]+)- by Air.*?Departs\s+(\d{2}:\d{2}).*?(\d{2}:\d{2}).*?Arrives",
                re.DOTALL
        )
            matches = pattern.finditer(full_text)

            for match in matches:
                date_str = match.group(1)
                from_city = match.group(2).strip()
                to_city = match.group(3).strip()
                dep_time_str = match.group(4)
                arr_time_str = match.group(5)

                # Combine date and time strings and parse to datetime objects
                datetime_format = "%d-%b-%Y %H:%M"
                dep_datetime = datetime.strptime(f"{date_str} {dep_time_str}", datetime_format)
                arr_datetime = datetime.strptime(f"{date_str} {arr_time_str}", datetime_format)
                if config.DEBUG_MODE: print(f"  -> Found flight in PDF: {from_city} to {to_city} on {dep_datetime} to {arr_datetime}")

                flights.append({
                    "from": from_city,
                    "to": to_city,
                    "date": dep_datetime.date(),
                    "departure": dep_datetime,
                    "arrival": arr_datetime
                })
        return flights
    except Exception as e:
        print(f"Error parsing PDF file {pdf_path}: {e}")
        return []

    
def classify_location(address, travel_city=None, hotel_reservations=None):
    """
    Classifies an address into a meaningful location name.

    Args:
        address: The address string to classify
        travel_city: Optional city context for the trip
        hotel_reservations: Optional list of hotel dicts with 'hotel_name' and 'address' keys

    Returns one of:
    - "Home" (if in Rajajinagar)
    - "Airport" (if contains airport keywords)
    - Hotel name (if matches a hotel reservation address)
    - "Hotel" (if contains generic hotel keywords)
    - "Restaurant" (if contains restaurant keywords)
    - Company name (if matches a known company location)
    - City name (fallback)
    """
    if not address or not isinstance(address, str):
        return "Unknown"

    addr_lower = address.lower()

    # Check for Home (Rajajinagar)
    if config.HOME_AREA.lower() in addr_lower:
        return "Home"

    # Check for Airport
    for keyword in config.AIRPORT_KEYWORDS:
        if keyword.lower() in addr_lower:
            return "Airport"

    # Check for specific hotel from reservations (by address match)
    if hotel_reservations:
        for hotel in hotel_reservations:
            hotel_addr = hotel.get("address", "").lower()
            hotel_name = hotel.get("hotel_name", "Hotel")
            # Check if key parts of the hotel address match
            # Extract street/road name for matching
            addr_parts = [p.strip() for p in hotel_addr.split(",") if len(p.strip()) > 3]
            for part in addr_parts[:2]:  # Check first 2 parts (usually street and area)
                if part in addr_lower:
                    return hotel_name
            # Also check if hotel name appears in address
            if hotel_name.lower() in addr_lower:
                return hotel_name

    # Check for generic Hotel keywords
    for keyword in config.HOTEL_KEYWORDS:
        if keyword.lower() in addr_lower:
            return "Hotel"

    # Check for Restaurant
    for keyword in config.RESTAURANT_KEYWORDS:
        if keyword.lower() in addr_lower:
            return "Restaurant"

    # Check for known companies (search all cities)
    for city, companies in config.COMPANIES.items():
        for company in companies:
            # Check if company name or part of it is in the address
            company_parts = company.lower().split()
            for part in company_parts:
                if len(part) > 3 and part in addr_lower:
                    return company

    # Fallback: return the city from the address
    city = find_fare_city(address)
    if city and city != "NA":
        return city

    return "Location"


def generate_uber_description(from_address, to_address, travel_city=None, hotel_reservations=None):
    """
    Generates a descriptive Uber ride description like:
    - "Home to Airport"
    - "Airport to Taj Samudra"
    - "Taj Samudra to Airport"

    Args:
        from_address: Pickup address
        to_address: Dropoff address
        travel_city: Optional city context
        hotel_reservations: Optional list of hotel dicts for matching
    """
    from_loc = classify_location(from_address, travel_city, hotel_reservations)
    to_loc = classify_location(to_address, travel_city, hotel_reservations)

    return f"{from_loc} to {to_loc}"


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
    Supports both old and new Uber email formats.
    """
    soup = BeautifulSoup(email_body, 'html.parser')
    details = {"from": "N/A", "to": "N/A", "fare": "N/A", "date": None, "fare-city": "N/A"}

    try:
        # === FARE EXTRACTION ===
        # New format: <td class="total-fare-amount">₹317.20</td>
        fare_tag = soup.find('td', class_='total-fare-amount')
        if fare_tag:
            fare_match = re.search(r'[\d,]+\.\d{2}', fare_tag.get_text())
            if fare_match:
                details["fare"] = fare_match.group(0)
        else:
            # Old format: <td class="total_head">Total</td> followed by sibling
            total_header_tag = soup.find('td', class_='total_head', string='Total')
            if total_header_tag:
                total_value_tag = total_header_tag.find_next_sibling('td', class_='total_head')
                if total_value_tag:
                    fare_match = re.search(r'[\d,]+\.\d{2}', total_value_tag.get_text())
                    if fare_match:
                        details["fare"] = fare_match.group(0)

        # === DATE EXTRACTION ===
        # New format: <div class="date">Mar 21, 2026 , 11:04 AM</div>
        date_div = soup.find('div', class_='date')
        if date_div:
            date_text = date_div.get_text(strip=True)
            # Extract just the date part (before the time)
            date_match = re.search(r'([A-Za-z]+\s+\d{1,2},\s*\d{4})', date_text)
            if date_match:
                details["date"] = parse_date(date_match.group(1)).date()
        else:
            # Old format: <span class="Uber18_text_p1">
            header_date_tag = soup.find('span', class_='Uber18_text_p1', string=re.compile(r'\w+\s\d{1,2},\s\d{4}'))
            if header_date_tag:
                details["date"] = parse_date(header_date_tag.get_text(strip=True)).date()

        # === ADDRESS EXTRACTION ===
        # New format: <td class="address-point-desc"> contains the address
        address_descs = soup.find_all('td', class_='address-point-desc')
        addresses = []
        for desc in address_descs:
            address = desc.get_text(strip=True)
            if len(address) > 10 and address not in addresses:
                addresses.append(address)

        # Fallback: Old format using time pattern
        if len(addresses) < 2:
            time_pattern = re.compile(r'\d{1,2}:\d{2}\s*(?:AM|PM)')
            time_tags = soup.find_all(string=time_pattern)
            for tag in time_tags:
                tr_time = tag.find_parent('tr')
                if tr_time:
                    tr_address = tr_time.find_next_sibling('tr')
                    if tr_address:
                        address_tag = tr_address.find('td')
                        if address_tag:
                            address = address_tag.get_text(strip=True)
                            if len(address) > 10 and address not in addresses:
                                addresses.append(address)

        if len(addresses) >= 2:
            details["from"] = addresses[0]
            details["to"] = addresses[1]
            details["fare-city"] = find_fare_city(details["from"])

    except Exception as e:
        print(f"Warning: An error occurred while parsing Uber receipt: {e}")

    return details
