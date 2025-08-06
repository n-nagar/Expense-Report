# utils.py
# Utility functions for parsing data from PDF files and email content.

import re
import pdfplumber
from datetime import datetime
from bs4 import BeautifulSoup, NavigableString
from dateutil.parser import parse as parse_date

def parse_flight_pdf(pdf_path):
    """
    Parses a flight confirmation PDF to extract travel details for all flight legs.
    Uses pdfplumber to read the PDF text and regex to find itinerary lines.
    Returns a list of flight dictionaries.
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
                print(f"  -> Found flight in PDF: {from_city} to {to_city} on {travel_date}")
        return flights
    except Exception as e:
        print(f"Error parsing PDF file {pdf_path}: {e}")
        return []

def parse_uber_receipt_email(email_body):
    """
    Parses the HTML content of an Uber receipt email to extract trip details.
    This logic is now highly specific to the Uber receipt's HTML structure and more robust.
    """
    soup = BeautifulSoup(email_body, 'html.parser')
    details = {"from": "N/A", "to": "N/A", "fare": "N/A", "date": None}

    try:
        # Find the total fare by looking for the specific class and text.
        total_header_tag = soup.find('td', class_='total_head', string='Total')
        if total_header_tag:
            total_value_tag = total_header_tag.find_next_sibling('td', class_='total_head')
            if total_value_tag:
                fare_match = re.search(r'[\d,]+\.\d{2}', total_value_tag.get_text())
                if fare_match:
                    details["fare"] = fare_match.group(0)

        # Find date from a more reliable element in the receipt header.
        header_date_tag = soup.find('span', class_='Uber18_text_p1', string=re.compile(r'\w+\s\d{1,2},\s\d{4}'))
        if header_date_tag:
            details["date"] = parse_date(header_date_tag.get_text(strip=True)).date()
        
        # Find addresses by locating the time, then finding the address in the next table row.
        time_pattern = re.compile(r'\d{1,2}:\d{2}\s*(?:AM|PM)')
        time_tags = soup.find_all(string=time_pattern)
        
        addresses = []
        for tag in time_tags:
            # The time is in a <td> inside a <tr>. The address is in the next <tr>.
            tr_time = tag.find_parent('tr')
            if tr_time:
                tr_address = tr_time.find_next_sibling('tr')
                if tr_address:
                    address_tag = tr_address.find('td')
                    if address_tag:
                        address = address_tag.get_text(strip=True)
                        if len(address) > 5: # Basic check for a valid address
                            addresses.append(address)
        
        if len(addresses) >= 2:
            details["from"] = addresses[0]
            details["to"] = addresses[1]

    except Exception as e:
        print(f"Warning: An error occurred while parsing Uber receipt: {e}")

    return details
