# yahoo_service.py
# This module handles connecting to Yahoo Mail via IMAP to fetch Uber receipts.

import imaplib
import email
from datetime import datetime
import config
import utils # Import the utils module to access the new function

IMAP_SERVER = "imap.mail.yahoo.com"

def connect_to_yahoo(email_address, app_password):
    """Connects and logs into the Yahoo IMAP server."""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(email_address, app_password)
        mail.select("inbox")
        print("Successfully connected to Yahoo Mail.")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Error connecting to Yahoo Mail: {e}")
        print("Please check your email and app password in config.py.")
        print("Ensure you have generated and are using a 16-character 'App Password'.")
        return None

def search_uber_receipts(mail_session, travel_date, usd_to_inr_rate):
    """
    Searches for Uber receipts on a specific date, saving only those over $10.
    """
    date_str = travel_date.strftime("%d-%b-%Y") # e.g., 29-Jul-2025
    search_query = f'(FROM "noreply@uber.com" SUBJECT "trip with Uber" ON "{date_str}")'
    if config.DEBUG_MODE: print(f"Executing Yahoo search with query: {search_query}")

    try:
        _, selected_mails = mail_session.search(None, search_query)
        email_ids = selected_mails[0].split()
        if not email_ids:
            return []

        if config.DEBUG_MODE: print(f"Found {len(email_ids)} Uber receipt(s) for {date_str}.")
        
        receipts = []
        for email_id in email_ids:
            _, data = mail_session.fetch(email_id, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    ctype = part.get_content_type()
                    if ctype == "text/html":
                        body = part.get_payload(decode=True)
                        break
            else:
                body = msg.get_payload(decode=True)

            if body:
                html_string = body.decode('utf-8', 'ignore')
                # Parse the details first to get the fare
                uber_details = utils.parse_uber_receipt_email(html_string)
                
                # Check if the fare exceeds $10
                save_receipt = False
                try:
                    inr_fare = float(uber_details.get("fare", "0").replace(",", ""))
                    usd_equivalent = inr_fare / usd_to_inr_rate
                    if usd_equivalent > 10:
                        save_receipt = True
                        if config.DEBUG_MODE: print(f"  -> Ride fare is ₹{inr_fare:.2f} (${usd_equivalent:.2f}), saving receipt.")
                    else:
                        if config.DEBUG_MODE: print(f"  -> Ride fare is ₹{inr_fare:.2f} (${usd_equivalent:.2f}), skipping receipt save.")
                except (ValueError, TypeError):
                    print("  -> Could not parse fare to check against $10 limit.")

                filepath = None
                total_receipts = len(receipts) - 1
                duplicate_found = (total_receipts >= 0 and receipts[total_receipts].get("fare", "0") == uber_details.get("fare", "0") and receipts[total_receipts].get("date") == uber_details.get("date") and receipts[total_receipts].get("from") == uber_details.get("from") and receipts[total_receipts].get("to") == uber_details.get("to"))

                if save_receipt and not duplicate_found:
                    html_filename = f"uber_receipt_{travel_date.strftime('%Y%m%d')}_{email_id.decode()}.html"
                    with open(html_filename, "w", encoding="utf-8") as html_file:
                        html_file.write(html_string)
                    filepath = html_filename
                
                # Always add the ride details to the list for the spreadsheet
                uber_details["filepath"] = filepath
                if not duplicate_found:
                    receipts.append(uber_details)
                else:
                    if config.DEBUG_MODE: print("  -> Duplicate receipt found, skipping.")
        
        return receipts

    except Exception as e:
        print(f"An error occurred while searching Yahoo Mail: {e}")
        return []

def close_connection(mail_session):
    """Closes the IMAP connection."""
    if mail_session:
        mail_session.logout()
        if config.DEBUG_MODE: print("Disconnected from Yahoo Mail.")
