# yahoo_service.py
# This module handles connecting to Yahoo Mail via IMAP to fetch Uber receipts.

import imaplib
import email
from datetime import datetime
import config
import utils # Import the utils module to access the new function
import os

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
                current_from = uber_details.get("from", "N/A")
                current_to = uber_details.get("to", "N/A")
                current_fare = uber_details.get("fare", "0")
                has_valid_addresses = current_from != "N/A" and current_to != "N/A"

                # Check for duplicate receipts (same fare on same date)
                duplicate_index = None
                for i, existing in enumerate(receipts):
                    if existing.get("fare", "0") == current_fare:
                        duplicate_index = i
                        break

                if duplicate_index is not None:
                    existing = receipts[duplicate_index]
                    existing_has_valid = existing.get("from", "N/A") != "N/A" and existing.get("to", "N/A") != "N/A"

                    if has_valid_addresses and not existing_has_valid:
                        # Replace N/A receipt with this one that has valid addresses
                        if config.DEBUG_MODE:
                            print(f"  -> Replacing N/A receipt with valid address version.")
                        # Delete old PDF if exists
                        if existing.get("filepath") and os.path.exists(existing["filepath"]):
                            os.remove(existing["filepath"])
                        receipts.pop(duplicate_index)
                        # Continue to add this receipt below
                    else:
                        # Skip this duplicate (either both have valid addresses or this one has N/A)
                        if config.DEBUG_MODE:
                            print("  -> Duplicate receipt found, skipping.")
                        continue

                if save_receipt:
                    html_filename = f"uber_receipt_{travel_date.strftime('%Y%m%d')}_{email_id.decode()}.html"
                    with open(html_filename, "w", encoding="utf-8") as html_file:
                        html_file.write(html_string)

                    # Convert HTML to PDF
                    pdf_filename = html_filename.replace(".html", ".pdf")
                    utils.html_to_pdf_chrome(html_filename, pdf_filename)

                    # Delete the HTML after conversion
                    os.remove(html_filename)

                    filepath = pdf_filename

                # Add receipt to list
                uber_details["filepath"] = filepath
                receipts.append(uber_details)
        
        return receipts

    except Exception as e:
        print(f"An error occurred while searching Yahoo Mail: {e}")
        return []

def close_connection(mail_session):
    """Closes the IMAP connection."""
    if mail_session:
        mail_session.logout()
        if config.DEBUG_MODE: print("Disconnected from Yahoo Mail.")
