# yahoo_service.py
# This module handles connecting to Yahoo Mail via IMAP to fetch Uber receipts.

import imaplib
import email
from datetime import datetime

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

def search_uber_receipts(mail_session, travel_date):
    """
    Searches for Uber receipts on a specific date.
    Returns a list of email bodies.
    """
    date_str = travel_date.strftime("%d-%b-%Y") # e.g., 29-Jul-2025
    search_query = f'(FROM "noreply@uber.com" SUBJECT "trip with Uber" ON "{date_str}")'
    print(f"Executing Yahoo search with query: {search_query}")

    try:
        _, selected_mails = mail_session.search(None, search_query)
        email_ids = selected_mails[0].split()
        if not email_ids:
            return []

        print(f"Found {len(email_ids)} Uber receipt(s) for {date_str}.")
        
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
                # Save the HTML body for a printable receipt. This is more stable than PDF conversion.
                html_string = body.decode('utf-8', 'ignore')
                html_filename = f"uber_receipt_{travel_date.strftime('%Y%m%d')}_{email_id.decode()}.html"
                
                with open(html_filename, "w", encoding="utf-8") as html_file:
                    html_file.write(html_string)

                print(f"  -> Successfully saved Uber receipt to {html_filename}")
                receipts.append({"body": html_string, "filepath": html_filename})
        
        return receipts

    except Exception as e:
        print(f"An error occurred while searching Yahoo Mail: {e}")
        return []

def close_connection(mail_session):
    """Closes the IMAP connection."""
    if mail_session:
        mail_session.logout()
        print("Disconnected from Yahoo Mail.")
