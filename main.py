# main.py
# The main script to orchestrate the entire expense reporting process.

import os
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from googleapiclient.discovery import build
import calendar

# Import project modules
import config
import google_services
import yahoo_service
import utils

def main():
    """Main function to run the expense automation."""
    print("--- Starting Expense Report Automation ---")

    # 1. Determine the date range for the report
    today = date.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_report_month = first_day_of_current_month - relativedelta(days=1)
    report_month_date = last_day_of_report_month.replace(day=1)

    report_month = report_month_date.month
    report_year = report_month_date.year

    search_start_date = report_month_date - relativedelta(months=1)
    gmail_search_after = search_start_date.strftime('%Y/%m/%d')
    gmail_search_before = first_day_of_current_month.strftime('%Y/%m/%d')

    print(f"Generating report for: {report_month_date.strftime('%B %Y')}")
    print(f"Searching for travel bookings received from {gmail_search_after} to {gmail_search_before}")

    # 2. Authenticate with Google Services
    creds = google_services.authenticate()
    if not creds:
        print("Failed to authenticate with Google. Exiting.")
        return
        
    gmail_service = build("gmail", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)

    # 3. Search Gmail for travel confirmation emails
    all_flights = []
    travel_pdf_paths = []
    
    query = f'from:"{config.TRAVEL_EMAIL_SENDER}" has:attachment after:{gmail_search_after} before:{gmail_search_before}'
    messages = google_services.search_gmail(gmail_service, query)
    
    print(f"Found {len(messages)} potential travel emails in Gmail.")

    for msg in messages:
        msg_id = msg['id']
        message_details = gmail_service.users().messages().get(userId='me', id=msg_id).execute()
        
        parts_to_search = list(message_details['payload'].get('parts', []))
        while parts_to_search:
            part = parts_to_search.pop(0)
            if part.get("parts"):
                parts_to_search.extend(part.get("parts"))

            filename = part.get('filename')
            if filename and filename.startswith('GCMA') and filename.endswith('.pdf'):
                pdf_path = google_services.get_gmail_attachment(gmail_service, msg_id, filename)
                if pdf_path:
                    flights_in_pdf = utils.parse_flight_pdf(pdf_path)
                    all_flights.extend(flights_in_pdf)
                    if pdf_path not in travel_pdf_paths:
                        travel_pdf_paths.append(pdf_path)

    # Filter for flights within the report month and sort them
    relevant_flights = sorted([f for f in all_flights if f['date'].month == report_month and f['date'].year == report_year], key=lambda x: x['date'])
    
    if not relevant_flights:
        print("No relevant travel bookings found for the report month. Exiting.")
        return

    # 4. Create Travel Calendar to determine nightly location
    print("\n--- Building Travel Calendar ---")
    _, num_days_in_month = calendar.monthrange(report_year, report_month)
    travel_calendar = {}
    current_location = "Bangalore"

    # Create a dictionary to easily look up flights by date
    flights_by_date = {}
    for f in relevant_flights:
        if f['date'] not in flights_by_date:
            flights_by_date[f['date']] = []
        flights_by_date[f['date']].append(f)

    for day_num in range(1, num_days_in_month + 1):
        current_date = date(report_year, report_month, day_num)
        
        # On the night of a travel day, the location is the destination city.
        if current_date in flights_by_date:
            # Handle multiple flights in one day (connections)
            last_flight_of_day = flights_by_date[current_date][-1]
            if "bangalore" in last_flight_of_day['from'].lower():
                 current_location = last_flight_of_day['to']
        
        travel_calendar[current_date] = current_location

        # If a return flight landed today, the location for the *next* day resets to Bangalore.
        if current_date in flights_by_date:
            last_flight_of_day = flights_by_date[current_date][-1]
            if "bangalore" in last_flight_of_day['to'].lower():
                current_location = "Bangalore"


    # 5. Search Yahoo Mail for Uber receipts on travel dates
    uber_data = []
    uber_receipt_paths = []
    unique_travel_dates = sorted(list(set(f['date'] for f in relevant_flights)))
    
    yahoo_mail = yahoo_service.connect_to_yahoo(config.YAHOO_EMAIL, config.YAHOO_APP_PASSWORD)
    if yahoo_mail:
        for travel_date in unique_travel_dates:
            receipts = yahoo_service.search_uber_receipts(yahoo_mail, travel_date)
            if receipts:
                for receipt in receipts:
                    uber_details = utils.parse_uber_receipt_email(receipt['body'])
                    uber_details['date'] = travel_date
                    uber_data.append(uber_details)
                    uber_receipt_paths.append(receipt['filepath'])
        yahoo_service.close_connection(yahoo_mail)

    # 6. Create Google Drive folder and upload files
    drive_folder_name = report_month_date.strftime("%m-%Y")
    folder_id = google_services.create_drive_folder(drive_service, drive_folder_name)
    
    if folder_id:
        for path in travel_pdf_paths + uber_receipt_paths:
            google_services.upload_file_to_drive(drive_service, path, folder_id)
            os.remove(path)

    # 7. Create and populate Google Sheet
    sheet_name = config.DRIVE_SHEET_NAME.format(month_name=report_month_date.strftime('%B'), year=report_year)
    spreadsheet_id = google_services.create_google_sheet(drive_service, sheet_name, folder_id)

    if spreadsheet_id:
        # Define the structure of the tabs based on the Excel template
        tab_configs = [
            {
                "name": "Per Diem & Lodging",
                "headers": ["Date(s) Claimed:", "City, Country*", "State Department M&IE (Per Diem Rate)*", "Breakfast**", "Lunch**", "Dinner**", "Incidentals**", "Total M&IE for Date", "Lodging Cost For Night", "Running Total", "Comments"]
            },
            {
                "name": "Reimbursements",
                "headers": ["Expenditure Date (When):", "Receipt # *", "Expenditure Location (Where)", "Expenditure Currency", "Expenditure Description - Who, What, Why", "Receipt Amt in Receipt Currency", "Rate of Exchange", "US Dollar Equivalent", "Comments"]
            }
        ]
        google_services.setup_spreadsheet_tabs(sheets_service, spreadsheet_id, tab_configs)

        # Prepare Per Diem data
        per_diem_rows = []
        for day_num in range(1, num_days_in_month + 1):
            current_date = date(report_year, report_month, day_num)
            location = travel_calendar.get(current_date, "Bangalore")
            
            rates = config.PER_DIEM_RATES_USD.get(location.lower(), {})
            
            bfast = rates.get("breakfast", "")
            lunch = rates.get("lunch", "")
            dinner = rates.get("dinner", "")
            incidentals = rates.get("incidentals", "")
            total_mie = sum(filter(None, [bfast, lunch, dinner, incidentals])) if any([bfast, lunch, dinner, incidentals]) else ""

            per_diem_rows.append([
                current_date.strftime('%Y-%m-%d'), 
                f"{location}, India", 
                "", # State Dept Rate - to be filled manually
                bfast, lunch, dinner, incidentals,
                total_mie,
                "", # Lodging - to be filled manually
                "", # Running Total - formula field
                ""  # Comments
            ])

        google_services.append_values(sheets_service, spreadsheet_id, "Per Diem & Lodging", per_diem_rows)

        # Prepare Reimbursements data
        reimbursement_rows = []
        receipt_counter = 1
        for item in sorted(uber_data, key=lambda x: x['date']):
            description = f"Uber from {item.get('from', 'N/A')} to {item.get('to', 'N/A')}"
            reimbursement_rows.append([
                item['date'].strftime('%Y-%m-%d'),
                f"Uber-{receipt_counter}",
                "Bangalore", # Assuming all Ubers are in Bangalore for now
                "INR",
                description,
                item.get('fare', 'N/A'),
                "", # Exchange Rate - to be filled manually
                "", # USD Equivalent - formula field
                ""  # Notes
            ])
            receipt_counter += 1
        
        if reimbursement_rows:
            google_services.append_values(sheets_service, spreadsheet_id, "Reimbursements", reimbursement_rows)

    print("\n--- Expense Report Automation Finished Successfully! ---")
    print(f"Check your Google Drive for the '{drive_folder_name}' folder.")

if __name__ == "__main__":
    main()
