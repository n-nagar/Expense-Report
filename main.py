# main.py
# The main script to orchestrate the entire expense reporting process.

import os
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from googleapiclient.discovery import build
from collections import defaultdict
import calendar
import time

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

    # Scrape Per Diem and Currency Rates
    if config.DEBUG_MODE: print("\n--- Scraping Per Diem & Currency Rates ---")
    per_diem_rates = utils.get_per_diem_rates_with_selenium(report_year, report_month)
    mie_breakdown = utils.get_mie_breakdown()
    usd_to_inr_rate = utils.get_usd_to_inr_rate(report_month_date)
    
    if not per_diem_rates or not mie_breakdown or not usd_to_inr_rate:
        print("Could not retrieve per diem or currency rates. Exiting.")
        return

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
    
    search_start_date = report_month_date - relativedelta(months=1)
    gmail_search_after = search_start_date.strftime('%Y/%m/%d')
    gmail_search_before = first_day_of_current_month.strftime('%Y/%m/%d')
    
    query = f'from:"{config.TRAVEL_EMAIL_SENDER}" has:attachment after:{gmail_search_after} before:{gmail_search_before}'
    if config.DEBUG_MODE: print(f"\nSearching Gmail for travel emails from {gmail_search_after} to {gmail_search_before}...")
    messages = google_services.search_gmail(gmail_service, query)
    
    if config.DEBUG_MODE: print(f"\nFound {len(messages)} potential travel emails in Gmail.")

    for msg in messages:
        msg_id = msg['id']
        message_details = gmail_service.users().messages().get(userId='me', id=msg_id).execute()
        
        parts_to_search = list(message_details['payload'].get('parts', []))
        while parts_to_search:
            part = parts_to_search.pop(0)
            if part.get("parts"):
                parts_to_search.extend(part.get("parts"))

            filename = part.get('filename')
#            if filename and filename.startswith('GCMA') and filename.endswith('.pdf'):
            if filename and filename.endswith('.pdf'):
                pdf_path = google_services.get_gmail_attachment(gmail_service, msg_id, filename)
                if pdf_path:
                    flights_in_pdf = utils.parse_flight_pdf(pdf_path)
                    all_flights.extend(flights_in_pdf)
                    if pdf_path not in travel_pdf_paths:
                        travel_pdf_paths.append(pdf_path)

    # Filter for flights within the report month and sort them
    relevant_flights = sorted([f for f in all_flights if f['departure'].month == report_month and f['departure'].year == report_year], key=lambda x: x['departure'])
    
    if not relevant_flights:
        print("No relevant travel bookings found for the report month. Exiting.")
        return

    # 4. Create Travel Calendar to determine nightly location
    if config.DEBUG_MODE: print("\n--- Building Travel Calendar ---")
    _, num_days_in_month = calendar.monthrange(report_year, report_month)
    travel_calendar = {}
    current_location = "Bangalore"

    flights_by_date = defaultdict(list)

    for f in relevant_flights:
        if 'date' in f and isinstance(f['date'], date):  # ensure key exists and is a date object
            flights_by_date[f['date']].append(f)

    unique_travel_dates = []
    for day_num in range(1, num_days_in_month + 1):
        current_date = date(report_year, report_month, day_num)
        unique_travel_dates.append(current_date)
        travel_calendar[current_date] = current_location
        if current_date in flights_by_date:
            flight = flights_by_date[current_date][-1]  # Get the last flight of the day
            if "bangalore" in flight['to'].lower():
                travel_calendar[current_date] = flight['from']
                current_location = "Bangalore"
            else:
                travel_calendar[current_date] = flight['to']
                current_location = flight['to']

    # 5. Search Yahoo Mail for Uber receipts
    uber_data = []
    uber_receipt_paths = []
    yahoo_mail = yahoo_service.connect_to_yahoo(config.YAHOO_EMAIL, config.YAHOO_APP_PASSWORD)
    if yahoo_mail:
        for travel_date in unique_travel_dates:
            receipts = yahoo_service.search_uber_receipts(yahoo_mail, travel_date, usd_to_inr_rate)
            if receipts:
                for receipt_details in receipts:
                    receipt_details['date'] = travel_date
                    uber_data.append(receipt_details)
                    if receipt_details.get("filepath"):
                        uber_receipt_paths.append(receipt_details["filepath"])
        yahoo_service.close_connection(yahoo_mail)

    # 6. Create Google Drive folder and upload files
    if config.SAVE_TO_DRIVE:
        base_folder_name = "https://drive.google.com/drive/folders/1dGFeh28Bzb0jPnJ9VmMoRR3xR-Avkp9Y?usp=sharing"
        drive_folder_name = report_month_date.strftime("%m-%Y")
        folder_id = google_services.create_drive_folder(drive_service, drive_folder_name)    
        if folder_id:
            for path in travel_pdf_paths + uber_receipt_paths:
                google_services.upload_file_to_drive(drive_service, path, folder_id)
                os.remove(path)
                time.sleep(1) 
            # 7. Create and populate Google Sheet
            sheet_name = config.DRIVE_SHEET_NAME.format(month_name=report_month_date.strftime('%B'), year=report_year)
            spreadsheet_id = google_services.create_google_sheet(drive_service, sheet_name, folder_id)

            if spreadsheet_id:
                tab_configs = [
                    {"name": "Per Diem & Lodging", "headers": ["Date(s) Claimed:", "City, Country*", "State Department M&IE (Per Diem Rate)*", "Breakfast**", "Lunch**", "Dinner**", "Incidentals**", "Total M&IE for Date", "Lodging Cost For Night", "Running Total", "Comments"]},
                    {"name": "Reimbursements", "headers": ["Expenditure Date (When):", "Receipt # *", "Expenditure Location (Where)", "Expenditure Currency", "Expenditure Description - Who, What, Why", "Receipt Amt in Receipt Currency", "Rate of Exchange", "US Dollar Equivalent", "Comments"]}
                ]
                google_services.setup_spreadsheet_tabs(sheets_service, spreadsheet_id, tab_configs)

        # Prepare Per Diem data
        per_diem_rows = []
        row_counter = 2 # Start from row 2 for the first data row
        running_total_formula = f"=H{row_counter}"
        for day_num in range(1, num_days_in_month + 1):
            current_date = date(report_year, report_month, day_num)
            location = travel_calendar.get(current_date, "Bangalore")
            
            bfast, lunch, dinner, incidentals, total_mie_rate, lodging = [""] * 6
            
            if "Bangalore" in location:
                rates = config.PER_DIEM_RATES_USD["Bangalore"]
                bfast, lunch, dinner, incidentals = rates["breakfast"], rates["lunch"], rates["dinner"], rates["incidentals"]
                total_mie_rate = rates.get("total_mie", "")
            else:
                city_rates = per_diem_rates.get(location, per_diem_rates.get("Other"))
                if city_rates:
                    lodging = city_rates["lodging"]
                    total_mie_rate = city_rates["total_mie"]
                    breakdown = mie_breakdown.get(total_mie_rate, {})
                    
                    bfast = config.PER_DIEM_RATES_USD["Bangalore"]["breakfast"]
                    lunch = breakdown.get("lunch", "")
                    dinner = breakdown.get("dinner", "")
                    incidentals = breakdown.get("incidentals", "")

            total_formula = f"=SUM(D{row_counter}:G{row_counter})"

            per_diem_rows.append([
                current_date.strftime('%Y-%m-%d'), f"{location}, India", total_mie_rate,
                bfast, lunch, dinner, incidentals, total_formula, "", running_total_formula, ""
            ])
            row_counter += 1
            running_total_formula = f"=J{row_counter-1}+H{row_counter}"

        per_diem_rows.append([
            "", "", "", "", "", "", "TOTAL PER DIEM", f"=SUM(H2:H{row_counter-1})", "", f"=J{row_counter-1}", ""
        ])

        if (config.SAVE_TO_DRIVE and per_diem_rows):
            google_services.append_values(sheets_service, spreadsheet_id, "Per Diem & Lodging", per_diem_rows)

        # Prepare Reimbursements data
        reimbursement_rows = []
        row_counter = 2 # Start from row 2 for the first data row
        for item in sorted(uber_data, key=lambda x: x['date']):
            description = f"Uber from {item.get('from', 'N/A')} to {item.get('to', 'N/A')}"
            reimbursement_rows.append([
                item['date'].strftime('%Y-%m-%d'),
                item['filepath'] or "",
                item['fare-city'] or "N/A",
                "INR",
                description,
                item.get('fare', 'N/A'),
                usd_to_inr_rate,
                f"=F{row_counter}/G{row_counter}",
                ""
            ])
            row_counter += 1
        
        reimbursement_rows.append([
            "",
            "",
            "",
            "",
            "",
            "",
            "TOTAL REIMBURSEMENTS",
            f"=SUM(H2:H{row_counter-1})",
            ""
        ])
        if config.SAVE_TO_DRIVE and reimbursement_rows:
            google_services.append_values(sheets_service, spreadsheet_id, "Reimbursements", reimbursement_rows)

    print("\n--- Expense Report Automation Finished Successfully! ---")
    if config.SAVE_TO_DRIVE:
        print(f"Your report has been saved to Google Drive in the folder '{drive_folder_name}'.")

if __name__ == "__main__":
    main()
