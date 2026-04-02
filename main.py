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

def get_report_month_year():
    """Prompts the user to select the month and year for the expense report."""
    today = date.today()
    default_month = (today.replace(day=1) - relativedelta(days=1)).month
    default_year = (today.replace(day=1) - relativedelta(days=1)).year

    print("\n--- Select Report Period ---")
    print(f"Default: {calendar.month_name[default_month]} {default_year} (previous month)")

    # Prompt for year
    year_input = input(f"Enter year [{default_year}]: ").strip()
    if year_input:
        try:
            report_year = int(year_input)
        except ValueError:
            print(f"Invalid year. Using default: {default_year}")
            report_year = default_year
    else:
        report_year = default_year

    # Prompt for month
    print("\nMonths: 1=Jan, 2=Feb, 3=Mar, 4=Apr, 5=May, 6=Jun,")
    print("        7=Jul, 8=Aug, 9=Sep, 10=Oct, 11=Nov, 12=Dec")
    month_input = input(f"Enter month number [{default_month}]: ").strip()
    if month_input:
        try:
            report_month = int(month_input)
            if report_month < 1 or report_month > 12:
                print(f"Invalid month. Using default: {default_month}")
                report_month = default_month
        except ValueError:
            print(f"Invalid month. Using default: {default_month}")
            report_month = default_month
    else:
        report_month = default_month

    # Prompt for per diem start day (for partial months)
    start_day_input = input(f"Per diem start day [1]: ").strip()
    if start_day_input:
        try:
            start_day = int(start_day_input)
            _, max_day = calendar.monthrange(report_year, report_month)
            if start_day < 1 or start_day > max_day:
                print(f"Invalid day. Using default: 1")
                start_day = 1
        except ValueError:
            print(f"Invalid day. Using default: 1")
            start_day = 1
    else:
        start_day = 1

    print(f"\nGenerating report for: {calendar.month_name[report_month]} {report_year} (starting day {start_day})")
    return report_month, report_year, start_day


def main():
    """Main function to run the expense automation."""
    print("--- Starting Expense Report Automation ---")

    # 1. Determine the date range for the report (user prompted)
    report_month, report_year, per_diem_start_day = get_report_month_year()
    report_month_date = date(report_year, report_month, 1)

    # Scrape Per Diem and Currency Rates
    if config.DEBUG_MODE: print("\n--- Scraping Per Diem & Currency Rates ---")
    # Scrape rates for India
    per_diem_rates = utils.get_per_diem_rates_with_selenium(report_year, report_month, "India")
    # Also scrape rates for Sri Lanka (for Colombo trips)
    sri_lanka_rates = utils.get_per_diem_rates_with_selenium(report_year, report_month, "Sri Lanka")
    if sri_lanka_rates:
        # Mark Sri Lanka cities and merge into per_diem_rates
        for city, rates in sri_lanka_rates.items():
            rates["country"] = "Sri Lanka"
            per_diem_rates[city] = rates
    # Mark India cities
    for city in per_diem_rates:
        if "country" not in per_diem_rates[city]:
            per_diem_rates[city]["country"] = "India"

    mie_breakdown = utils.get_mie_breakdown()
    exchange_rates = utils.get_exchange_rates(report_month_date)
    usd_to_inr_rate = exchange_rates.get("INR")
    usd_to_lkr_rate = exchange_rates.get("LKR")

    if not per_diem_rates or not mie_breakdown or not usd_to_inr_rate:
        print("Could not retrieve per diem or currency rates. Exiting.")
        return

    if config.DEBUG_MODE:
        print(f"Exchange rates: USD to INR = {usd_to_inr_rate}, USD to LKR = {usd_to_lkr_rate}")

    # 2. Authenticate with Google Services
    creds = google_services.authenticate()
    if not creds:
        print("Failed to authenticate with Google. Exiting.")
        return
        
    gmail_service = build("gmail", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    calendar_service = build("calendar", "v3", credentials=creds)

    # 3. Search Gmail for travel confirmation emails
    # Look back 2 months for early bookings
    all_flights = []
    hotel_reservations = []
    travel_pdf_paths = []

    search_start_date = report_month_date - relativedelta(months=2)
    search_end_date = report_month_date + relativedelta(months=1)
    gmail_search_after = search_start_date.strftime('%Y/%m/%d')
    gmail_search_before = search_end_date.strftime('%Y/%m/%d')
    
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
            if filename and filename.endswith('.pdf'):
                pdf_path = google_services.get_gmail_attachment(gmail_service, msg_id, filename)
                if pdf_path:
                    # Try parsing as flight PDF
                    flights_in_pdf = utils.parse_flight_pdf(pdf_path)
                    has_relevant_flights = any(
                        f['departure'].month == report_month and f['departure'].year == report_year
                        for f in flights_in_pdf
                    )

                    # Try parsing as hotel reservation PDF
                    hotel_info = utils.parse_hotel_reservation_pdf(pdf_path)
                    has_relevant_hotel = False
                    if hotel_info and hotel_info.get("checkin_date"):
                        checkin = hotel_info["checkin_date"]
                        has_relevant_hotel = checkin.month == report_month and checkin.year == report_year

                    if has_relevant_flights:
                        all_flights.extend(flights_in_pdf)
                        if pdf_path not in travel_pdf_paths:
                            travel_pdf_paths.append(pdf_path)
                    elif has_relevant_hotel:
                        hotel_reservations.append(hotel_info)
                        if pdf_path not in travel_pdf_paths:
                            travel_pdf_paths.append(pdf_path)
                    else:
                        # Delete PDF that's not for the report month
                        os.remove(pdf_path)
                        if config.DEBUG_MODE:
                            print(f"  -> Skipped PDF (not for {calendar.month_name[report_month]} {report_year}): {pdf_path}")

    # Filter for flights within the report month and sort them
    relevant_flights = sorted([f for f in all_flights if f['departure'].month == report_month and f['departure'].year == report_year], key=lambda x: x['departure'])

    # 3b. Search Google Calendar for Bangalore company meetings
    if config.DEBUG_MODE: print("\n--- Searching Calendar for Bangalore Company Meetings ---")
    _, num_days_in_month = calendar.monthrange(report_year, report_month)
    month_start = date(report_year, report_month, 1)
    month_end = date(report_year, report_month, num_days_in_month)

    # Returns dict mapping date -> company name
    bangalore_meetings = google_services.search_calendar_events(
        calendar_service,
        config.BANGALORE_COMPANIES,
        month_start,
        month_end
    )
    if config.DEBUG_MODE: print(f"Found {len(bangalore_meetings)} Bangalore company meeting dates.")

    # Continue even if no flights or meetings - still generate per diem report
    if not relevant_flights and not bangalore_meetings:
        print("No travel bookings or Bangalore meetings found - generating per diem only report.")

    # 4. Create Travel Calendar to determine nightly location
    if config.DEBUG_MODE: print("\n--- Building Travel Calendar ---")
    travel_calendar = {}
    current_location = "Bangalore"

    flights_by_date = defaultdict(list)

    for f in relevant_flights:
        if 'date' in f and isinstance(f['date'], date):  # ensure key exists and is a date object
            flights_by_date[f['date']].append(f)

    unique_travel_dates = []
    for day_num in range(per_diem_start_day, num_days_in_month + 1):
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
    # Include both travel dates and Bangalore company meeting dates
    uber_search_dates = set(unique_travel_dates)
    for meeting_date in bangalore_meetings.keys():
        uber_search_dates.add(meeting_date)
    uber_search_dates = sorted(uber_search_dates)

    if config.DEBUG_MODE:
        print(f"\n--- Searching Uber Receipts for {len(uber_search_dates)} dates ---")
        if bangalore_meetings:
            print(f"  (includes {len(bangalore_meetings)} Bangalore company meeting dates)")

    uber_data = []
    uber_receipt_paths = []
    yahoo_mail = yahoo_service.connect_to_yahoo(config.YAHOO_EMAIL, config.YAHOO_APP_PASSWORD)
    if yahoo_mail:
        for search_date in uber_search_dates:
            receipts = yahoo_service.search_uber_receipts(yahoo_mail, search_date, usd_to_inr_rate)
            if receipts:
                for receipt_details in receipts:
                    receipt_details['date'] = search_date
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

    # 7. Create and populate Google Sheet (MATCH MARCH TEMPLATE)
    sheet_name = config.DRIVE_SHEET_NAME.format(month_name=report_month_date.strftime('%B'), year=report_year)

    # 👉 Copy the March template (keeps tabs/formatting/header row positions)
    spreadsheet_id = google_services.copy_and_convert_to_sheet(
        drive_service,
        config.TEMPLATE_SPREADSHEET_ID,
        sheet_name,
        folder_id
    )

    # Prepare Per Diem data
    per_diem_rows = []
    start_row_pd = 12       # matches March template
    row_counter = start_row_pd
    running_total_formula = f"=H{row_counter}"   # Running Total (col J) starts off equal to Total M&IE (col H)

    for day_num in range(per_diem_start_day, num_days_in_month + 1):
        current_date = date(report_year, report_month, day_num)
        location = travel_calendar.get(current_date, "Bangalore")

        bfast, lunch, dinner, incidentals, total_mie_rate, lodging = [""] * 6
        country = "India"  # Default country

        if "Bangalore" in location:
            rates = config.PER_DIEM_RATES_USD["Bangalore"]
            bfast, lunch, dinner, incidentals = rates["breakfast"], rates["lunch"], rates["dinner"], rates["incidentals"]
            total_mie_rate = rates.get("total_mie", "")
        else:
            city_rates = per_diem_rates.get(location, per_diem_rates.get("Other"))
            if city_rates:
                lodging = city_rates["lodging"]
                total_mie_rate = city_rates["total_mie"]
                country = city_rates.get("country", "India")
                breakdown = mie_breakdown.get(total_mie_rate, {})
                bfast = config.PER_DIEM_RATES_USD["Bangalore"]["breakfast"]
                lunch = breakdown.get("lunch", "")
                dinner = breakdown.get("dinner", "")
                incidentals = breakdown.get("incidentals", "")

        total_formula = f"=SUM(D{row_counter}:G{row_counter})"   # H = D+E+F+G

        per_diem_rows.append([
            current_date.strftime('%Y-%m-%d'),   # A: Date(s) Claimed:
            f"{location}, {country}",            # B
            total_mie_rate,                      # C: State Dept M&IE (Per Diem Rate)*
            bfast, lunch, dinner, incidentals,   # D-G
            total_formula,                       # H: Total M&IE for Date
            "",                                  # I: Lodging Cost For Night (left blank if N/A)
            running_total_formula,               # J: Running Total
            ""                                   # K: Comments
        ])
        row_counter += 1
        running_total_formula = f"=J{row_counter-1}+H{row_counter}"

    # Add the total row - label in column G, formulas in H and J
    per_diem_rows.append([
        "",                                       # A
        "",                                       # B
        "",                                       # C
        "",                                       # D
        "",                                       # E
        "",                                       # F
        "TOTAL PER DIEM",                         # G: Label
        f"=SUM(H{start_row_pd}:H{row_counter-1})", # H: Sum of all daily totals
        "",                                       # I
        f"=J{row_counter-1}",                     # J: Final running total
        ""                                        # K
    ])

    if config.SAVE_TO_DRIVE and per_diem_rows:
        # Clear old data BELOW the first data row to avoid trailing junk (keep header row 11 and first data row 12 safe)
        google_services.clear_values(sheets_service, spreadsheet_id, "Per Diem & Lodging!A12:K9999")
        # Write from the first data row (A12)
        google_services.update_values(sheets_service, spreadsheet_id, f"Per Diem & Lodging!A{start_row_pd}", per_diem_rows)

    # Prepare Reimbursements data
    reimbursement_rows = []
    start_row_rb = 13   # matches March template
    row_counter = start_row_rb

    for item in sorted(uber_data, key=lambda x: x['date']):
        # Get the travel city for this date to help with location classification
        travel_city = travel_calendar.get(item['date'], "Bangalore")
        expense_date = item['date']

        # Determine the company being coached for this expense
        # For Bangalore: must have a calendar appointment with a coaching company
        # For travel cities: look up company from config.COMPANIES
        if travel_city == "Bangalore" or "Bangalore" in travel_city:
            if expense_date in bangalore_meetings:
                company_name = bangalore_meetings[expense_date]
            else:
                # No Bangalore company meeting on this date - skip this expense
                if config.DEBUG_MODE:
                    print(f"  -> Skipping Bangalore expense on {expense_date}: no company meeting in calendar")
                continue
        else:
            # Non-Bangalore travel - look up company by city
            company_name = ""
            for city, companies in config.COMPANIES.items():
                if city.lower() in travel_city.lower() or travel_city.lower() in city.lower():
                    company_name = companies[0]
                    break

        # Generate descriptive ride description (e.g., "Home to Airport", "Taj Samudra to Airport")
        description = utils.generate_uber_description(
            item.get('from', ''),
            item.get('to', ''),
            travel_city,
            hotel_reservations
        )
        # Use the correct currency and exchange rate based on what was detected in the receipt
        currency = item.get('currency', 'INR')
        if currency == 'LKR' and usd_to_lkr_rate:
            exchange_rate = usd_to_lkr_rate
        else:
            exchange_rate = usd_to_inr_rate
            currency = 'INR'  # Default to INR if LKR rate not available

        reimbursement_rows.append([
            item['date'].strftime('%Y-%m-%d'),   # A: Expenditure Date (When):
            item['filepath'] or "",              # B: Receipt # *
            item['fare-city'] or "N/A",          # C: Location (Where)
            currency,                            # D: Currency (INR or LKR)
            description,                         # E: Description
            item.get('fare', 'N/A'),             # F: Receipt Amt in Receipt Currency
            exchange_rate,                       # G: Rate of Exchange (USD to currency)
            f"=F{row_counter}/G{row_counter}",   # H: US Dollar Equivalent
            company_name                         # I: Company being coached
        ])
        row_counter += 1

    # Add the total row - label in column G, formula in H
    reimbursement_rows.append([
        "",                                        # A
        "",                                        # B
        "",                                        # C
        "",                                        # D
        "",                                        # E
        "",                                        # F
        "TOTAL REIMBURSEMENT",                     # G: Label (singular to match template)
        f"=SUM(H{start_row_rb}:H{row_counter-1})", # H: Sum of all USD equivalents
        ""                                         # I
    ])

    if config.SAVE_TO_DRIVE and reimbursement_rows:
        # Clear existing data below first data row (keep header at row 12, first data row 13 will be overwritten)
        google_services.clear_values(sheets_service, spreadsheet_id, "Reimbursements!A13:I9999")
        # Write starting at first data row (A13)
        google_services.update_values(sheets_service, spreadsheet_id, f"Reimbursements!A{start_row_rb}", reimbursement_rows)

    print("\n--- Expense Report Automation Finished Successfully! ---")
    if config.SAVE_TO_DRIVE:
        print(f"Your report has been saved to Google Drive in the folder '{drive_folder_name}'.")

if __name__ == "__main__":
    main()
