# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Expense report automation for Stanford Seed program coaching. The user is a volunteer coach for Stanford University GSB's Seed program in India, traveling monthly to meet with 8 companies:
- 3 in Bangalore (local, home base)
- 2 in Vadodara, India
- 1 in Colombo, Sri Lanka
- 1 in Indore, India
- 1 in Siliguri, India

Travel is booked by Stanford agents who send confirmations to a Stanford University Gmail account, establishing travel start/end dates. This system consolidates flight confirmations and Uber receipts, then generates a Google Drive folder with all details and receipts formatted according to Stanford's expense report rules and templates.

## Running the Application

```bash
python main.py
```

The application will prompt for the report month and year (defaults to previous month). This allows generating reports for any past month, not just the most recent one.

## Running Tests

Run all integration tests (verifies email access, template access, API connectivity):
```bash
python test_integrations.py
```

Or with pytest for more detailed output:
```bash
python -m pytest test_integrations.py -v
```

Test the per diem scraper separately:
```bash
python test_scraper.py
```

## Architecture

```
main.py                 # Orchestration - runs the complete workflow
├── config.py           # Credentials, constants, MIE breakdown table
├── google_services.py  # OAuth2 + Gmail/Drive/Sheets API operations
├── yahoo_service.py    # IMAP connection to Yahoo Mail for Uber receipts
├── utils.py            # PDF/HTML parsing, web scraping, currency conversion
└── html2pdf.py         # Standalone HTML-to-PDF via Chrome headless
```

## Data Flow

1. **Prompt for report month/year** (defaults to previous month)
2. **Scrape reference data**:
   - US State Department per diem rates for India cities (lodging + M&IE totals) via Selenium
   - MIE breakdown table maps total M&IE to individual meal/incidental amounts
   - USD-to-INR exchange rate for the month (currently Frankfurter API; OANDA scraper available as fallback)
3. **Extract flights** - Search Gmail for travel PDFs (looks 2 months back for early bookings), parse with pdfplumber
4. **Search calendar for Bangalore meetings** - Find events matching Bangalore company names (local trips without flights)
5. **Build travel calendar** - Map each day to a location based on flight sequence
6. **Find Uber receipts** - Search Yahoo Mail via IMAP for travel dates AND Bangalore meeting dates
7. **Generate output** - Copy Google Sheet template, populate with data and formulas, upload PDFs to Drive

## Per Diem Business Rules

**Breakfast is always $12** - Fixed by contract regardless of location or State Department rates. This is set in `config.PER_DIEM_RATES_USD["Bangalore"]["breakfast"]` and applied to all days.

**For Bangalore days (home base):**
- Uses fixed rates from `config.PER_DIEM_RATES_USD["Bangalore"]` (breakfast, lunch, dinner, incidentals)
- No lodging claimed

**For travel days (other cities):**
- Lodging and total M&IE scraped from US State Department for that city
- Breakfast uses the fixed $12 rate
- Lunch, dinner, incidentals calculated from `MIE_BREAKDOWN` table based on total M&IE rate

## Currency Conversion

All Uber charges are in INR. The system fetches the USD-to-INR rate for the report month and converts receipts to USD equivalents in the spreadsheet. Only receipts exceeding $10 USD equivalent are saved.

## Key External Dependencies

- **Google APIs**: Gmail (read), Drive (create folders/upload), Sheets (read/write), Calendar (read)
- **Yahoo IMAP**: Requires App Password in config.py
- **Selenium + Chrome**: For scraping state.gov per diem tables and HTML-to-PDF conversion
- **Frankfurter API**: Daily exchange rates

## Google Sheets Structure

The template spreadsheet has two sheets:
- **"Per Diem & Lodging"** - Data starts row 12 (headers row 11)
- **"Reimbursements"** - Data starts row 13 (headers row 12)

Both sheets use formula injection for calculations - values are written as spreadsheet formulas, not computed values.

## Configuration (config.py)

Required settings:
- `YAHOO_EMAIL`, `YAHOO_APP_PASSWORD` - Yahoo Mail access
- `TRAVEL_EMAIL_SENDER` - Gmail sender for flight confirmations
- `TEMPLATE_SPREADSHEET_ID` - Google Sheet template to copy
- `PER_DIEM_RATES_USD` - Bangalore per diem rates
- `MIE_BREAKDOWN` - 265-entry dict mapping day number to meal/incidental amounts
- `BANGALORE_COMPANIES` - List of local company names to search in calendar (Amuse Labs, Axcend Systems, Padmini Aromatics)

Flags:
- `DEBUG_MODE` - Verbose output
- `SAVE_TO_DRIVE` - Toggle Drive uploads

## Authentication Setup

**Google (two accounts involved):**
- `credentials.json` - OAuth client credentials from Google Cloud Console (set up under personal Google account)
- `token.json` - OAuth token for the authenticated session (Stanford University Gmail account)
- When running, the OAuth flow prompts login to the Stanford account to access flight confirmation emails

**Yahoo Mail:**
- Personal Yahoo account where Uber receipts arrive
- Uses IMAP with App Password configured in `config.py`

Both credential files are gitignored.

## Important Patterns

- **Duplicate detection**: Uber receipts compared by fare/date/from/to fields
- **Fare filtering**: Only receipts >$10 USD equivalent are saved
- **Multiple flights/day**: Uses last flight of the day for location
- **Default location**: Bangalore when no flights recorded
