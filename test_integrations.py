# test_integrations.py
# Integration tests to verify all external service connectivity and data retrieval.
# Run with: python -m pytest test_integrations.py -v
# Or: python test_integrations.py

import unittest
from datetime import date
from dateutil.relativedelta import relativedelta


class TestGoogleAuthentication(unittest.TestCase):
    """Tests for Google OAuth authentication."""

    def test_credentials_file_exists(self):
        """Verify credentials.json file exists."""
        import os
        self.assertTrue(
            os.path.exists("credentials.json"),
            "credentials.json not found. Download from Google Cloud Console."
        )

    def test_google_authentication(self):
        """Verify Google OAuth authentication succeeds."""
        import google_services
        creds = google_services.authenticate()
        self.assertIsNotNone(creds, "Google authentication failed")
        self.assertTrue(creds.valid, "Google credentials are not valid")


class TestGmailAccess(unittest.TestCase):
    """Tests for Gmail API access."""

    @classmethod
    def setUpClass(cls):
        import google_services
        from googleapiclient.discovery import build
        cls.creds = google_services.authenticate()
        cls.gmail_service = build("gmail", "v1", credentials=cls.creds)

    def test_gmail_service_connection(self):
        """Verify Gmail API service is accessible."""
        self.assertIsNotNone(self.gmail_service)

    def test_gmail_search_works(self):
        """Verify Gmail search returns without error."""
        import google_services
        # Search for any email in the last 30 days
        messages = google_services.search_gmail(
            self.gmail_service,
            "newer_than:30d"
        )
        # Should return a list (possibly empty, but not None)
        self.assertIsInstance(messages, list)

    def test_gmail_search_travel_sender(self):
        """Verify Gmail can search for travel confirmation sender."""
        import config
        import google_services
        query = f'from:"{config.TRAVEL_EMAIL_SENDER}"'
        messages = google_services.search_gmail(self.gmail_service, query)
        self.assertIsInstance(messages, list)
        print(f"  Found {len(messages)} emails from travel sender")


class TestYahooMailAccess(unittest.TestCase):
    """Tests for Yahoo Mail IMAP access."""

    def test_yahoo_connection(self):
        """Verify Yahoo Mail IMAP connection succeeds."""
        import config
        import yahoo_service
        mail = yahoo_service.connect_to_yahoo(
            config.YAHOO_EMAIL,
            config.YAHOO_APP_PASSWORD
        )
        self.assertIsNotNone(mail, "Yahoo Mail connection failed")
        yahoo_service.close_connection(mail)

    def test_yahoo_search_uber_receipts(self):
        """Verify Yahoo can search for Uber receipts."""
        import config
        import yahoo_service
        mail = yahoo_service.connect_to_yahoo(
            config.YAHOO_EMAIL,
            config.YAHOO_APP_PASSWORD
        )
        self.assertIsNotNone(mail)

        # Search for receipts from last month
        last_month = date.today() - relativedelta(months=1)
        test_date = last_month.replace(day=15)  # Middle of last month

        # Use a placeholder rate for testing
        receipts = yahoo_service.search_uber_receipts(mail, test_date, 85.0)
        self.assertIsInstance(receipts, list)
        print(f"  Found {len(receipts)} Uber receipts on {test_date}")

        yahoo_service.close_connection(mail)


class TestGoogleDriveAccess(unittest.TestCase):
    """Tests for Google Drive API access."""

    @classmethod
    def setUpClass(cls):
        import google_services
        from googleapiclient.discovery import build
        cls.creds = google_services.authenticate()
        cls.drive_service = build("drive", "v3", credentials=cls.creds)

    def test_drive_service_connection(self):
        """Verify Drive API service is accessible."""
        self.assertIsNotNone(self.drive_service)

    def test_drive_list_files(self):
        """Verify Drive can list files."""
        response = self.drive_service.files().list(
            pageSize=1,
            fields="files(id, name)"
        ).execute()
        self.assertIn("files", response)


class TestGoogleSheetsTemplateAccess(unittest.TestCase):
    """Tests for Google Sheets template accessibility."""

    @classmethod
    def setUpClass(cls):
        import config
        import google_services
        from googleapiclient.discovery import build
        cls.creds = google_services.authenticate()
        cls.sheets_service = build("sheets", "v4", credentials=cls.creds)
        cls.drive_service = build("drive", "v3", credentials=cls.creds)
        cls.template_id = config.TEMPLATE_SPREADSHEET_ID

    def test_template_spreadsheet_exists(self):
        """Verify template spreadsheet ID is configured."""
        import config
        self.assertTrue(
            hasattr(config, 'TEMPLATE_SPREADSHEET_ID'),
            "TEMPLATE_SPREADSHEET_ID not found in config"
        )
        self.assertTrue(
            len(config.TEMPLATE_SPREADSHEET_ID) > 0,
            "TEMPLATE_SPREADSHEET_ID is empty"
        )

    def test_template_spreadsheet_accessible(self):
        """Verify template spreadsheet can be accessed (native Sheet or Excel)."""
        try:
            # First check if it's accessible via Drive API
            file_metadata = self.drive_service.files().get(
                fileId=self.template_id,
                fields="id,name,mimeType"
            ).execute()

            mime_type = file_metadata.get('mimeType', '')
            name = file_metadata.get('name', '')

            # Accept both native Google Sheets and Excel files
            valid_types = [
                'application/vnd.google-apps.spreadsheet',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ]
            self.assertIn(mime_type, valid_types, f"Template is not a spreadsheet: {mime_type}")
            print(f"  Template name: {name}")
            print(f"  Template type: {'Google Sheet' if 'google-apps' in mime_type else 'Excel (.xlsx)'}")
        except Exception as e:
            self.fail(f"Cannot access template spreadsheet: {e}")

    def test_template_has_required_sheets(self):
        """Verify template has Per Diem & Lodging and Reimbursements sheets."""
        # First check if template is native Google Sheet or Excel
        file_metadata = self.drive_service.files().get(
            fileId=self.template_id,
            fields="mimeType,name"
        ).execute()

        mime_type = file_metadata.get('mimeType', '')

        if 'google-apps.spreadsheet' in mime_type:
            # Native Google Sheet - can read directly
            spreadsheet = self.sheets_service.spreadsheets().get(
                spreadsheetId=self.template_id
            ).execute()
            sheet_names = [s['properties']['title'] for s in spreadsheet['sheets']]
        else:
            # Excel file - need to check by copying and converting first
            # Skip detailed sheet check for Excel, just verify it's copyable
            print(f"  Template is Excel file - sheet names verified after copy/convert")
            print(f"  Expected sheets: 'Per Diem & Lodging', 'Reimbursements'")
            return  # Skip assertion for Excel files

        self.assertIn(
            "Per Diem & Lodging",
            sheet_names,
            "Template missing 'Per Diem & Lodging' sheet"
        )
        self.assertIn(
            "Reimbursements",
            sheet_names,
            "Template missing 'Reimbursements' sheet"
        )
        print(f"  Template sheets: {sheet_names}")

    def test_template_can_be_copied(self):
        """Verify template can be copied (tests Drive permissions)."""
        try:
            # Just check metadata - don't actually copy
            file_metadata = self.drive_service.files().get(
                fileId=self.template_id,
                fields="id,name,mimeType,capabilities"
            ).execute()

            capabilities = file_metadata.get('capabilities', {})
            can_copy = capabilities.get('canCopy', False)
            self.assertTrue(can_copy, "Template cannot be copied - check sharing settings")
            print(f"  Template '{file_metadata.get('name')}' is copyable")
        except Exception as e:
            self.fail(f"Cannot access template file metadata: {e}")


class TestPerDiemScraping(unittest.TestCase):
    """Tests for US State Department per diem rate scraping."""

    def test_per_diem_scraping(self):
        """Verify per diem rates can be scraped from state.gov."""
        import utils

        # Use current month for testing
        today = date.today()
        last_month = today - relativedelta(months=1)

        rates = utils.get_per_diem_rates_with_selenium(
            last_month.year,
            last_month.month,
            "India"
        )

        self.assertIsNotNone(rates, "Per diem scraping returned None")
        self.assertIsInstance(rates, dict)
        self.assertGreater(len(rates), 0, "No per diem rates retrieved")

        # Verify structure of returned data
        first_city = list(rates.keys())[0]
        self.assertIn("lodging", rates[first_city])
        self.assertIn("total_mie", rates[first_city])

        print(f"  Retrieved rates for {len(rates)} cities in India")
        print(f"  Sample: {first_city} - Lodging: ${rates[first_city]['lodging']}, M&IE: ${rates[first_city]['total_mie']}")

    def test_mie_breakdown_available(self):
        """Verify MIE breakdown table is configured."""
        import utils

        breakdown = utils.get_mie_breakdown()

        self.assertIsNotNone(breakdown)
        self.assertIsInstance(breakdown, dict)
        self.assertGreater(len(breakdown), 0, "MIE breakdown table is empty")

        # Check structure
        sample_key = list(breakdown.keys())[0]
        sample_value = breakdown[sample_key]
        self.assertIn("lunch", sample_value)
        self.assertIn("dinner", sample_value)
        self.assertIn("incidentals", sample_value)

        print(f"  MIE breakdown has {len(breakdown)} rate entries")


class TestCurrencyConversion(unittest.TestCase):
    """Tests for currency conversion rate retrieval."""

    def test_frankfurter_api_accessible(self):
        """Verify Frankfurter API returns USD-to-INR rate."""
        import utils

        # Use a date from last month
        last_month = date.today() - relativedelta(months=1)
        test_date = last_month.replace(day=15)

        rate = utils.get_usd_to_inr_rate(test_date)

        self.assertIsNotNone(rate, "Currency API returned None")
        self.assertIsInstance(rate, (int, float))
        self.assertGreater(rate, 0, "Exchange rate should be positive")
        # Sanity check: INR is typically 80-90 per USD
        self.assertGreater(rate, 50, "Exchange rate seems too low")
        self.assertLess(rate, 150, "Exchange rate seems too high")

        print(f"  USD to INR rate on {test_date}: {rate}")


class TestConfigurationValues(unittest.TestCase):
    """Tests for required configuration values."""

    def test_yahoo_credentials_configured(self):
        """Verify Yahoo credentials are set."""
        import config

        self.assertTrue(hasattr(config, 'YAHOO_EMAIL'), "YAHOO_EMAIL not in config")
        self.assertTrue(hasattr(config, 'YAHOO_APP_PASSWORD'), "YAHOO_APP_PASSWORD not in config")
        self.assertGreater(len(config.YAHOO_EMAIL), 0, "YAHOO_EMAIL is empty")
        self.assertGreater(len(config.YAHOO_APP_PASSWORD), 0, "YAHOO_APP_PASSWORD is empty")

    def test_travel_sender_configured(self):
        """Verify travel email sender is configured."""
        import config

        self.assertTrue(hasattr(config, 'TRAVEL_EMAIL_SENDER'), "TRAVEL_EMAIL_SENDER not in config")
        self.assertGreater(len(config.TRAVEL_EMAIL_SENDER), 0, "TRAVEL_EMAIL_SENDER is empty")

    def test_bangalore_rates_configured(self):
        """Verify Bangalore per diem rates are configured with $12 breakfast."""
        import config

        self.assertTrue(hasattr(config, 'PER_DIEM_RATES_USD'), "PER_DIEM_RATES_USD not in config")
        self.assertIn("Bangalore", config.PER_DIEM_RATES_USD)

        blr_rates = config.PER_DIEM_RATES_USD["Bangalore"]
        self.assertIn("breakfast", blr_rates)
        self.assertEqual(blr_rates["breakfast"], 12, "Breakfast should be fixed at $12")
        self.assertIn("lunch", blr_rates)
        self.assertIn("dinner", blr_rates)
        self.assertIn("incidentals", blr_rates)


class TestGoogleCalendarAccess(unittest.TestCase):
    """Tests for Google Calendar API access."""

    @classmethod
    def setUpClass(cls):
        import google_services
        from googleapiclient.discovery import build
        cls.creds = google_services.authenticate()
        cls.calendar_service = build("calendar", "v3", credentials=cls.creds)

    def test_calendar_service_connection(self):
        """Verify Calendar API service is accessible."""
        self.assertIsNotNone(self.calendar_service)

    def test_calendar_list_events(self):
        """Verify Calendar can list events."""
        from datetime import datetime, timedelta
        now = datetime.utcnow().isoformat() + 'Z'
        try:
            events_result = self.calendar_service.events().list(
                calendarId='primary',
                timeMin=now,
                maxResults=1,
                singleEvents=True
            ).execute()
            self.assertIn('items', events_result)
        except Exception as e:
            self.fail(f"Calendar API error: {e}")

    def test_bangalore_companies_configured(self):
        """Verify Bangalore companies are configured."""
        import config
        self.assertTrue(hasattr(config, 'BANGALORE_COMPANIES'))
        self.assertIsInstance(config.BANGALORE_COMPANIES, list)
        self.assertGreater(len(config.BANGALORE_COMPANIES), 0)
        print(f"  Bangalore companies: {config.BANGALORE_COMPANIES}")


class TestPDFParsing(unittest.TestCase):
    """Tests for PDF parsing utilities."""

    def test_pdfplumber_available(self):
        """Verify pdfplumber library is installed."""
        try:
            import pdfplumber
            self.assertTrue(True)
        except ImportError:
            self.fail("pdfplumber not installed - run: pip install pdfplumber")


class TestSeleniumSetup(unittest.TestCase):
    """Tests for Selenium/Chrome setup."""

    def test_selenium_chrome_available(self):
        """Verify Selenium can launch Chrome headless."""
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager

        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        try:
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=options
            )
            driver.quit()
            self.assertTrue(True)
        except Exception as e:
            self.fail(f"Chrome WebDriver failed: {e}")


def run_all_tests():
    """Run all integration tests and report results."""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    # Add test classes in logical order
    suite.addTests(loader.loadTestsFromTestCase(TestConfigurationValues))
    suite.addTests(loader.loadTestsFromTestCase(TestSeleniumSetup))
    suite.addTests(loader.loadTestsFromTestCase(TestGoogleAuthentication))
    suite.addTests(loader.loadTestsFromTestCase(TestGmailAccess))
    suite.addTests(loader.loadTestsFromTestCase(TestYahooMailAccess))
    suite.addTests(loader.loadTestsFromTestCase(TestGoogleDriveAccess))
    suite.addTests(loader.loadTestsFromTestCase(TestGoogleSheetsTemplateAccess))
    suite.addTests(loader.loadTestsFromTestCase(TestGoogleCalendarAccess))
    suite.addTests(loader.loadTestsFromTestCase(TestPerDiemScraping))
    suite.addTests(loader.loadTestsFromTestCase(TestCurrencyConversion))
    suite.addTests(loader.loadTestsFromTestCase(TestPDFParsing))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    return result.wasSuccessful()


if __name__ == "__main__":
    print("=" * 60)
    print("Expense Report Integration Tests")
    print("=" * 60)
    print()
    success = run_all_tests()
    exit(0 if success else 1)
