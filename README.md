Automated Expense Reporter
This Python project automates the process of creating monthly expense reports by fetching flight and Uber details from your email accounts and organizing them in Google Drive.

Features
Connects to Gmail: Searches for flight booking confirmation emails from a specific sender.

Parses Attachments: Extracts travel dates and destinations from PDF attachment filenames.

Connects to Yahoo Mail: Uses the extracted travel dates to find corresponding Uber receipts.

Extracts Receipt Data: Parses Uber receipt emails to get trip details and fare information.

Organizes in Google Drive:

Creates a new folder for the expense month (e.g., 07-2025).

Creates a Google Sheet with separate tabs for 'Travel' and 'Uber' expenses.

Downloads and saves flight PDFs and Uber receipt emails to the Drive folder.

Prerequisites
Python 3.8 or newer.

A Google account with a configured Google Cloud project.

A Yahoo Mail account.

Setup Instructions
Follow these steps carefully to get the project running.

1. Clone the Repository

First, clone this project to your local machine using Git.

git clone <your-github-repository-url>
cd <repository-name>

2. Install Dependencies

Install all the required Python libraries using the requirements.txt file.

pip install -r requirements.txt

3. Configure Google Cloud Project

Since you cannot use your company's workspace, you'll use your personal Google account to set up the necessary APIs.

Create a Project: Go to the Google Cloud Console and create a new project.

Enable APIs: In your new project, go to the "APIs & Services" > "Library" section and enable the following APIs:

Gmail API

Google Drive API

Google Sheets API

Create OAuth 2.0 Credentials:

Go to "APIs & Services" > "Credentials".

Click "Create Credentials" and select "OAuth client ID".

If prompted, configure the consent screen. Choose "External" and provide an app name, user support email, and developer contact information.

For the application type, select "Desktop app".

After creation, a pop-up will appear. Click "DOWNLOAD JSON".

Rename the downloaded file to credentials.json and place it in the root directory of this project.

4. Configure Yahoo Mail for IMAP Access

To allow the script to access your Yahoo Mail securely, you must generate an "App Password".

Log in to your Yahoo account.

Go to your "Account Info" > "Account Security" page.

Find "App password" and click "Generate app password".

Select "Other App", give it a name (e.g., "Python Expense Script"), and click "Generate".

Yahoo will provide a 16-character password. Copy this password immediately. You will use this in the config.py file, not your regular Yahoo password.

5. Update the Configuration File

Open the config.py file and fill in your details.

YAHOO_EMAIL: Your full Yahoo email address.

YAHOO_APP_PASSWORD: The 16-character app password you just generated.

TRAVEL_EMAIL_SENDER: The email address of the travel agency.

6. Run the Application

Once everything is set up, you can run the script from your terminal:

python main.py

First Run: The first time you run the script, a new browser window or tab will open, asking you to authorize access to your Google Account. Please log in and grant the requested permissions. The script will then create a token.json file to store your authorization, so you won't have to do this again.

Subsequent Runs: The script will use the token.json file to automatically refresh your access.

The script will print its progress in the terminal and will notify you upon successful completion.