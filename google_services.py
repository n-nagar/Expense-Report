# google_services.py
# This module handles all interactions with Google APIs: Authentication, Gmail, Drive, and Sheets.

import os
import base64
import config as Config
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

# Scopes define the permissions the script will request from the user.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar.readonly"
]


def authenticate():
    """
    Handles user authentication for Google APIs.
    Creates a 'token.json' file to store access and refresh tokens.
    Automatically re-prompts if refresh token is expired, revoked, or invalid.
    Deletes token.json if it's no longer usable.
    """
    creds = None

    # Load existing credentials if they exist
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If no valid credentials, try refresh or full login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"⚠ Refresh token failed: {e}")
                # Delete bad token file
                if os.path.exists("token.json"):
                    os.remove("token.json")
                creds = None  # Force full re-authentication

        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

        # Save new credentials
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds


def search_gmail(service, query):
    """Searches Gmail for emails matching the given query."""
    try:
        response = service.users().messages().list(userId="me", q=query).execute()
        messages = response.get("messages", [])
        return messages
    except HttpError as error:
        print(f"An error occurred with Gmail search: {error}")
        return []

def get_gmail_attachment(service, msg_id, attachment_filename):
    """Downloads a specific attachment from a Gmail message."""
    try:
        message = service.users().messages().get(userId="me", id=msg_id).execute()
        parts_to_search = message["payload"].get("parts", [])
        
        queue = list(parts_to_search)
        while queue:
            part = queue.pop(0)
            if part.get("parts"):
                queue.extend(part.get("parts"))
            
            if part.get("filename") and part["filename"] == attachment_filename:
                if "data" in part["body"]:
                    data = part["body"]["data"]
                else:
                    att_id = part["body"]["attachmentId"]
                    att = service.users().messages().attachments().get(userId="me", messageId=msg_id, id=att_id).execute()
                    data = att["data"]

                file_data = base64.urlsafe_b64decode(data.encode("UTF-8"))
                with open(attachment_filename, "wb") as f:
                    f.write(file_data)
                return attachment_filename
    except HttpError as error:
        print(f"An error occurred while downloading attachment: {error}")
    return None

def create_drive_folder(service, folder_name):
    """Creates a folder in Google Drive if it doesn't already exist."""
    try:
        query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
        response = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        if response.get('files'):
            print(f"Folder '{folder_name}' already exists.")
            return response.get('files')[0].get('id')

        file_metadata = {"name": folder_name, "mimeType": "application/vnd.google-apps.folder"}
        folder = service.files().create(body=file_metadata, fields="id").execute()
        if Config.DEBUG_MODE: print(f"Created Google Drive folder: '{folder_name}'")
        return folder.get("id")
    except HttpError as error:
        print(f"An error occurred while creating Drive folder: {error}")
        return None

def upload_file_to_drive(service, file_path, folder_id):
    """Uploads a local file to a specified Google Drive folder."""
    try:
        file_metadata = {"name": os.path.basename(file_path), "parents": [folder_id]}
        media = MediaFileUpload(file_path, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
        if Config.DEBUG_MODE: print(f"Uploaded file '{os.path.basename(file_path)}' to Drive.")
        return file.get("id")
    except HttpError as error:
        print(f"An error occurred while uploading file: {error}")
    except Exception as e:
        print(f"A local file error occurred: {e}")

def create_google_sheet(drive_service, sheet_name, folder_id):
    """Creates a new Google Sheet in a specified Drive folder."""
    try:
        file_metadata = {
            "name": sheet_name,
            "parents": [folder_id],
            "mimeType": "application/vnd.google-apps.spreadsheet",
        }
        sheet = drive_service.files().create(body=file_metadata, fields="id").execute()
        if Config.DEBUG_MODE: print(f"Created Google Sheet: '{sheet_name}'")
        return sheet.get("id")
    except HttpError as error:
        print(f"An error occurred while creating Google Sheet: {error}")
        return None

def setup_spreadsheet_tabs(sheets_service, spreadsheet_id, tab_configs):
    """
    Sets up a spreadsheet by renaming the first tab and creating new tabs with headers.
    """
    requests = []
    
    # Request to rename the default "Sheet1" to the first tab's name
    requests.append({
        "updateSheetProperties": {
            "properties": {"sheetId": 0, "title": tab_configs[0]["name"]},
            "fields": "title"
        }
    })

    # Requests to add the rest of the tabs
    for config in tab_configs[1:]:
        requests.append({"addSheet": {"properties": {"title": config["name"]}}})

    try:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()

        if Config.DEBUG_MODE: print(f"Successfully created tabs: {[c['name'] for c in tab_configs]}")
        
        # Now, add the headers to each tab
        for config in tab_configs:
            append_values(sheets_service, spreadsheet_id, config["name"], [config["headers"]])

    except HttpError as error:
        print(f"An error occurred setting up spreadsheet tabs: {error}")

def append_values(sheets_service, spreadsheet_id, range_name, values):
    """Appends values to a sheet."""
    try:
        body = {"values": values}
        # CORRECTED: The range for an append operation should just be the sheet name.
        # This makes the request less ambiguous and prevents the API from duplicating rows.
        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="USER_ENTERED",
            body=body,
        ).execute()
        print(f"Successfully wrote {len(values)} row(s) to tab '{range_name}'.")
    except HttpError as error:
        print(f"An error occurred appending values to sheet: {error}")


def copy_google_sheet(drive_service, template_file_id: str, name: str, folder_id: str) -> str:
    """Copies a Google Sheet template to a new location."""
    body = {
        "name": name,
        "parents": [folder_id]
    }
    copied = drive_service.files().copy(
        fileId=template_file_id,
        body=body
    ).execute()
    return copied["id"]

def copy_and_convert_to_sheet(drive_service, template_file_id: str, name: str, folder_id: str) -> str:
    """
    Copies a file (e.g., .xlsx, .csv) and converts it into a native Google Sheet.
    
    Args:
        drive_service: The authenticated Google Drive API service object.
        template_file_id: The ID of the source file to copy (e.g., an Excel file).
        name: The name for the new Google Sheet.
        folder_id: The ID of the folder where the new sheet will be created.

    Returns:
        The file ID of the newly created Google Sheet.
    """
    body = {
        "name": name,
        "parents": [folder_id],
        # This is the key that forces the conversion to a Google Sheet ✨
        "mimeType": "application/vnd.google-apps.spreadsheet"
    }
    
    copied_sheet = drive_service.files().copy(
        fileId=template_file_id,
        body=body
    ).execute()
    
    print(f"Successfully created Google Sheet '{name}' with ID: {copied_sheet['id']}")
    return copied_sheet["id"]

def clear_values(sheets_service, spreadsheet_id: str, a1_range: str):
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=a1_range,
        body={}
    ).execute()

def update_values(sheets_service, spreadsheet_id: str, a1_range: str, values: list[list[str | float]]):
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=a1_range,
        valueInputOption="USER_ENTERED",   # keep formulas working
        body={"values": values}
    ).execute()


def search_calendar_events(calendar_service, company_names: list, start_date, end_date):
    """
    Searches Google Calendar for events matching any of the company names.
    Returns a list of dates where matching events are found.

    Args:
        calendar_service: The authenticated Google Calendar API service object.
        company_names: List of company names to search for in event titles.
        start_date: Start date for the search range (datetime.date).
        end_date: End date for the search range (datetime.date).

    Returns:
        List of datetime.date objects where matching calendar events exist.
    """
    from datetime import datetime, timedelta

    # Convert dates to RFC3339 format for Calendar API
    time_min = datetime.combine(start_date, datetime.min.time()).isoformat() + 'Z'
    time_max = datetime.combine(end_date + timedelta(days=1), datetime.min.time()).isoformat() + 'Z'

    matching_dates = []

    try:
        events_result = calendar_service.events().list(
            calendarId='primary',
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime'
        ).execute()

        events = events_result.get('items', [])

        for event in events:
            event_summary = event.get('summary', '').lower()

            # Check if any company name matches the event title
            for company in company_names:
                if company.lower() in event_summary:
                    # Extract the date from the event
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    if 'T' in start:
                        # DateTime format: 2026-01-15T10:00:00+05:30
                        event_date = datetime.fromisoformat(start.replace('Z', '+00:00')).date()
                    else:
                        # All-day event: 2026-01-15
                        event_date = datetime.strptime(start, '%Y-%m-%d').date()

                    if event_date not in matching_dates:
                        matching_dates.append(event_date)
                        if Config.DEBUG_MODE:
                            print(f"  -> Found calendar event: '{event.get('summary')}' on {event_date}")
                    break

        return sorted(matching_dates)

    except HttpError as error:
        print(f"An error occurred while searching calendar: {error}")
        return []

if __name__ == "__main__":
    creds = authenticate()
    if not creds:
        print("Failed to authenticate with Google. Exiting.")
        exit(1)
        
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    sheet_name = "HiHiHiHi"

    # 👉 Copy the March template (keeps tabs/formatting/header row positions)
    drive_folder_name = "TestTestTest"
    folder_id = create_drive_folder(drive_service, drive_folder_name)    
    spreadsheet_id = copy_and_convert_to_sheet(
        drive_service,
        Config.TEMPLATE_SPREADSHEET_ID,
        sheet_name,
        folder_id
    )
    if not spreadsheet_id:
        print("Failed to create Google Sheet. Exiting.")
        exit(1)
    clear_values(sheets_service, spreadsheet_id, "Per Diem & Lodging!A12:K9999")

