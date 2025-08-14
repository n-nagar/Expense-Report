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
    "https://www.googleapis.com/auth/spreadsheets"
]

def authenticate():
    """
    Handles user authentication for Google APIs.
    Creates a 'token.json' file to store access and refresh tokens.
    """
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
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
        print(f"Created Google Drive folder: '{folder_name}'")
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
        print(f"Uploaded file '{os.path.basename(file_path)}' to Drive.")
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
        print(f"Created Google Sheet: '{sheet_name}'")
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
        print(f"Successfully created tabs: {[c['name'] for c in tab_configs]}")
        
        # Now, add the headers to each tab
        for config in tab_configs:
            append_values(sheets_service, spreadsheet_id, config["name"], [config["headers"]])

    except HttpError as error:
        print(f"An error occurred setting up spreadsheet tabs: {error}")

def append_values(sheets_service, spreadsheet_id, range_name, values):
    """Appends values to a sheet."""
    try:
        body = {"values": values}
        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{range_name}!A1",
            valueInputOption="USER_ENTERED",
            body=body,
        ).execute()
        print(f"Successfully wrote {len(values)} row(s) to tab '{range_name}'.")
    except HttpError as error:
        print(f"An error occurred appending values to sheet: {error}")
