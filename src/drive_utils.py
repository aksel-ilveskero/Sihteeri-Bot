import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError

from ascii import print_logo

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/documents", "https://www.googleapis.com/auth/spreadsheets"]

def login() -> Resource:
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("credentials/token.json"):
    creds = Credentials.from_authorized_user_file("credentials/token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials/credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("credentials/token.json", "w") as token:
      token.write(creds.to_json())

  try:
    drive_service = build("drive", "v3", credentials=creds)
    sheet_service = build("sheets", "v4", credentials=creds)
    doc_service = build("docs", "v1", credentials=creds)
    user_info = drive_service.about().get(fields="user").execute()
    print_logo()
    print(f"Kirjauduttu Driveen käyttäjällä {user_info["user"]["displayName"]}, {user_info["user"]["emailAddress"]}\n")

    print(type(drive_service))
    
    return drive_service, sheet_service, doc_service

  except HttpError as error:
    # TODO(developer) - Handle errors from drive API.
    print(f"An error occurred: {error}")

    return None

if __name__ == "__main__":
  login()