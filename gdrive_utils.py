"""
Google Drive API helper functions.

Authenticates as YOUR OWN Google account (via a one-time OAuth refresh
token) rather than a Service Account. This matters because Service
Accounts have zero storage quota on personal Gmail accounts and cannot
create/upload files -- only a real user identity can. Reading/downloading
doesn't need quota, but writing the report back does, so we authenticate
as you for everything.
"""
import os
import io
import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_drive_service():
    """
    Builds an authenticated Drive client using OAuth credentials for your
    own Google account: a client ID, client secret, and a long-lived
    refresh token (generated once via get_refresh_token.py).
    """
    client_id = os.environ.get("GOOGLE_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")
    refresh_token = os.environ.get("GOOGLE_REFRESH_TOKEN")
    if not all([client_id, client_secret, refresh_token]):
        raise RuntimeError(
            "GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET and GOOGLE_REFRESH_TOKEN "
            "environment variables must all be set. Run get_refresh_token.py "
            "once locally to obtain these."
        )
    creds = Credentials(
        token=None,
        refresh_token=refresh_token,
        client_id=client_id,
        client_secret=client_secret,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=SCOPES,
    )
    creds.refresh(Request())
    return build("drive", "v3", credentials=creds)


SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")


def list_data_files_in_folder(service, folder_id):
    """
    Returns a list of {id, name} dicts for every supported data file
    (.xlsx, .xls, .csv) in the given folder. Filters by filename extension
    rather than mimeType, because Drive can store the same file under a
    few different mimeTypes depending on upload settings.
    """
    query = f"'{folder_id}' in parents and trashed=false"
    results = []
    page_token = None
    while True:
        response = service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
        ).execute()
        results.extend(response.get("files", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            break

    matched = [f for f in results if f["name"].lower().endswith(SUPPORTED_EXTENSIONS)]
    skipped = [f for f in results if f not in matched]
    if skipped:
        names = ", ".join(f["name"] for f in skipped)
        print(f"[INFO] Ignoring non-data files in folder: {names}")
    return matched


def download_file(service, file_id, dest_path):
    """Downloads a Drive file's raw bytes to a local path."""
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(dest_path, "wb")
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()


def upload_or_replace_file(service, local_path, folder_id, filename):
    """
    Uploads local_path to the given Drive folder as `filename`.
    If a file with that name already exists there, it's updated in place
    (so the same shareable link always shows the latest report) instead of
    creating duplicate copies on every run.
    """
    existing = service.files().list(
        q=f"'{folder_id}' in parents and name='{filename}' and trashed=false",
        fields="files(id)",
    ).execute().get("files", [])

    media = MediaFileUpload(local_path, mimetype=EXCEL_MIME, resumable=True)

    if existing:
        file_id = existing[0]["id"]
        service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    else:
        metadata = {"name": filename, "parents": [folder_id]}
        created = service.files().create(body=metadata, media_body=media, fields="id").execute()
        return created["id"]
