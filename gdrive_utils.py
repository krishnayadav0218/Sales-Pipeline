"""
Google Drive API helper functions.
Authenticates using a Service Account (no browser login needed --
essential for a headless server like Render).
"""
import os
import io
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_drive_service():
    """
    Loads the Service Account credentials from the GOOGLE_SERVICE_ACCOUNT_JSON
    environment variable (the full JSON key content, pasted as a single string
    into Render's environment variable settings) and returns an authenticated
    Drive API client.
    """
    raw_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw_json:
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_JSON environment variable not set. "
            "Paste your service account key's full JSON content into it."
        )
    info = json.loads(raw_json)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
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
