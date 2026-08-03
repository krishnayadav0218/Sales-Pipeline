"""
Run this ONCE on your own computer (not on GitHub) to get a new
GOOGLE_REFRESH_TOKEN when the old one expires or gets revoked.

Steps before running:
1. pip install google-auth-oauthlib
2. Make sure you have a `credentials.json` file (downloaded from Google
   Cloud Console -> APIs & Services -> Credentials -> your OAuth Client ID
   -> Download JSON) in the SAME folder as this script.
3. Run:  python get_refresh_token.py
4. A browser window will open -> log in with the Google account that
   owns the Drive folders -> allow access.
5. This script will print your new CLIENT_ID, CLIENT_SECRET, and
   REFRESH_TOKEN. Copy the REFRESH_TOKEN value and update the
   GOOGLE_REFRESH_TOKEN secret in GitHub (Settings -> Secrets and
   variables -> Actions).
"""
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]

def main():
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
    creds = flow.run_local_server(port=0)

    print("\n" + "=" * 60)
    print("SUCCESS! Copy these values into your GitHub repo secrets:")
    print("=" * 60)
    print(f"GOOGLE_CLIENT_ID     = {creds.client_id}")
    print(f"GOOGLE_CLIENT_SECRET = {creds.client_secret}")
    print(f"GOOGLE_REFRESH_TOKEN = {creds.refresh_token}")
    print("=" * 60)
    print("\nMost likely only GOOGLE_REFRESH_TOKEN has changed --")
    print("update just that one secret in GitHub unless the others")
    print("also differ from what you had before.")

if __name__ == "__main__":
    main()
