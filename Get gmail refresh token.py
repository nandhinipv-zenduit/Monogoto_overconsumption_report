"""
One-off helper: mint a Gmail refresh token for the reporting job.

Run this ONCE, locally, signed in as the mailbox that will send the reports
(GMAIL_USERNAME). It prints a refresh token; paste that into the GitHub
Actions secret GMAIL_REFRESH_TOKEN. You never need to run it again unless
the token is revoked.

BEFORE RUNNING — create the OAuth client:
  1. console.cloud.google.com -> pick/create a project
  2. APIs & Services -> Library -> enable "Gmail API"
  3. APIs & Services -> OAuth consent screen
       - Internal (if zenduit.com is a Workspace org) — strongly preferred,
         because External + "Testing" expires refresh tokens after 7 DAYS
         and this job would then break every week.
       - Add scope: https://mail.google.com/
  4. Credentials -> Create credentials -> OAuth client ID
       - Application type: Desktop app
       - Copy the client ID and client secret

THEN:
    pip install google-auth-oauthlib
    GMAIL_CLIENT_ID=xxx.apps.googleusercontent.com \
    GMAIL_CLIENT_SECRET=GOCSPX-xxx \
    python get_gmail_refresh_token.py

A browser window opens; consent as the SENDING account (not your personal
one — the token is bound to whoever consents, and it must match
GMAIL_USERNAME or SMTP will still return 535).
"""

import os
import sys

try:
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    sys.exit("Install the dependency first:  pip install google-auth-oauthlib")

# Full-mailbox scope. gmail.send is NOT sufficient for SMTP AUTH XOAUTH2 —
# using it is the most common reason this whole setup still fails with 535.
SCOPES = ["https://mail.google.com/"]

client_id = os.getenv("GMAIL_CLIENT_ID")
client_secret = os.getenv("GMAIL_CLIENT_SECRET")
if not client_id or not client_secret:
    sys.exit("Set GMAIL_CLIENT_ID and GMAIL_CLIENT_SECRET in the environment first.")

flow = InstalledAppFlow.from_client_config(
    {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": ["http://localhost"],
        }
    },
    SCOPES,
)

# access_type=offline + prompt=consent forces Google to hand back a REFRESH
# token. Without prompt=consent, a repeat authorisation returns only an
# access token and you'd be left wondering where the refresh token went.
creds = flow.run_local_server(
    port=0,
    access_type="offline",
    prompt="consent",
)

print("\n" + "=" * 70)
print("Account consented :", getattr(creds, "id_token", None) or "(check the browser prompt)")
print("Granted scopes    :", creds.scopes)
print("\nGitHub Actions secrets to set:")
print(f"  GMAIL_USERNAME      = <the mailbox you just consented as>")
print(f"  GMAIL_CLIENT_ID     = {client_id}")
print(f"  GMAIL_CLIENT_SECRET = {client_secret}")
print(f"  GMAIL_REFRESH_TOKEN = {creds.refresh_token}")
print("\nThen DELETE the old GMAIL_PASS secret — it is no longer read.")
print("=" * 70)
