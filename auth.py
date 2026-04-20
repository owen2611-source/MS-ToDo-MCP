#!/usr/bin/env python3
"""
One-time auth script – run this locally to get your GRAPH_REFRESH_TOKEN.

Usage:
    pip install httpx
    export AZURE_CLIENT_ID=<your-client-id>
    python auth.py

You only need to run this once. Copy the printed env vars to your
hosting platform (Fly.io: fly secrets set ...).

App registration (free, 2 minutes):
  1. Go to https://entra.microsoft.com
  2. Sign in with your Microsoft account (the one you use for To Do)
  3. Identity → Applications → App registrations → New registration
  4. Name: "MS To Do MCP"  |  Accounts: "Personal Microsoft accounts only"
  5. Register → copy the Application (client) ID
  6. Authentication → Add a platform → Mobile and desktop applications
     → check https://login.microsoftonline.com/common/oauth2/nativeclient → Configure
  7. Under "Advanced settings" toggle "Allow public client flows" to Yes → Save
  8. API permissions → Add a permission → Microsoft Graph → Delegated
     → add Tasks.ReadWrite and User.Read → Grant admin consent
"""

import os
import sys
import time

try:
    import httpx
except ImportError:
    sys.exit("Run: pip install httpx")

CLIENT_ID = os.environ.get("AZURE_CLIENT_ID") or input("Enter your Azure Client ID: ").strip()

SCOPE = "Tasks.ReadWrite User.Read offline_access"
DEVICE_CODE_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode"
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"


def main() -> None:
    r = httpx.post(DEVICE_CODE_URL, data={"client_id": CLIENT_ID, "scope": SCOPE})
    if r.status_code != 200:
        sys.exit(f"Device code request failed: {r.text}")
    data = r.json()

    print(f"\n1. Open this URL in your browser:\n   {data['verification_uri']}")
    print(f"\n2. Enter this code: {data['user_code']}\n")
    print("Waiting for you to sign in", end="", flush=True)

    interval = int(data.get("interval", 5))
    deadline = time.time() + int(data.get("expires_in", 900))

    while time.time() < deadline:
        time.sleep(interval)
        resp = httpx.post(TOKEN_URL, data={
            "client_id": CLIENT_ID,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": data["device_code"],
        })
        body = resp.json()

        if "access_token" in body:
            print("\n\nAuthenticated! Set these env vars on your hosting platform:\n")
            print(f"  AZURE_CLIENT_ID={CLIENT_ID}")
            print(f"  GRAPH_REFRESH_TOKEN={body['refresh_token']}")
            print("\nFly.io command:")
            print(f"  fly secrets set AZURE_CLIENT_ID={CLIENT_ID} GRAPH_REFRESH_TOKEN={body['refresh_token']}")
            return

        error = body.get("error", "")
        if error == "authorization_pending":
            print(".", end="", flush=True)
        elif error == "slow_down":
            interval += 5
            print(".", end="", flush=True)
        else:
            sys.exit(f"\nAuth error: {body}")

    sys.exit("\nCode expired. Run again.")


if __name__ == "__main__":
    main()
