"""
Diagnostic script — run this in GitHub Actions to see exactly what
the Graph API can access for your user. Add it as a workflow step
temporarily to diagnose the OneDrive connection issue.
"""
import os
import requests

CLIENT_ID           = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET       = os.environ["AZURE_CLIENT_SECRET"]
TENANT_ID           = os.environ["AZURE_TENANT_ID"]
ONEDRIVE_USER_EMAIL = os.environ["ONEDRIVE_USER_EMAIL"]

def get_token():
    url = "https://login.microsoftonline.com/{}/oauth2/v2.0/token".format(TENANT_ID)
    payload = {
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default"
    }
    r = requests.post(url, data=payload)
    r.raise_for_status()
    return r.json()["access_token"]

def check(label, url, token):
    headers = {"Authorization": "Bearer {}".format(token)}
    r = requests.get(url, headers=headers)
    print("\n--- {} ---".format(label))
    print("URL:    {}".format(url))
    print("Status: {}".format(r.status_code))
    if r.status_code == 200:
        data = r.json()
        # Print key fields only to avoid massive output
        for key in ["id", "displayName", "userPrincipalName", 
                    "mail", "webUrl", "name", "driveType"]:
            if key in data:
                print("{}: {}".format(key, data[key]))
    else:
        print("Error: {}".format(r.text[:500]))

token = get_token()
print("Token acquired successfully.")
print("Testing user email: {}".format(ONEDRIVE_USER_EMAIL))

# Test 1: Can we find the user at all?
check(
    "User lookup",
    "https://graph.microsoft.com/v1.0/users/{}".format(ONEDRIVE_USER_EMAIL),
    token
)

# Test 2: Can we access their drive?
check(
    "Drive root",
    "https://graph.microsoft.com/v1.0/users/{}/drive".format(ONEDRIVE_USER_EMAIL),
    token
)

# Test 3: Can we list drive root contents?
check(
    "Drive root children",
    "https://graph.microsoft.com/v1.0/users/{}/drive/root/children".format(
        ONEDRIVE_USER_EMAIL),
    token
)

# Test 4: Try with the user's ID instead of email (in case UPN lookup is broken)
print("\n--- Fetching user ID for alternate lookup ---")
headers = {"Authorization": "Bearer {}".format(token)}
r = requests.get(
    "https://graph.microsoft.com/v1.0/users/{}".format(ONEDRIVE_USER_EMAIL),
    headers=headers
)
if r.status_code == 200:
    user_id = r.json().get("id")
    print("User ID: {}".format(user_id))
    check(
        "Drive root by user ID",
        "https://graph.microsoft.com/v1.0/users/{}/drive".format(user_id),
        token
    )
else:
    print("Could not retrieve user ID — user lookup failed.")
