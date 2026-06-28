"""
One-time authentication script for ms-mail-mcp.
Run this once to sign in with your work Microsoft account.
After success, server.py refreshes tokens automatically.
"""
import os
import sys
import msal

AUTHORITY  = "https://login.microsoftonline.com/organizations"
SCOPES     = ["Mail.Read", "Mail.ReadWrite", "Mail.Send", "User.Read", "offline_access"]
CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")

CLIENT_ID = os.environ.get("MS_CLIENT_ID", "").strip()
if not CLIENT_ID:
    CLIENT_ID = input("Enter your Azure App Client ID: ").strip()
if not CLIENT_ID:
    print("Client ID is required. Exiting.")
    sys.exit(1)

cache = msal.SerializableTokenCache()
if os.path.exists(CACHE_PATH):
    cache.deserialize(open(CACHE_PATH).read())

app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

# Try silent refresh first (re-running after previous auth)
accounts = app.get_accounts()
result = None
if accounts:
    print(f"Found cached account: {accounts[0].get('username', '?')} — refreshing token...")
    result = app.acquire_token_silent(SCOPES, account=accounts[0])

if not result:
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print(f"Failed to start device flow: {flow}")
        sys.exit(1)
    print(f"\n{flow['message']}\n")
    print("Steps:")
    print("  1. Open the URL above in your browser (Safari is fine)")
    print("  2. Enter the code shown above")
    print("  3. Sign in with your work Microsoft account (@infinitelocus.com)")
    print("\nWaiting for you to complete sign-in...\n")
    result = app.acquire_token_by_device_flow(flow)  # blocks until user finishes

if "access_token" in result:
    if cache.has_state_changed:
        with open(CACHE_PATH, "w") as f:
            f.write(cache.serialize())
    username = result.get("id_token_claims", {}).get("preferred_username", "unknown")
    print(f"Authenticated as: {username}")
    print(f"Token cache saved to: {CACHE_PATH}")
    print("\nSetup complete. You can now register and start the MCP server.")
else:
    print(f"\nAuthentication failed: {result.get('error_description', str(result))}")
    sys.exit(1)
