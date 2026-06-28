# ms-mail-mcp

A Model Context Protocol (MCP) server that connects Microsoft Outlook / Microsoft 365
to Claude Code via the Microsoft Graph API. Works with any account — web, mobile, Mac,
Windows. No Outlook desktop app required.

## Tools

| Tool | Description |
|---|---|
| `list_emails` | List recent inbox emails |
| `list_folders` | List all mail folders with their IDs |
| `list_folder_emails` | List emails from any folder by ID or well-known name |
| `read_email` | Read full content of an email by message ID |
| `search_emails` | Search emails by keyword across all folders |
| `send_email` | Send an email (with optional CC/BCC) |
| `send_email_with_attachment` | Send an email with a file attachment (max 4 MB) |
| `reply_email` | Reply or reply-all to an email |
| `forward_email` | Forward an email to new recipients |
| `delete_email` | Move to Deleted Items or permanently delete |
| `mark_email` | Mark as read or unread |
| `move_email` | Move email to a folder |

Well-known folder names for `list_folder_emails` / `move_email`:
`inbox` · `sentItems` · `deletedItems` · `drafts` · `junkemail`

---

## One-Time Setup (do this once, on the machine where Claude Code runs)

### Step 1 — Register an Azure App

Go to [portal.azure.com](https://portal.azure.com) → sign in with your work account → search **"App registrations"**.

1. Click **New registration**
2. **Name:** `ms-mail-mcp` (anything works)
3. **Supported account types:** *Accounts in this organizational directory only*
4. **Redirect URI:** leave blank → click **Register**
5. Copy the **Application (client) ID** — you'll need it as `MS_CLIENT_ID`

**Add API permissions:**
6. Left sidebar → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
7. Search and add each of these:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `User.Read`
   - `offline_access`
8. Click **Add permissions**. You do NOT need to click "Grant admin consent" for delegated permissions.

**Enable device code flow (required):**
9. Left sidebar → **Authentication**
10. Scroll to **Advanced settings** → set **Allow public client flows** → **Yes**
11. Click **Save**

> **Corporate IT note:** If your org has disabled user consent for apps, you'll see an error
> during `auth_setup.py` asking for admin approval. Forward the error to your IT/admin team
> and ask them to grant consent for the `ms-mail-mcp` app. This is a one-time step.

---

### Step 2 — Install dependencies

```bash
pip3 install msal requests mcp
```

Or with uv:
```bash
uv pip install msal requests mcp
```

---

### Step 3 — Clone this repo

```bash
git clone https://github.com/manish31-oss/ms-mail-mcp ~/ms-mail-mcp
```

---

### Step 4 — Authenticate (one-time, ~2 minutes)

```bash
MS_CLIENT_ID=your_client_id_here python3 ~/ms-mail-mcp/auth_setup.py
```

The script will print a URL and a short code. Open the URL in Safari, enter the code,
sign in with your `@infinitelocus.com` account. The script saves `token_cache.json`
next to `server.py`. You never need to re-authenticate — the server refreshes tokens
automatically using the saved refresh token.

---

### Step 5 — Register with Claude Code

```bash
claude mcp add ms-mail -s user \
  -e MS_CLIENT_ID=your_client_id_here \
  -- python3 ~/ms-mail-mcp/server.py
```

Verify:
```bash
claude mcp list
```

---

### Step 6 — Restart Claude Code

After restart, test with:
```
List my last 5 inbox emails.
```

---

## Notes

- Credentials are passed via environment variables — never hardcoded.
- Token cache (`token_cache.json`) is excluded from git via `.gitignore`. Never commit it.
- Microsoft access tokens expire in 1 hour; the server caches them in memory for 50 minutes
  and auto-refreshes via MSAL when needed. The refresh token lasts up to 90 days of inactivity.
- `delete_email` default is `permanent=False` (moves to Deleted Items). Pass `permanent=True`
  to skip the bin entirely.
- Attachment uploads >4 MB require Microsoft's upload session API — not implemented here.
  Use `send_email_with_attachment` for files under 4 MB.

---

## Uninstall

```bash
claude mcp remove ms-mail -s user
pip3 uninstall msal requests
rm -rf ~/ms-mail-mcp
```

---

## Paste-Ready Setup Prompt (for MacBook — paste into Claude Code at office)

Copy everything in the block below and paste it into Claude Code on your MacBook.
Before pasting, replace `YOUR_CLIENT_ID` with your actual Azure App Client ID from Step 1.

---

```
Set up the Microsoft Outlook MCP server on this Mac. MY_CLIENT_ID is: YOUR_CLIENT_ID

Step 1 — Check Python 3
Run: python3 --version
If Python 3 is not installed, run: brew install python3

Step 2 — Clone the repo
Run: git clone https://github.com/manish31-oss/ms-mail-mcp ~/ms-mail-mcp
If git is not installed, run: xcode-select --install

Step 3 — Install dependencies
Run: pip3 install msal requests mcp

Step 4 — Run one-time authentication
Run: MS_CLIENT_ID=YOUR_CLIENT_ID python3 ~/ms-mail-mcp/auth_setup.py
Show me the full output including the URL and code.
Then STOP and tell me to open the URL in Safari and sign in.
Wait for me to confirm I have completed sign-in before continuing.

Step 5 — Register with Claude Code
Run: claude mcp add ms-mail -s user -e MS_CLIENT_ID=YOUR_CLIENT_ID -- python3 ~/ms-mail-mcp/server.py
Then run: claude mcp list
Confirm "ms-mail" appears.

Step 6 — Tell me to restart Claude Code.
After I restart, call list_folders to confirm the MCP is live.
```
