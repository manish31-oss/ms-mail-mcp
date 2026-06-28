# ms-mail-mcp

A Model Context Protocol (MCP) server that gives Claude Code full access to your
Microsoft Outlook / Microsoft 365 account via the **Microsoft Graph API**.

Works everywhere — no Outlook desktop app required. Compatible with web, mobile,
Mac, Windows, and Linux.

---

## Tools (31 total)

### Mail — Reading
| Tool | Description |
|---|---|
| `list_emails` | List recent inbox emails with read/unread status and attachment flag |
| `list_folders` | List all mail folders with total and unread counts |
| `list_folder_emails` | List emails from any folder by ID or well-known name |
| `read_email` | Read full content of an email (From, To, CC, Subject, Body) |
| `search_emails` | Search emails by keyword across all folders (supports KQL operators) |

### Mail — Composing
| Tool | Description |
|---|---|
| `send_email` | Send an email immediately (CC and BCC optional) |
| `send_email_with_attachment` | Send an email with a file attachment (max 4 MB) |
| `reply_email` | Reply or reply-all to an email |
| `forward_email` | Forward an email to new recipients |
| `create_draft` | Save a draft without sending — returns draft ID |
| `send_draft` | Send a previously saved draft by its message ID |

### Mail — Managing
| Tool | Description |
|---|---|
| `delete_email` | Move to Deleted Items (default) or permanently delete |
| `mark_email` | Mark as read or unread |
| `move_email` | Move email to a folder |
| `copy_email` | Copy email to a folder, leaving original in place |
| `flag_email` | Set follow-up flag: `flagged` · `notFlagged` · `complete` |

### Mail — Attachments
| Tool | Description |
|---|---|
| `get_attachments` | List all file attachments on an email (with IDs, sizes) |
| `save_attachment` | Download and save an attachment to a local directory |

### Folders
| Tool | Description |
|---|---|
| `create_folder` | Create a new mail folder (optionally nested inside another) |
| `delete_folder` | Permanently delete a folder and all its contents |
| `rename_folder` | Rename an existing folder |

### Calendar
| Tool | Description |
|---|---|
| `list_events` | List upcoming events within the next N days (IST timezone) |
| `get_event` | Get full details of an event including attendees and response status |
| `create_event` | Create a meeting or appointment with attendees |
| `update_event` | Update subject, time, location, or body of an existing event |
| `delete_event` | Delete a calendar event |
| `respond_to_event` | Accept, tentatively accept, or decline a meeting invite |

### Contacts
| Tool | Description |
|---|---|
| `list_contacts` | List or search contacts by name, email, or company |
| `get_contact` | Get full details of a contact |
| `create_contact` | Create a new contact in Outlook |

### Profile
| Tool | Description |
|---|---|
| `get_profile` | Get current user's name, email, job title, department, and phone |

**Well-known folder names** (pass these as `folder_id` where accepted):
`inbox` · `sentItems` · `deletedItems` · `drafts` · `junkemail` · `archive`

---

## One-Time Setup

### Step 1 — Register an Azure App (5 minutes, any browser)

1. Go to [portal.azure.com](https://portal.azure.com) → sign in with your work account
2. Search for **"App registrations"** → **New registration**
3. **Name:** `ms-mail-mcp` (any name works)
4. **Supported account types:** *Accounts in this organizational directory only*
5. **Redirect URI:** leave blank → click **Register**
6. Copy the **Application (client) ID** — you will need this as `MS_CLIENT_ID`

**Add API permissions:**

7. Left sidebar → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
8. Search and add all of these:

   | Permission | Used for |
   |---|---|
   | `Mail.Read` | Read emails and folders |
   | `Mail.ReadWrite` | Create, move, delete emails and folders |
   | `Mail.Send` | Send emails |
   | `Calendars.ReadWrite` | Read and manage calendar events |
   | `Contacts.ReadWrite` | Read and manage contacts |
   | `User.Read` | Get your profile |
   | `offline_access` | Keep you signed in (refresh tokens) |

9. Click **Add permissions**

**Enable device code flow (required):**

10. Left sidebar → **Authentication**
11. Scroll to **Advanced settings** → set **Allow public client flows** → **Yes**
12. Click **Save**

> **Corporate IT note:** If your org requires admin approval for apps, you'll see a
> consent error during `auth_setup.py`. Forward the error to your IT/admin and ask
> them to grant consent for the `ms-mail-mcp` app registration. One-time step.

---

### Step 2 — Install dependencies

```bash
pip3 install msal requests mcp
```

---

### Step 3 — Get the code

```bash
git clone https://github.com/manish31-oss/ms-mail-mcp ~/ms-mail-mcp
```

---

### Step 4 — One-time authentication

```bash
MS_CLIENT_ID=your_client_id_here python3 ~/ms-mail-mcp/auth_setup.py
```

The script prints a short URL and a code. Open the URL in any browser, enter the code,
and sign in with your work Microsoft account. The script saves `token_cache.json` next
to `server.py`. You never need to re-authenticate — the server refreshes tokens silently.

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
Call get_profile to confirm the MCP is connected.
```

---

## Notes

- Credentials are passed via environment variables — never hardcoded.
- `token_cache.json` is in `.gitignore` — your tokens never leave your machine.
- Microsoft access tokens expire in 1 hour. The server caches them 50 min in memory
  and calls MSAL for a silent refresh when needed. The refresh token lasts 90 days
  of inactivity (up to 1 year if used regularly).
- `delete_email` default is `permanent=False` (moves to Deleted Items). Pass
  `permanent=True` to skip the bin entirely.
- `delete_folder` is irreversible — it deletes the folder and every email in it.
- `send_email_with_attachment` / `save_attachment` support files up to ~4 MB.
  Files above 4 MB require Microsoft's upload session API (not implemented here).

---

## Uninstall

```bash
claude mcp remove ms-mail -s user
pip3 uninstall msal requests
rm -rf ~/ms-mail-mcp
```

---

## Paste-Ready Setup Prompts

Copy the relevant block below and paste it into Claude Code. Replace `YOUR_CLIENT_ID`
with your Azure App Client ID from Step 1 before pasting.

---

### For macOS (Safari / web-only Outlook users)

```
Set up the Microsoft Outlook MCP server on this Mac.
MY_CLIENT_ID is: YOUR_CLIENT_ID

Step 1 — Check Python 3
Run: python3 --version
If not found, install with: brew install python3

Step 2 — Clone the repo
Run: git clone https://github.com/manish31-oss/ms-mail-mcp ~/ms-mail-mcp
If git is missing, run first: xcode-select --install

Step 3 — Install dependencies
Run: pip3 install msal requests mcp

Step 4 — One-time authentication
Run: MS_CLIENT_ID=YOUR_CLIENT_ID python3 ~/ms-mail-mcp/auth_setup.py
Show me the full output.
Then STOP — tell me to open the URL shown in Safari and sign in with my work account.
Wait for me to confirm I have completed sign-in before moving to Step 5.

Step 5 — Register with Claude Code
Run: claude mcp add ms-mail -s user -e MS_CLIENT_ID=YOUR_CLIENT_ID -- python3 ~/ms-mail-mcp/server.py
Run: claude mcp list
Confirm "ms-mail" appears.

Step 6 — Tell me to restart Claude Code.
After I confirm the restart, call get_profile to verify the MCP is live.
```

---

### For Windows (PowerShell)

```
Set up the Microsoft Outlook MCP server on this Windows machine.
MY_CLIENT_ID is: YOUR_CLIENT_ID

Step 1 — Check Python
Run in PowerShell: python --version
If Python is not installed, download from https://python.org and install it.
During install, check "Add Python to PATH". Then verify: python --version

Step 2 — Clone the repo
Run: git clone https://github.com/manish31-oss/ms-mail-mcp "$env:USERPROFILE\ms-mail-mcp"
If git is not installed, download from https://git-scm.com and install, then retry.

Step 3 — Install dependencies
Run: pip install msal requests mcp

Step 4 — One-time authentication
Run: $env:MS_CLIENT_ID="YOUR_CLIENT_ID"; python "$env:USERPROFILE\ms-mail-mcp\auth_setup.py"
Show me the full output.
Then STOP — tell me to open the URL shown in my browser and sign in with my work Microsoft account.
Wait for me to confirm I have completed sign-in before moving to Step 5.

Step 5 — Register with Claude Code
Run: claude mcp add ms-mail -s user -e MS_CLIENT_ID=YOUR_CLIENT_ID -- python "$env:USERPROFILE\ms-mail-mcp\server.py"
Run: claude mcp list
Confirm "ms-mail" appears.

Step 6 — Tell me to restart Claude Code.
After I confirm the restart, call get_profile to verify the MCP is live.
```
