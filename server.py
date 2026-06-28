import os
import time
import base64
import msal
import requests
from datetime import datetime, timedelta, timezone
from mcp.server.fastmcp import FastMCP

CLIENT_ID  = os.environ["MS_CLIENT_ID"]
AUTHORITY  = "https://login.microsoftonline.com/organizations"
SCOPES     = [
    "Mail.Read", "Mail.ReadWrite", "Mail.Send",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "User.Read", "offline_access",
]
CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")
API_BASE   = "https://graph.microsoft.com/v1.0"
TIMEOUT    = 30

mcp = FastMCP("ms-mail")

# ── Token management ──────────────────────────────────────────────────────────

_token_mem = {"token": None, "expires_at": 0}

def _load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        cache.deserialize(open(CACHE_PATH).read())
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        with open(CACHE_PATH, "w") as f:
            f.write(cache.serialize())

def get_access_token():
    if _token_mem["token"] and time.time() < _token_mem["expires_at"]:
        return _token_mem["token"]
    cache = _load_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if not accounts:
        raise RuntimeError("No cached credentials. Run: python3 auth_setup.py")
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save_cache(cache)
    if not result or "access_token" not in result:
        raise RuntimeError(f"Token refresh failed: {result}")
    _token_mem["token"] = result["access_token"]
    _token_mem["expires_at"] = time.time() + 3000  # cache 50 min
    return _token_mem["token"]

# ── HTTP helpers ──────────────────────────────────────────────────────────────

def _req(method, path, payload=None, params=None, extra_headers=None):
    headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}
    if extra_headers:
        headers.update(extra_headers)
    return requests.request(method, f"{API_BASE}{path}", headers=headers, json=payload, params=params, timeout=TIMEOUT)

def _get(path, params=None, extra_headers=None):
    return _req("GET", path, params=params, extra_headers=extra_headers).json()

def _post(path, payload=None):
    return _req("POST", path, payload=payload)

def _patch(path, payload):
    return _req("PATCH", path, payload=payload)

def _delete(path):
    return _req("DELETE", path)

def _ok(r, expected_status, msg):
    if r.status_code == expected_status:
        return msg
    try:
        err = r.json()
        return f"Error {r.status_code}: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"

def _recipients(csv: str) -> list:
    return [{"emailAddress": {"address": a.strip()}} for a in csv.split(",") if a.strip()]


# ════════════════════════════════════════════════════════════════════════════════
# MAIL — READING
# ════════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def list_emails(count: int = 10) -> str:
    """List recent emails from your inbox."""
    data = _get("/me/mailFolders/inbox/messages", {
        "$top": count,
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments",
        "$orderby": "receivedDateTime desc",
    })
    msgs = data.get("value", [])
    if not msgs:
        return "Inbox is empty."
    return "\n".join(
        f"ID:{m['id']}\n  From: {m['from']['emailAddress']['address']}\n"
        f"  Subject: {'[UNREAD] ' if not m['isRead'] else ''}{m['subject']}\n"
        f"  Date: {m['receivedDateTime']}  Attachments: {m['hasAttachments']}"
        for m in msgs)


@mcp.tool()
def list_folders() -> str:
    """List all top-level mail folders with their IDs, total count, and unread count."""
    data = _get("/me/mailFolders", {"$top": 100})
    folders = data.get("value", [])
    if not folders:
        return "No folders found."
    return "\n".join(
        f"ID:{f['id']} | {f['displayName']} | Total:{f['totalItemCount']} | Unread:{f['unreadItemCount']}"
        for f in folders)


@mcp.tool()
def list_folder_emails(folder_id: str, count: int = 50) -> str:
    """List emails from a specific folder.
    Pass a folder ID (from list_folders) or a well-known name:
    inbox · sentItems · deletedItems · drafts · junkemail · archive"""
    data = _get(f"/me/mailFolders/{folder_id}/messages", {
        "$top": count,
        "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments",
        "$orderby": "receivedDateTime desc",
    })
    msgs = data.get("value", [])
    if not msgs:
        return f"No emails in '{folder_id}'."
    return "\n".join(
        f"ID:{m['id']} | From:{m['from']['emailAddress']['address']} | "
        f"{'[UNREAD] ' if not m['isRead'] else ''}{m['subject']} | {m['receivedDateTime']}"
        for m in msgs)


@mcp.tool()
def read_email(message_id: str) -> str:
    """Read the full content of an email by its message ID."""
    data = _get(f"/me/messages/{message_id}")
    if "error" in data:
        return f"Error: {data['error'].get('message', str(data))}"
    from_addr = data.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
    to_list   = ", ".join(r["emailAddress"]["address"] for r in data.get("toRecipients", []))
    cc_list   = ", ".join(r["emailAddress"]["address"] for r in data.get("ccRecipients", []))
    subject   = data.get("subject", "")
    received  = data.get("receivedDateTime", "")
    body      = data.get("body", {}).get("content", "")
    header    = f"From: {from_addr}\nTo: {to_list}\n"
    if cc_list:
        header += f"CC: {cc_list}\n"
    header += f"Subject: {subject}\nDate: {received}\n\n"
    return header + body


@mcp.tool()
def search_emails(query: str, count: int = 50) -> str:
    """Search emails by keyword — matches subject, sender, and body across all folders.
    Supports KQL operators: from:name@domain.com, subject:keyword, hasAttachments:true"""
    data = _get("/me/messages", {
        "$search": f'"{query}"',
        "$top": count,
        "$select": "id,subject,from,receivedDateTime,isRead",
    })
    msgs = data.get("value", [])
    if not msgs:
        return f"No emails found matching '{query}'."
    return "\n".join(
        f"ID:{m['id']} | From:{m['from']['emailAddress']['address']} | "
        f"{'[UNREAD] ' if not m['isRead'] else ''}{m['subject']} | {m['receivedDateTime']}"
        for m in msgs)


# ════════════════════════════════════════════════════════════════════════════════
# MAIL — COMPOSING
# ════════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def send_email(to: str, subject: str, body: str, cc: str = "", bcc: str = "") -> str:
    """Send an email immediately. to/cc/bcc accept comma-separated addresses."""
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": _recipients(to),
        },
        "saveToSentItems": True,
    }
    if cc:  payload["message"]["ccRecipients"]  = _recipients(cc)
    if bcc: payload["message"]["bccRecipients"] = _recipients(bcc)
    return _ok(_post("/me/sendMail", payload), 202, "Sent successfully.")


@mcp.tool()
def send_email_with_attachment(to: str, subject: str, body: str, file_path: str,
                               cc: str = "", bcc: str = "") -> str:
    """Send an email with a file attachment (max 4 MB). file_path must be an absolute path."""
    if not os.path.exists(file_path):
        return f"File not found: {file_path}"
    with open(file_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": _recipients(to),
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": os.path.basename(file_path),
                "contentBytes": content_b64,
            }],
        },
        "saveToSentItems": True,
    }
    if cc:  payload["message"]["ccRecipients"]  = _recipients(cc)
    if bcc: payload["message"]["bccRecipients"] = _recipients(bcc)
    return _ok(_post("/me/sendMail", payload), 202, "Sent with attachment successfully.")


@mcp.tool()
def reply_email(message_id: str, body: str, reply_all: bool = False) -> str:
    """Reply to an email. Set reply_all=True to reply to all recipients."""
    endpoint = "replyAll" if reply_all else "reply"
    return _ok(
        _post(f"/me/messages/{message_id}/{endpoint}", {"comment": body}),
        202, "Replied successfully.")


@mcp.tool()
def forward_email(message_id: str, to: str, body: str = "") -> str:
    """Forward an email to one or more recipients (comma-separated)."""
    return _ok(
        _post(f"/me/messages/{message_id}/forward",
              {"comment": body, "toRecipients": _recipients(to)}),
        202, "Forwarded successfully.")


@mcp.tool()
def create_draft(to: str, subject: str, body: str, cc: str = "", bcc: str = "") -> str:
    """Save a draft email without sending it. Returns the draft message ID for later use."""
    payload = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": body},
        "toRecipients": _recipients(to),
    }
    if cc:  payload["ccRecipients"]  = _recipients(cc)
    if bcc: payload["bccRecipients"] = _recipients(bcc)
    r = _post("/me/messages", payload)
    if r.status_code == 201:
        draft = r.json()
        return f"Draft saved. ID: {draft['id']}"
    try:
        err = r.json()
        return f"Error {r.status_code}: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


@mcp.tool()
def send_draft(message_id: str) -> str:
    """Send a previously saved draft by its message ID."""
    return _ok(_post(f"/me/messages/{message_id}/send", {}), 202, "Draft sent successfully.")


# ════════════════════════════════════════════════════════════════════════════════
# MAIL — MANAGING
# ════════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def delete_email(message_id: str, permanent: bool = False) -> str:
    """Delete an email. permanent=False (default) moves to Deleted Items. permanent=True deletes forever."""
    if permanent:
        return _ok(_post(f"/me/messages/{message_id}/permanentDelete", {}), 204, "Permanently deleted.")
    return _ok(_delete(f"/me/messages/{message_id}"), 204, "Moved to Deleted Items.")


@mcp.tool()
def mark_email(message_id: str, as_read: bool = True) -> str:
    """Mark an email as read (as_read=True) or unread (as_read=False)."""
    r = _patch(f"/me/messages/{message_id}", {"isRead": as_read})
    if r.status_code == 200:
        return f"Marked as {'read' if as_read else 'unread'}."
    try:
        err = r.json()
        return f"Error: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


@mcp.tool()
def move_email(message_id: str, dest_folder_id: str) -> str:
    """Move an email to a folder. Accepts a folder ID or well-known name (e.g. inbox, drafts)."""
    return _ok(
        _post(f"/me/messages/{message_id}/move", {"destinationId": dest_folder_id}),
        201, "Moved successfully.")


@mcp.tool()
def copy_email(message_id: str, dest_folder_id: str) -> str:
    """Copy an email to a folder, leaving the original in place."""
    r = _post(f"/me/messages/{message_id}/copy", {"destinationId": dest_folder_id})
    if r.status_code == 201:
        return f"Copied successfully. New ID: {r.json().get('id', '?')}"
    try:
        err = r.json()
        return f"Error {r.status_code}: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


@mcp.tool()
def flag_email(message_id: str, flag_status: str = "flagged") -> str:
    """Set the follow-up flag on an email.
    flag_status options: 'flagged' | 'notFlagged' | 'complete'"""
    r = _patch(f"/me/messages/{message_id}", {"flag": {"flagStatus": flag_status}})
    if r.status_code == 200:
        return f"Flag set to '{flag_status}'."
    try:
        err = r.json()
        return f"Error: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


# ════════════════════════════════════════════════════════════════════════════════
# MAIL — ATTACHMENTS
# ════════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def get_attachments(message_id: str) -> str:
    """List all file attachments for an email (excludes inline images)."""
    data = _get(f"/me/messages/{message_id}/attachments",
                {"$select": "id,name,size,contentType,isInline"})
    attachments = [a for a in data.get("value", []) if not a.get("isInline")]
    if not attachments:
        return "No file attachments found."
    return "\n".join(
        f"ID:{a['id']} | {a['name']} | {a['contentType']} | {a.get('size', '?')} bytes"
        for a in attachments)


@mcp.tool()
def save_attachment(message_id: str, attachment_id: str, save_dir: str) -> str:
    """Download and save an email attachment to disk. save_dir must be an absolute path."""
    if not os.path.isdir(save_dir):
        return f"Directory not found: {save_dir}"
    data = _get(f"/me/messages/{message_id}/attachments/{attachment_id}")
    if "error" in data:
        return f"Error: {data['error'].get('message', str(data))}"
    filename = data.get("name", "attachment")
    content_b64 = data.get("contentBytes", "")
    if not content_b64:
        return "Attachment has no content."
    content = base64.b64decode(content_b64)
    out_path = os.path.join(save_dir, filename)
    with open(out_path, "wb") as f:
        f.write(content)
    return f"Saved to: {out_path} ({len(content)} bytes)"


# ════════════════════════════════════════════════════════════════════════════════
# FOLDERS
# ════════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def create_folder(folder_name: str, parent_folder_id: str = "") -> str:
    """Create a new mail folder. Optionally pass parent_folder_id to nest it inside another folder."""
    path = (f"/me/mailFolders/{parent_folder_id}/childFolders"
            if parent_folder_id else "/me/mailFolders")
    r = _post(path, {"displayName": folder_name})
    if r.status_code == 201:
        return f"Folder created. ID: {r.json().get('id', '?')}"
    try:
        err = r.json()
        return f"Error {r.status_code}: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


@mcp.tool()
def delete_folder(folder_id: str) -> str:
    """Permanently delete a mail folder and ALL emails inside it. This cannot be undone."""
    return _ok(_delete(f"/me/mailFolders/{folder_id}"), 204, "Folder deleted permanently.")


@mcp.tool()
def rename_folder(folder_id: str, new_name: str) -> str:
    """Rename a mail folder."""
    r = _patch(f"/me/mailFolders/{folder_id}", {"displayName": new_name})
    if r.status_code == 200:
        return f"Folder renamed to '{new_name}'."
    try:
        err = r.json()
        return f"Error: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


# ════════════════════════════════════════════════════════════════════════════════
# CALENDAR
# ════════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def list_events(days_ahead: int = 7, count: int = 20) -> str:
    """List upcoming calendar events within the next N days. Times shown in Asia/Kolkata (IST)."""
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days_ahead)
    data = _get("/me/calendarView", {
        "startDateTime": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "endDateTime":   end.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "$top": count,
        "$select": "id,subject,start,end,organizer,location,isOnlineMeeting,responseStatus",
        "$orderby": "start/dateTime",
    }, extra_headers={"Prefer": 'outlook.timezone="Asia/Kolkata"'})
    events = data.get("value", [])
    if not events:
        return f"No events in the next {days_ahead} days."
    return "\n".join(
        f"ID:{e['id']}\n"
        f"  Subject: {e.get('subject', '')}\n"
        f"  Start: {e['start']['dateTime']}  End: {e['end']['dateTime']}\n"
        f"  Organizer: {e.get('organizer', {}).get('emailAddress', {}).get('address', '?')}\n"
        f"  Location: {e.get('location', {}).get('displayName', 'N/A')}\n"
        f"  Online: {e.get('isOnlineMeeting', False)}  "
        f"  Response: {e.get('responseStatus', {}).get('response', '?')}"
        for e in events)


@mcp.tool()
def get_event(event_id: str) -> str:
    """Get full details of a calendar event by its ID."""
    data = _get(f"/me/events/{event_id}")
    if "error" in data:
        return f"Error: {data['error'].get('message', str(data))}"
    attendees = ", ".join(
        f"{a['emailAddress']['address']} ({a.get('status', {}).get('response', '?')})"
        for a in data.get("attendees", []))
    body = data.get("body", {}).get("content", "")
    return (
        f"Subject: {data.get('subject', '')}\n"
        f"Start: {data['start']['dateTime']} ({data['start'].get('timeZone', '')})\n"
        f"End:   {data['end']['dateTime']} ({data['end'].get('timeZone', '')})\n"
        f"Location: {data.get('location', {}).get('displayName', 'N/A')}\n"
        f"Organizer: {data.get('organizer', {}).get('emailAddress', {}).get('address', '?')}\n"
        f"Attendees: {attendees or 'None'}\n"
        f"Online meeting: {data.get('isOnlineMeeting', False)}\n\n"
        f"{body}"
    )


@mcp.tool()
def create_event(subject: str, start: str, end: str,
                 attendees: str = "", body: str = "",
                 location: str = "", tz: str = "Asia/Kolkata",
                 is_online_meeting: bool = False) -> str:
    """Create a calendar event / meeting.
    start / end format: 'YYYY-MM-DDTHH:MM:SS'  (e.g. '2026-07-01T10:00:00')
    attendees: comma-separated email addresses
    tz: IANA timezone string (default 'Asia/Kolkata')"""
    payload = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": body},
        "start": {"dateTime": start, "timeZone": tz},
        "end":   {"dateTime": end,   "timeZone": tz},
        "isOnlineMeeting": is_online_meeting,
    }
    if location:
        payload["location"] = {"displayName": location}
    if attendees:
        payload["attendees"] = [
            {"emailAddress": {"address": a.strip()}, "type": "required"}
            for a in attendees.split(",") if a.strip()
        ]
    r = _post("/me/calendar/events", payload)
    if r.status_code == 201:
        return f"Event created. ID: {r.json().get('id', '?')}"
    try:
        err = r.json()
        return f"Error {r.status_code}: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


@mcp.tool()
def update_event(event_id: str, subject: str = "", start: str = "", end: str = "",
                 body: str = "", location: str = "", tz: str = "Asia/Kolkata") -> str:
    """Update fields of an existing calendar event. Only pass the fields you want to change."""
    payload = {}
    if subject:  payload["subject"] = subject
    if body:     payload["body"] = {"contentType": "HTML", "content": body}
    if start:    payload["start"] = {"dateTime": start, "timeZone": tz}
    if end:      payload["end"]   = {"dateTime": end,   "timeZone": tz}
    if location: payload["location"] = {"displayName": location}
    if not payload:
        return "Nothing to update — provide at least one field."
    r = _patch(f"/me/events/{event_id}", payload)
    if r.status_code == 200:
        return "Event updated successfully."
    try:
        err = r.json()
        return f"Error: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


@mcp.tool()
def delete_event(event_id: str) -> str:
    """Delete a calendar event permanently."""
    return _ok(_delete(f"/me/events/{event_id}"), 204, "Event deleted.")


@mcp.tool()
def respond_to_event(event_id: str, response: str, comment: str = "") -> str:
    """Respond to a meeting invitation.
    response options: 'accept' | 'tentativelyAccept' | 'decline'"""
    valid = {"accept", "tentativelyAccept", "decline"}
    if response not in valid:
        return f"Invalid response '{response}'. Use: accept | tentativelyAccept | decline"
    return _ok(
        _post(f"/me/events/{event_id}/{response}",
              {"comment": comment, "sendResponse": True}),
        202, f"Response '{response}' sent.")


# ════════════════════════════════════════════════════════════════════════════════
# CONTACTS
# ════════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def list_contacts(query: str = "", count: int = 25) -> str:
    """List or search contacts. Pass a query string to filter by name, email, or company."""
    params = {
        "$top": count,
        "$select": "id,displayName,emailAddresses,jobTitle,companyName,mobilePhone",
    }
    if query:
        params["$search"] = f'"{query}"'
    data = _get("/me/contacts", params)
    contacts = data.get("value", [])
    if not contacts:
        return "No contacts found."
    lines = []
    for c in contacts:
        emails = ", ".join(e["address"] for e in c.get("emailAddresses", []))
        lines.append(
            f"ID:{c['id']} | {c.get('displayName', '?')} | {emails} | "
            f"{c.get('jobTitle', '')} @ {c.get('companyName', '')}"
        )
    return "\n".join(lines)


@mcp.tool()
def get_contact(contact_id: str) -> str:
    """Get full details of a contact by their ID."""
    data = _get(f"/me/contacts/{contact_id}")
    if "error" in data:
        return f"Error: {data['error'].get('message', str(data))}"
    emails  = ", ".join(e["address"] for e in data.get("emailAddresses", []))
    phones  = data.get("mobilePhone", "") or data.get("businessPhones", [""])[0]
    return (
        f"Name: {data.get('displayName', '?')}\n"
        f"Email: {emails}\n"
        f"Phone: {phones}\n"
        f"Title: {data.get('jobTitle', 'N/A')}\n"
        f"Company: {data.get('companyName', 'N/A')}\n"
        f"Department: {data.get('department', 'N/A')}"
    )


@mcp.tool()
def create_contact(first_name: str, last_name: str = "", email: str = "",
                   phone: str = "", job_title: str = "", company: str = "") -> str:
    """Create a new contact in your Outlook contacts."""
    payload = {"givenName": first_name}
    if last_name:  payload["surname"]      = last_name
    if job_title:  payload["jobTitle"]     = job_title
    if company:    payload["companyName"]  = company
    if phone:      payload["mobilePhone"]  = phone
    if email:
        payload["emailAddresses"] = [{"address": email, "name": f"{first_name} {last_name}".strip()}]
    r = _post("/me/contacts", payload)
    if r.status_code == 201:
        return f"Contact created. ID: {r.json().get('id', '?')}"
    try:
        err = r.json()
        return f"Error {r.status_code}: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


# ════════════════════════════════════════════════════════════════════════════════
# PROFILE
# ════════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def get_profile() -> str:
    """Get the current signed-in user's profile information."""
    data = _get("/me", {
        "$select": "displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone"
    })
    if "error" in data:
        return f"Error: {data['error'].get('message', str(data))}"
    return (
        f"Name:       {data.get('displayName', '?')}\n"
        f"Email:      {data.get('mail', data.get('userPrincipalName', '?'))}\n"
        f"Job Title:  {data.get('jobTitle', 'N/A')}\n"
        f"Department: {data.get('department', 'N/A')}\n"
        f"Office:     {data.get('officeLocation', 'N/A')}\n"
        f"Phone:      {data.get('mobilePhone', 'N/A')}"
    )


if __name__ == "__main__":
    mcp.run()
