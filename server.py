import os
import time
import base64
import msal
import requests
from mcp.server.fastmcp import FastMCP

CLIENT_ID  = os.environ["MS_CLIENT_ID"]
AUTHORITY  = "https://login.microsoftonline.com/organizations"
SCOPES     = ["Mail.Read", "Mail.ReadWrite", "Mail.Send", "User.Read", "offline_access"]
CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")
API_BASE   = "https://graph.microsoft.com/v1.0"
TIMEOUT    = 30

mcp = FastMCP("ms-mail")

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

def _req(method, path, payload=None, params=None):
    headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}
    return requests.request(method, f"{API_BASE}{path}", headers=headers, json=payload, params=params, timeout=TIMEOUT)

def _get(path, params=None):
    return _req("GET", path, params=params).json()

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


@mcp.tool()
def list_emails(count: int = 10) -> str:
    """List recent emails from inbox"""
    data = _get("/me/mailFolders/inbox/messages", {
        "$top": count,
        "$select": "id,subject,from,receivedDateTime,isRead",
        "$orderby": "receivedDateTime desc",
    })
    msgs = data.get("value", [])
    if not msgs:
        return "No emails found"
    return "\n".join(
        f"ID:{m['id']} | From:{m['from']['emailAddress']['address']} | "
        f"{'[UNREAD] ' if not m['isRead'] else ''}{m['subject']} | {m['receivedDateTime']}"
        for m in msgs)


@mcp.tool()
def list_folders() -> str:
    """List all mail folders with their IDs"""
    data = _get("/me/mailFolders", {"$top": 50})
    folders = data.get("value", [])
    if not folders:
        return "No folders found"
    return "\n".join(
        f"ID:{f['id']} | {f['displayName']} | Total:{f['totalItemCount']} | Unread:{f['unreadItemCount']}"
        for f in folders)


@mcp.tool()
def list_folder_emails(folder_id: str, count: int = 50) -> str:
    """List emails from a specific folder by folder ID. Use list_folders to get IDs.
    Well-known folder names also work: inbox, sentItems, deletedItems, drafts, junkemail."""
    data = _get(f"/me/mailFolders/{folder_id}/messages", {
        "$top": count,
        "$select": "id,subject,from,toRecipients,receivedDateTime,isRead",
        "$orderby": "receivedDateTime desc",
    })
    msgs = data.get("value", [])
    if not msgs:
        return f"No emails in folder '{folder_id}'"
    return "\n".join(
        f"ID:{m['id']} | From:{m['from']['emailAddress']['address']} | {m['subject']} | {m['receivedDateTime']}"
        for m in msgs)


@mcp.tool()
def read_email(message_id: str) -> str:
    """Read full content of an email by message ID"""
    data = _get(f"/me/messages/{message_id}")
    if "error" in data:
        return f"Error: {data['error'].get('message', str(data))}"
    from_addr = data.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
    subject   = data.get("subject", "")
    received  = data.get("receivedDateTime", "")
    body      = data.get("body", {}).get("content", "")
    return f"From: {from_addr}\nSubject: {subject}\nDate: {received}\n\n{body}"


@mcp.tool()
def search_emails(query: str) -> str:
    """Search emails by keyword — matches subject, sender, body across all folders"""
    data = _get("/me/messages", {
        "$search": f'"{query}"',
        "$top": 50,
        "$select": "id,subject,from,receivedDateTime",
    })
    msgs = data.get("value", [])
    if not msgs:
        return f"No emails found matching '{query}'"
    return "\n".join(
        f"ID:{m['id']} | From:{m['from']['emailAddress']['address']} | {m['subject']} | {m['receivedDateTime']}"
        for m in msgs)


@mcp.tool()
def send_email(to: str, subject: str, body: str, cc: str = "", bcc: str = "") -> str:
    """Send an email. cc and bcc are optional comma-separated addresses."""
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": [{"emailAddress": {"address": a.strip()}} for a in to.split(",")],
        },
        "saveToSentItems": True,
    }
    if cc:  payload["message"]["ccRecipients"]  = [{"emailAddress": {"address": a.strip()}} for a in cc.split(",")]
    if bcc: payload["message"]["bccRecipients"] = [{"emailAddress": {"address": a.strip()}} for a in bcc.split(",")]
    return _ok(_post("/me/sendMail", payload), 202, "Sent successfully")


@mcp.tool()
def reply_email(message_id: str, body: str, reply_all: bool = False) -> str:
    """Reply to an email. Set reply_all=True to reply to all recipients."""
    endpoint = "replyAll" if reply_all else "reply"
    return _ok(_post(f"/me/messages/{message_id}/{endpoint}", {"comment": body}), 202, "Replied successfully")


@mcp.tool()
def forward_email(message_id: str, to: str, body: str = "") -> str:
    """Forward an email to one or more recipients (comma-separated)."""
    to_list = [{"emailAddress": {"address": a.strip()}} for a in to.split(",")]
    return _ok(
        _post(f"/me/messages/{message_id}/forward", {"comment": body, "toRecipients": to_list}),
        202, "Forwarded successfully")


@mcp.tool()
def delete_email(message_id: str, permanent: bool = False) -> str:
    """Delete an email. permanent=False (default) moves to Deleted Items. permanent=True deletes forever."""
    if permanent:
        return _ok(_post(f"/me/messages/{message_id}/permanentDelete", {}), 204, "Permanently deleted")
    return _ok(_delete(f"/me/messages/{message_id}"), 204, "Deleted (moved to Deleted Items)")


@mcp.tool()
def mark_email(message_id: str, as_read: bool = True) -> str:
    """Mark an email as read or unread."""
    r = _patch(f"/me/messages/{message_id}", {"isRead": as_read})
    if r.status_code == 200:
        return f"Marked as {'read' if as_read else 'unread'}"
    try:
        err = r.json()
        return f"Error: {err.get('error', {}).get('message', str(err))}"
    except Exception:
        return f"Unexpected status: {r.status_code}"


@mcp.tool()
def move_email(message_id: str, dest_folder_id: str) -> str:
    """Move an email to a folder. Use list_folders to get folder IDs.
    Well-known folder names also work: inbox, sentItems, deletedItems, drafts, junkemail."""
    return _ok(
        _post(f"/me/messages/{message_id}/move", {"destinationId": dest_folder_id}),
        201, "Moved successfully")


@mcp.tool()
def send_email_with_attachment(to: str, subject: str, body: str, file_path: str, cc: str = "", bcc: str = "") -> str:
    """Send an email with a file attachment (max 4 MB). file_path must be an absolute path."""
    if not os.path.exists(file_path):
        return f"File not found: {file_path}"
    with open(file_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": [{"emailAddress": {"address": a.strip()}} for a in to.split(",")],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": os.path.basename(file_path),
                "contentBytes": content_b64,
            }],
        },
        "saveToSentItems": True,
    }
    if cc:  payload["message"]["ccRecipients"]  = [{"emailAddress": {"address": a.strip()}} for a in cc.split(",")]
    if bcc: payload["message"]["bccRecipients"] = [{"emailAddress": {"address": a.strip()}} for a in bcc.split(",")]
    return _ok(_post("/me/sendMail", payload), 202, "Sent with attachment successfully")


if __name__ == "__main__":
    mcp.run()
