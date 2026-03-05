"""
Google Integration Module for Email Intelligence Dashboard
============================================================
Provides GoogleAuth, GmailClient, and GoogleCalendarClient classes
that normalize Gmail/Calendar API responses to match the Microsoft Graph
format used by the dashboard. This allows the scoring engine, card
rendering, and all UI components to work identically with both providers.

Required packages:
    pip install google-auth google-auth-oauthlib google-api-python-client

Usage:
    from google_client import GoogleAuth, GmailClient, GoogleCalendarClient
"""

import base64
import email as email_lib
import json
import os
import re
import threading
from datetime import datetime, timedelta, timezone
from html import unescape
from urllib.parse import quote as urlquote

# Lazy imports — these are only needed if Google account is configured
_google_imported = False
_import_error = None

def _ensure_google_imports():
    """Lazy-import Google libraries. Returns True if available."""
    global _google_imported, _import_error
    if _google_imported:
        return True
    if _import_error:
        return False
    try:
        global Credentials, InstalledAppFlow, Request, build
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
        _google_imported = True
        return True
    except ImportError as e:
        _import_error = str(e)
        return False


def is_google_available():
    """Check if Google API libraries are installed."""
    return _ensure_google_imports()


def get_google_import_error():
    """Return the import error message if Google libs aren't available."""
    _ensure_google_imports()
    return _import_error


# ═══════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/userinfo.email",
]

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".outlook_dashboard")
GOOGLE_TOKEN_FILE = os.path.join(CONFIG_DIR, "google_token.json")
# Look for credentials in user config dir first, then bundled with the app
_BUNDLED_CREDS = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                              "assets", "google_credentials.json")
GOOGLE_CREDS_FILE = os.path.join(CONFIG_DIR, "google_credentials.json")


def _ensure_config_dir():
    os.makedirs(CONFIG_DIR, exist_ok=True)


# Network error patterns (connection drops, firewall kills, DNS failures)
_NETWORK_ERROR_PATTERNS = (
    "ConnectionError", "ConnectionReset", "ConnectionAborted",
    "10053", "10054", "forcibly closed", "aborted by the software",
    "NameResolutionError", "getaddrinfo failed", "NewConnectionError",
    "ConnectTimeoutError", "ReadTimeout", "SSLError", "ProxyError",
    "RemoteDisconnected", "Max retries", "timed out",
    "TransportError", "ServerNotFoundError",
)


def _is_network_error(err):
    """Check if an exception is a network/connectivity issue vs auth revocation."""
    err_str = str(err)
    return any(p in err_str for p in _NETWORK_ERROR_PATTERNS)


# ═══════════════════════════════════════════════════════════════
# GOOGLE OAUTH AUTHENTICATION
# ═══════════════════════════════════════════════════════════════

class GoogleAuth:
    """OAuth 2.0 authentication for Google APIs (Gmail + Calendar).

    Requires a google_credentials.json (OAuth client ID) file in the
    config directory. On first auth, opens a browser for consent. Tokens
    are cached to google_token.json for silent refresh on subsequent runs.
    """

    def __init__(self, credentials_file=None):
        if not _ensure_google_imports():
            raise ImportError(
                f"Google API libraries not installed: {_import_error}\n"
                "Run: pip install google-auth google-auth-oauthlib google-api-python-client"
            )
        self._creds_file = credentials_file or GOOGLE_CREDS_FILE
        self._token_file = GOOGLE_TOKEN_FILE
        self._creds = None
        self._lock = threading.RLock()
        self._load_token()

    def _load_token(self):
        """Load cached credentials from disk."""
        if os.path.exists(self._token_file):
            try:
                self._creds = Credentials.from_authorized_user_file(
                    self._token_file, GOOGLE_SCOPES)
            except Exception:
                self._creds = None

    def _save_token(self):
        """Persist credentials to disk."""
        if self._creds:
            _ensure_config_dir()
            with open(self._token_file, "w") as f:
                f.write(self._creds.to_json())

    def get_credentials(self):
        """Return valid credentials, refreshing or re-authenticating as needed.
        If refresh fails due to network error, raises instead of opening browser."""
        with self._lock:
            if self._creds and self._creds.valid:
                return self._creds
            if self._creds and self._creds.expired and self._creds.refresh_token:
                try:
                    self._creds.refresh(Request())
                    self._save_token()
                    return self._creds
                except Exception as e:
                    if _is_network_error(e):
                        # Network is down — don't try interactive auth, just raise
                        raise ConnectionError(
                            f"Google token refresh failed (offline): {e}") from e
                    pass  # Auth revocation — fall through to interactive
            return self._authenticate_interactive()

    def _authenticate_interactive(self):
        """Run browser-based OAuth consent flow."""
        # Check user-specified path, then bundled path
        creds_path = self._creds_file
        if not os.path.exists(creds_path) and os.path.exists(_BUNDLED_CREDS):
            creds_path = _BUNDLED_CREDS
        if not os.path.exists(creds_path):
            raise FileNotFoundError(
                f"Google credentials file not found: {self._creds_file}\n\n"
                "To set up Google integration:\n"
                "1. Place google_credentials.json in the assets/ directory, OR\n"
                "2. Place it in your config directory:\n"
                f"   {self._creds_file}"
            )
        flow = InstalledAppFlow.from_client_secrets_file(
            creds_path, GOOGLE_SCOPES)
        self._creds = flow.run_local_server(port=8401, open_browser=True)
        self._save_token()
        return self._creds

    def logout(self):
        """Clear stored credentials."""
        self._creds = None
        if os.path.exists(self._token_file):
            os.remove(self._token_file)

    def is_authenticated(self):
        """Check if we have valid or refreshable credentials."""
        if not self._creds:
            return False
        if self._creds.valid:
            return True
        if self._creds.expired and self._creds.refresh_token:
            return True
        return False


# ═══════════════════════════════════════════════════════════════
# GMAIL API CLIENT
# ═══════════════════════════════════════════════════════════════

class GmailClient:
    """Gmail API client that normalizes responses to Microsoft Graph format.

    All returned email dicts use the same field names as GraphClient:
    id, subject, bodyPreview, body, from, toRecipients, ccRecipients,
    receivedDateTime, isRead, importance, flag, hasAttachments,
    conversationId, inferenceClassification, categories, parentFolderId
    """

    def __init__(self, auth):
        self._auth = auth
        self._service = None
        self._lock = threading.RLock()
        self._label_cache = {}  # label_id -> name
        self._folder_cache = {}  # name.lower() -> label_id

    def _get_service(self):
        """Get or create the Gmail API service. Rebuilds if credentials were refreshed.
        Thread-safe — prevents concurrent service rebuilds that corrupt HTTP connections."""
        with self._lock:
            creds = self._auth.get_credentials()
            if self._service is None or getattr(self, '_last_creds_token', None) != creds.token:
                self._service = build("gmail", "v1", credentials=creds)
                self._last_creds_token = creds.token
            return self._service

    def get_me(self):
        """Get user profile info, normalized to Graph format."""
        svc = self._get_service()
        profile = self._api_call(lambda: svc.users().getProfile(userId="me").execute())
        email_addr = profile.get("emailAddress", "")
        # Try to get display name from OAuth userinfo
        try:
            from googleapiclient.discovery import build as _build
            people_svc = _build("oauth2", "v2", credentials=self._auth.get_credentials())
            info = self._api_call(lambda: people_svc.userinfo().get().execute())
            name = info.get("name", email_addr.split("@")[0])
        except Exception:
            name = email_addr.split("@")[0]
        return {
            "displayName": name,
            "mail": email_addr,
            "userPrincipalName": email_addr,
            "jobTitle": "",
            "businessPhones": [],
            "mobilePhone": "",
            "_provider": "google",
        }

    def get_emails(self, top=30, skip=0, folder="inbox",
                   before=None, after=None, search=None):
        """Fetch emails, normalized to Graph format."""
        svc = self._get_service()

        # Map folder names to Gmail label IDs
        label_map = {
            "inbox": "INBOX",
            "sentitems": "SENT",
            "drafts": "DRAFT",
            "deleteditems": "TRASH",
            "junkemail": "SPAM",
            "archive": None,  # Gmail archive = no INBOX label
        }
        label_id = label_map.get(folder.lower(), folder.upper())

        # Build query
        q_parts = []
        if search:
            q_parts.append(search)
        if before:
            q_parts.append(f"before:{before.strftime('%Y/%m/%d')}")
        if after:
            q_parts.append(f"after:{after.strftime('%Y/%m/%d')}")

        kwargs = {"userId": "me", "maxResults": top}
        if label_id:
            kwargs["labelIds"] = [label_id]
        if q_parts:
            kwargs["q"] = " ".join(q_parts)

        # Note: Gmail doesn't support skip/offset, only pageToken.
        # For simplicity, we fetch from the beginning each time.
        # The dashboard's incremental refresh uses ID dedup anyway.
        result = self._api_call(lambda: svc.users().messages().list(**kwargs).execute())
        message_ids = result.get("messages", [])

        if not message_ids:
            return {"value": []}

        # Fetch full message details (batched)
        emails = []
        for msg_ref in message_ids:
            try:
                msg = self._api_call(lambda mid=msg_ref["id"]: svc.users().messages().get(
                    userId="me", id=mid,
                    format="full"
                ).execute())
                normalized = self._normalize_message(msg)
                if normalized:
                    emails.append(normalized)
            except Exception:
                continue

        return {"value": emails}

    def _normalize_message(self, msg):
        """Convert a Gmail API message to Graph-compatible dict."""
        headers = {h["name"].lower(): h["value"]
                   for h in msg.get("payload", {}).get("headers", [])}

        # Parse sender
        from_raw = headers.get("from", "")
        from_name, from_email = self._parse_address(from_raw)

        # Parse recipients
        to_raw = headers.get("to", "")
        cc_raw = headers.get("cc", "")
        to_recips = self._parse_address_list(to_raw)
        cc_recips = self._parse_address_list(cc_raw)

        # Parse date
        internal_date = msg.get("internalDate", "0")
        try:
            dt = datetime.fromtimestamp(int(internal_date) / 1000, tz=timezone.utc)
            received_dt = dt.strftime("%Y-%m-%dT%H:%M:%S.0000000Z")
        except Exception:
            received_dt = ""

        # Labels → read status, importance, folder
        labels = set(msg.get("labelIds", []))
        is_read = "UNREAD" not in labels
        is_important = "IMPORTANT" in labels
        is_starred = "STARRED" in labels
        in_inbox = "INBOX" in labels

        # Body
        body_html, body_text = self._extract_body(msg.get("payload", {}))
        snippet = msg.get("snippet", "")

        # Attachments
        has_attachments = any(
            part.get("filename")
            for part in self._flatten_parts(msg.get("payload", {}))
            if part.get("filename")
        )

        return {
            "id": msg["id"],
            "subject": headers.get("subject", "(no subject)"),
            "bodyPreview": unescape(snippet)[:255] if snippet else "",
            "body": {
                "contentType": "html" if body_html else "text",
                "content": body_html or body_text or snippet,
            },
            "from": {
                "emailAddress": {
                    "name": from_name or from_email,
                    "address": from_email,
                }
            },
            "toRecipients": to_recips,
            "ccRecipients": cc_recips,
            "receivedDateTime": received_dt,
            "isRead": is_read,
            "importance": "high" if is_important else "normal",
            "flag": {"flagStatus": "flagged" if is_starred else "notFlagged"},
            "hasAttachments": has_attachments,
            "conversationId": msg.get("threadId", ""),
            "inferenceClassification": "focused" if in_inbox else "other",
            "categories": [],
            "parentFolderId": "inbox" if in_inbox else "archive",
            # Provider tag for dual-account support
            "_provider": "google",
            "_gmail_id": msg["id"],
            "_gmail_thread_id": msg.get("threadId", ""),
            "_gmail_labels": list(labels),
        }

    def _extract_body(self, payload):
        """Extract HTML and plain text body from Gmail payload."""
        html_body = ""
        text_body = ""

        parts = self._flatten_parts(payload)
        for part in parts:
            mime = part.get("mimeType", "")
            data = part.get("body", {}).get("data", "")
            if not data:
                continue
            try:
                decoded = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
            except Exception:
                continue
            if mime == "text/html" and not html_body:
                html_body = decoded
            elif mime == "text/plain" and not text_body:
                text_body = decoded

        return html_body, text_body

    def _flatten_parts(self, payload):
        """Recursively flatten MIME parts."""
        parts = []
        if "parts" in payload:
            for part in payload["parts"]:
                parts.extend(self._flatten_parts(part))
        else:
            parts.append(payload)
        return parts

    def _parse_address(self, raw):
        """Parse 'Name <email>' into (name, email)."""
        if not raw:
            return "", ""
        match = re.match(r'^"?([^"<]*)"?\s*<?([^>]*)>?$', raw.strip())
        if match:
            name = match.group(1).strip().strip('"')
            addr = match.group(2).strip()
            return name, addr.lower()
        return raw.strip(), raw.strip().lower()

    def _parse_address_list(self, raw):
        """Parse comma-separated addresses to Graph-format recipient list."""
        if not raw:
            return []
        result = []
        for part in raw.split(","):
            name, addr = self._parse_address(part.strip())
            if addr:
                result.append({
                    "emailAddress": {"name": name or addr, "address": addr}
                })
        return result

    # ── Actions ────────────────────────────────────────────────

    def _api_call(self, fn):
        """Execute a Gmail API call under lock to prevent concurrent HTTP requests
        from corrupting the shared connection pool (causes SSL/WinError crashes)."""
        with self._lock:
            return fn()

    def mark_as_read(self, message_id):
        """Mark message as read."""
        svc = self._get_service()
        self._api_call(lambda: svc.users().messages().modify(
            userId="me", id=message_id,
            body={"removeLabelIds": ["UNREAD"]}
        ).execute())

    def archive_email(self, message_id):
        """Archive message (remove INBOX label)."""
        svc = self._get_service()
        self._api_call(lambda: svc.users().messages().modify(
            userId="me", id=message_id,
            body={"removeLabelIds": ["INBOX"]}
        ).execute())

    def delete_email(self, message_id):
        """Trash message."""
        svc = self._get_service()
        self._api_call(lambda: svc.users().messages().trash(userId="me", id=message_id).execute())

    def reply_to_email(self, message_id, reply_body_html, extra_cc=None,
                       bcc=None, attachments=None):
        """Send a reply to a message."""
        return self._send_reply(message_id, reply_body_html, reply_all=False,
                                extra_cc=extra_cc, bcc=bcc)

    def reply_all_to_email(self, message_id, reply_body_html, extra_cc=None,
                           bcc=None, attachments=None):
        """Send a reply-all to a message."""
        return self._send_reply(message_id, reply_body_html, reply_all=True,
                                extra_cc=extra_cc, bcc=bcc)

    # ── Draft-compatible API (mirrors GraphClient interface) ──────────────────
    # Gmail has no server-side draft/send-queue concept, so these methods send
    # immediately and return a sentinel dict {"id": "_google_sent_", "_sent": True}
    # that compose.py checks to skip the MS-specific queue steps.

    def create_reply_draft(self, message_id, body_html, extra_cc=None,
                           subject=None, to_recipients=None):
        """Send reply immediately (Gmail has no draft queue). Returns sent sentinel."""
        self._send_reply(message_id, body_html, reply_all=False, extra_cc=extra_cc)
        return {"id": "_google_sent_", "_sent": True}

    def create_reply_all_draft(self, message_id, body_html, extra_cc=None,
                               subject=None, to_recipients=None):
        """Send reply-all immediately (Gmail has no draft queue). Returns sent sentinel."""
        self._send_reply(message_id, body_html, reply_all=True, extra_cc=extra_cc)
        return {"id": "_google_sent_", "_sent": True}

    def create_forward_draft(self, message_id, to_addresses, comment_html=""):
        """Forward immediately (Gmail has no draft queue). Returns sent sentinel."""
        self.forward_email(message_id, to_addresses, comment_html)
        return {"id": "_google_sent_", "_sent": True}

    def _send_reply(self, message_id, body_html, reply_all=False,
                    extra_cc=None, bcc=None):
        """Build and send a reply message."""
        svc = self._get_service()
        # Get original message for headers
        orig = self._api_call(lambda: svc.users().messages().get(
            userId="me", id=message_id, format="metadata",
            metadataHeaders=["From", "To", "Cc", "Subject", "Message-ID", "References", "In-Reply-To"]
        ).execute())
        headers = {h["name"].lower(): h["value"]
                   for h in orig.get("payload", {}).get("headers", [])}

        # Build recipients
        to_addr = headers.get("from", "")  # Reply to sender
        cc_addr = ""
        if reply_all:
            # Add original To and Cc (minus ourselves)
            profile = self._api_call(lambda: svc.users().getProfile(userId="me").execute())
            my_email = profile.get("emailAddress", "").lower()
            all_addrs = set()
            for field in ["to", "cc"]:
                for part in headers.get(field, "").split(","):
                    _, addr = self._parse_address(part.strip())
                    if addr and addr.lower() != my_email:
                        all_addrs.add(part.strip())
            cc_addr = ", ".join(all_addrs)

        if extra_cc:
            if cc_addr:
                cc_addr += ", " + ", ".join(extra_cc)
            else:
                cc_addr = ", ".join(extra_cc)

        # Build subject
        subject = headers.get("subject", "")
        if not subject.lower().startswith("re:"):
            subject = f"Re: {subject}"

        # Build MIME message
        import email.mime.text
        import email.mime.multipart
        msg = email.mime.multipart.MIMEMultipart("alternative")
        msg["To"] = to_addr
        if cc_addr:
            msg["Cc"] = cc_addr
        if bcc:
            msg["Bcc"] = ", ".join(bcc)
        msg["Subject"] = subject
        msg["In-Reply-To"] = headers.get("message-id", "")
        msg["References"] = headers.get("references", "") + " " + headers.get("message-id", "")

        # Add HTML body
        html_part = email.mime.text.MIMEText(body_html, "html")
        msg.attach(html_part)

        # Encode and send
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
        self._api_call(lambda: svc.users().messages().send(
            userId="me",
            body={"raw": raw, "threadId": orig.get("threadId", "")}
        ).execute())

    def forward_email(self, message_id, to_addresses, comment_html=""):
        """Forward a message."""
        svc = self._get_service()
        orig = self._api_call(lambda: svc.users().messages().get(
            userId="me", id=message_id, format="full"
        ).execute())
        headers = {h["name"].lower(): h["value"]
                   for h in orig.get("payload", {}).get("headers", [])}

        subject = headers.get("subject", "")
        if not subject.lower().startswith("fwd:"):
            subject = f"Fwd: {subject}"

        html_body, text_body = self._extract_body(orig.get("payload", {}))
        fwd_body = comment_html + "<br/><hr/>" + (html_body or f"<pre>{text_body}</pre>")

        import email.mime.text
        import email.mime.multipart
        msg = email.mime.multipart.MIMEMultipart("alternative")
        msg["To"] = ", ".join(to_addresses)
        msg["Subject"] = subject
        html_part = email.mime.text.MIMEText(fwd_body, "html")
        msg.attach(html_part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
        self._api_call(lambda: svc.users().messages().send(userId="me", body={"raw": raw}).execute())

    def create_draft(self, subject, body_html, to_addresses, cc_addresses=None):
        """Create a draft email."""
        svc = self._get_service()
        import email.mime.text
        import email.mime.multipart
        msg = email.mime.multipart.MIMEMultipart("alternative")
        msg["To"] = ", ".join(to_addresses) if isinstance(to_addresses, list) else to_addresses
        if cc_addresses:
            msg["Cc"] = ", ".join(cc_addresses) if isinstance(cc_addresses, list) else cc_addresses
        msg["Subject"] = subject
        html_part = email.mime.text.MIMEText(body_html, "html")
        msg.attach(html_part)
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
        draft = self._api_call(lambda: svc.users().drafts().create(
            userId="me", body={"message": {"raw": raw}}
        ).execute())
        return {"id": draft["id"]}

    def send_draft(self, draft_id):
        """Send a draft."""
        svc = self._get_service()
        self._api_call(lambda: svc.users().drafts().send(
            userId="me", body={"id": draft_id}
        ).execute())

    def get_address_book(self):
        """Get contacts from Gmail sent messages (Google People API requires separate scope)."""
        contacts = {}
        try:
            svc = self._get_service()
            result = self._api_call(lambda: svc.users().messages().list(
                userId="me", labelIds=["SENT"], maxResults=100
            ).execute())
            for msg_ref in result.get("messages", [])[:100]:
                try:
                    msg = self._api_call(lambda mid=msg_ref["id"]: svc.users().messages().get(
                        userId="me", id=mid, format="metadata",
                        metadataHeaders=["To", "Cc"]
                    ).execute())
                    headers = {h["name"].lower(): h["value"]
                               for h in msg.get("payload", {}).get("headers", [])}
                    for field in ["to", "cc"]:
                        for part in headers.get(field, "").split(","):
                            name, addr = self._parse_address(part.strip())
                            if addr and "@" in addr and addr not in contacts:
                                contacts[addr] = name or addr
                except Exception:
                    continue
        except Exception:
            pass
        return [{"name": name, "email": email} for email, name in contacts.items()]

    def get_todays_events(self):
        """Fetch today's calendar events (delegates to GoogleCalendarClient)."""
        try:
            cal = GoogleCalendarClient(self._auth)
            return cal.get_todays_events()
        except Exception:
            return []

    # ── Stub methods for Graph-compatibility ───────────────────
    # These are called by the dashboard but aren't fully applicable to Gmail.
    # They return safe defaults so the dashboard doesn't crash.

    def get_email_detail(self, message_id):
        """Get full email detail."""
        svc = self._get_service()
        msg = self._api_call(lambda: svc.users().messages().get(
            userId="me", id=message_id, format="full").execute())
        return self._normalize_message(msg)

    def is_meeting_request(self, message_id):
        """Gmail meeting invites come as calendar notifications, not email types."""
        try:
            svc = self._get_service()
            msg = self._api_call(lambda: svc.users().messages().get(
                userId="me", id=message_id, format="metadata",
                metadataHeaders=["Subject", "From"]
            ).execute())
            headers = {h["name"].lower(): h["value"]
                       for h in msg.get("payload", {}).get("headers", [])}
            subject = headers.get("subject", "").lower()
            from_addr = headers.get("from", "").lower()
            if "calendar-notification@google.com" in from_addr:
                if "canceled" in subject or "cancelled" in subject:
                    return "cancellation"
                if "invitation" in subject or "updated invitation" in subject:
                    return "request"
            return False
        except Exception:
            return False

    def get_event_times(self, message_id):
        """Not directly available from Gmail — return None."""
        return None

    def accept_event(self, message_id):
        """Gmail calendar events must be accepted via Calendar API, not Gmail."""
        pass  # Would need Calendar API integration

    def get_mail_folders(self):
        """Return Gmail labels as folder list."""
        try:
            svc = self._get_service()
            result = self._api_call(lambda: svc.users().labels().list(userId="me").execute())
            return [{"id": l["id"], "displayName": l["name"]}
                    for l in result.get("labels", [])]
        except Exception:
            return []

    def get_attachments(self, message_id):
        """Get attachment metadata for a message."""
        try:
            svc = self._get_service()
            msg = self._api_call(lambda: svc.users().messages().get(
                userId="me", id=message_id, format="full").execute())
            attachments = []
            for part in self._flatten_parts(msg.get("payload", {})):
                if part.get("filename"):
                    attachments.append({
                        "id": part.get("body", {}).get("attachmentId", ""),
                        "name": part["filename"],
                        "contentType": part.get("mimeType", ""),
                        "size": part.get("body", {}).get("size", 0),
                    })
            return attachments
        except Exception:
            return []

    def set_email_categories(self, message_id, categories):
        """Gmail doesn't have categories — use labels as approximation."""
        pass  # Could map to Gmail labels if needed

    def snooze_email(self, message_id, folder_name="Future Action"):
        """Gmail doesn't have native snooze via API — archive as workaround."""
        self.archive_email(message_id)

    def move_to_inbox(self, message_id):
        """Move message back to inbox."""
        svc = self._get_service()
        self._api_call(lambda: svc.users().messages().modify(
            userId="me", id=message_id,
            body={"addLabelIds": ["INBOX"]}
        ).execute())

    def get_snoozed_emails(self, folder_name="Future Action"):
        """Gmail doesn't have a snooze folder — return empty."""
        return []

    def get_sent_count_for_conversation(self, conversation_id):
        """Check if user has sent messages in this thread."""
        try:
            svc = self._get_service()
            result = self._api_call(lambda: svc.users().messages().list(
                userId="me", labelIds=["SENT"],
                q=f"rfc822msgid:{conversation_id}",
                maxResults=5
            ).execute())
            return len(result.get("messages", []))
        except Exception:
            return 0


# ═══════════════════════════════════════════════════════════════
# GOOGLE CALENDAR CLIENT
# ═══════════════════════════════════════════════════════════════

class GoogleCalendarClient:
    """Google Calendar API client that normalizes events to Graph format."""

    def __init__(self, auth):
        self._auth = auth
        self._service = None
        self._lock = threading.RLock()

    def _get_service(self):
        """Get or create Calendar service. Rebuilds if credentials were refreshed."""
        with self._lock:
            creds = self._auth.get_credentials()
            if self._service is None or getattr(self, '_last_creds_token', None) != creds.token:
                self._service = build("calendar", "v3", credentials=creds)
                self._last_creds_token = creds.token
            return self._service

    def _api_call(self, fn):
        """Execute a Calendar API call under lock."""
        with self._lock:
            return fn()

    def get_todays_events(self):
        """Fetch today's calendar events, normalized to Graph format."""
        svc = self._get_service()
        now = datetime.now(timezone.utc)
        start_of_day = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end_of_day = start_of_day + timedelta(days=1)

        result = self._api_call(lambda: svc.events().list(
            calendarId="primary",
            timeMin=start_of_day.isoformat(),
            timeMax=end_of_day.isoformat(),
            maxResults=50,
            singleEvents=True,
            orderBy="startTime"
        ).execute())

        events = []
        for ev in result.get("items", []):
            normalized = self._normalize_event(ev)
            if normalized:
                events.append(normalized)
        return events

    def _normalize_event(self, ev):
        """Convert Google Calendar event to Graph-compatible format."""
        # Handle all-day events (date vs dateTime)
        start = ev.get("start", {})
        end = ev.get("end", {})
        is_all_day = "date" in start and "dateTime" not in start

        if is_all_day:
            start_dt = start.get("date", "") + "T00:00:00.0000000Z"
            end_dt = end.get("date", "") + "T00:00:00.0000000Z"
        else:
            start_dt = self._to_utc(start.get("dateTime", ""))
            end_dt = self._to_utc(end.get("dateTime", ""))

        # Attendees
        attendees = []
        for att in ev.get("attendees", []):
            attendees.append({
                "emailAddress": {
                    "name": att.get("displayName", att.get("email", "")),
                    "address": att.get("email", ""),
                },
                "type": "required",
                "status": {
                    "response": self._map_response(att.get("responseStatus", "needsAction")),
                },
            })

        organizer = ev.get("organizer", {})
        location = ev.get("location", "")

        return {
            "id": ev.get("id", ""),
            "subject": ev.get("summary", "(no title)"),
            "start": {"dateTime": start_dt, "timeZone": "UTC"},
            "end": {"dateTime": end_dt, "timeZone": "UTC"},
            "isAllDay": is_all_day,
            "attendees": attendees,
            "organizer": {
                "emailAddress": {
                    "name": organizer.get("displayName", ""),
                    "address": organizer.get("email", ""),
                }
            },
            "location": {"displayName": location} if location else {},
            "isCancelled": ev.get("status") == "cancelled",
            "_provider": "google",
        }

    def _to_utc(self, dt_str):
        """Convert a datetime string to UTC ISO format."""
        if not dt_str:
            return ""
        try:
            dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
            utc_dt = dt.astimezone(timezone.utc)
            return utc_dt.strftime("%Y-%m-%dT%H:%M:%S.0000000Z")
        except Exception:
            return dt_str

    def _map_response(self, google_status):
        """Map Google Calendar response status to Graph format."""
        mapping = {
            "accepted": "accepted",
            "declined": "declined",
            "tentative": "tentativelyAccepted",
            "needsAction": "notResponded",
        }
        return mapping.get(google_status, "none")
