# Graph API
import json, os, re, threading
from datetime import datetime, timedelta, timezone
from urllib.parse import quote as urlquote
import requests
from .config import CONFIG_DIR, OFFLINE_QUEUE_FILE, log

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

EMAIL_FIELDS = (
    "id,subject,bodyPreview,body,from,toRecipients,ccRecipients,"
    "receivedDateTime,isRead,importance,flag,hasAttachments,"
    "conversationId,inferenceClassification,categories,parentFolderId"
)

NETWORK_ERRORS = (
    "NameResolutionError", "ConnectionError", "Max retries",
    "getaddrinfo failed", "NewConnectionError", "ConnectTimeoutError",
    "ReadTimeout", "SSLError", "ProxyError", "ConnectionResetError",
    "ConnectionAborted", "10054", "10053", "forcibly closed",
    "RemoteDisconnected", "aborted by the software",
    "timed out", "TimeoutError", "Read timed out",
    "operation timed out", "ReadTimeoutError",
    "WRONG_VERSION_NUMBER", "_ssl.c", "CERTIFICATE_VERIFY_FAILED",
    "SSL:", "[SSL]", "EOF occurred", "504 Server Error", "502 Server Error",
    "503 Server Error", "Gateway Timeout",
)

def is_network_error(err_str):
    """Check if an error string indicates a network/connectivity issue."""
    return any(ne in err_str for ne in NETWORK_ERRORS)


class OfflineQueue:
    """Persistent disk-based queue for actions taken while offline.
    Actions are serialized to JSON and replayed when connectivity returns."""

    def __init__(self, path=OFFLINE_QUEUE_FILE):
        self._path = path
        self._lock = threading.Lock()
        os.makedirs(os.path.dirname(path), exist_ok=True)

    def _load(self):
        try:
            with open(self._path, "r") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return []

    def _save(self, items):
        tmp = self._path + ".tmp"
        with open(tmp, "w") as f:
            json.dump(items, f)
        os.replace(tmp, self._path)

    def enqueue(self, action, **kwargs):
        """Add an action to the offline queue. Returns the queued item."""
        item = {"action": action, "ts": datetime.now(timezone.utc).isoformat(), **kwargs}
        with self._lock:
            items = self._load()
            items.append(item)
            self._save(items)
        return item

    def peek_all(self):
        """Return all queued items without removing them."""
        with self._lock:
            return self._load()

    def clear(self):
        """Remove all items from the queue."""
        with self._lock:
            self._save([])

    def remove_completed(self, count):
        """Remove the first `count` items (successfully replayed)."""
        with self._lock:
            items = self._load()
            self._save(items[count:])

    def is_empty(self):
        with self._lock:
            return len(self._load()) == 0

    @property
    def count(self):
        with self._lock:
            return len(self._load())


class GraphClient:
    def __init__(self, get_token_fn):
        self._get_token = get_token_fn
        self._folder_cache = {}  # name.lower() -> id

    def _headers(self):
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    def _get(self, url, params=None, extra_headers=None):
        h = self._headers()
        if extra_headers:
            h.update(extra_headers)
        r = requests.get(url, headers=h, params=params, timeout=30)
        r.raise_for_status()
        return r.json()

    def _post(self, url, json_data=None):
        r = requests.post(url, headers=self._headers(), json=json_data, timeout=60)
        r.raise_for_status()
        return r

    def _patch(self, url, json_data):
        r = requests.patch(url, headers=self._headers(), json=json_data, timeout=60)
        r.raise_for_status()
        return r

    def _msg_url(self, message_id, suffix=""):
        """Build a safe URL for a message, encoding the message ID for special chars."""
        safe_id = urlquote(message_id, safe="")
        base = f"{GRAPH_BASE}/me/messages/{safe_id}"
        return f"{base}/{suffix}" if suffix else base

    def get_me(self):
        return self._get(f"{GRAPH_BASE}/me",
                         params={"$select": "displayName,mail,userPrincipalName,"
                                            "jobTitle,businessPhones,mobilePhone"})

    def get_address_book(self):
        """Fetch contacts from People API, Contacts API, and recent sent items.
        Returns list of dicts: [{name, email}, ...]"""
        contacts = {}  # email -> name, deduped

        # 1. People API — frequently contacted people (most relevant)
        try:
            result = self._get(f"{GRAPH_BASE}/me/people",
                               params={"$top": 200,
                                       "$select": "displayName,emailAddresses,personType"})
            for person in result.get("value", []):
                name = person.get("displayName", "")
                for ea in person.get("emailAddresses", []):
                    addr = (ea.get("address") or "").lower().strip()
                    if addr and "@" in addr:
                        contacts[addr] = name
        except Exception:
            pass

        # 2. Contacts API — full address book
        try:
            result = self._get(f"{GRAPH_BASE}/me/contacts",
                               params={"$top": 500,
                                       "$select": "displayName,emailAddresses"})
            for contact in result.get("value", []):
                name = contact.get("displayName", "")
                for ea in contact.get("emailAddresses", []):
                    addr = (ea.get("address") or "").lower().strip()
                    if addr and "@" in addr and addr not in contacts:
                        contacts[addr] = name
        except Exception:
            pass

        # 3. Recent sent items — last 100 recipients
        try:
            result = self._get(f"{GRAPH_BASE}/me/mailFolders/sentitems/messages",
                               params={"$top": 100,
                                       "$select": "toRecipients,ccRecipients",
                                       "$orderby": "sentDateTime desc"})
            for msg in result.get("value", []):
                for recip_list in [msg.get("toRecipients", []), msg.get("ccRecipients", [])]:
                    for recip in recip_list:
                        ea = recip.get("emailAddress", {})
                        addr = (ea.get("address") or "").lower().strip()
                        name = ea.get("name") or ""
                        if addr and "@" in addr and addr not in contacts:
                            contacts[addr] = name
        except Exception:
            pass

        return [{"name": name, "email": email} for email, name in contacts.items()]

    def get_todays_events(self):
        """Fetch today's calendar events with attendees. Times returned in UTC."""
        try:
            now = datetime.now(timezone.utc)
            start_of_day = now.replace(hour=0, minute=0, second=0, microsecond=0)
            end_of_day = start_of_day + timedelta(days=1)
            result = self._get(f"{GRAPH_BASE}/me/calendarView",
                               params={"startDateTime": start_of_day.strftime("%Y-%m-%dT%H:%M:%S.0000000Z"),
                                       "endDateTime": end_of_day.strftime("%Y-%m-%dT%H:%M:%S.0000000Z"),
                                       "$top": 50,
                                       "$select": "id,subject,start,end,isAllDay,attendees,organizer,location,isCancelled",
                                       "$orderby": "start/dateTime"},
                               extra_headers={"Prefer": 'outlook.timezone="UTC"'})
            return result.get("value", [])
        except Exception:
            return []

    def get_emails(self, top=30, skip=0, folder="inbox",
                   before=None, after=None, search=None):
        # $search requires /me/messages, not /me/mailFolders/.../messages
        if search:
            url = f"{GRAPH_BASE}/me/messages"
        else:
            url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
        params = {
            "$top": top, "$skip": skip,
            "$orderby": "receivedDateTime desc",
            "$select": EMAIL_FIELDS,
        }
        filters = []
        if before:
            filters.append(f"receivedDateTime lt {before.strftime('%Y-%m-%dT%H:%M:%SZ')}")
        if after:
            filters.append(f"receivedDateTime ge {after.strftime('%Y-%m-%dT%H:%M:%SZ')}")
        if filters:
            params["$filter"] = " and ".join(filters)
        if search:
            params.pop("$filter", None)
            params.pop("$orderby", None)
            params.pop("$skip", None)
            params["$search"] = f'"{search}"'
        return self._get(url, params=params)

    def get_email_detail(self, message_id):
        return self._get(self._msg_url(message_id),
                         params={"$select": EMAIL_FIELDS})

    def is_meeting_request(self, message_id):
        """Check if a message is a meeting/event message via @odata.type or Google Calendar patterns.
        Returns False if not a meeting message, 'request' if it's an actionable invite,
        'cancellation' or 'response' for non-actionable meeting messages."""
        try:
            # First try with meetingMessageType (works for native Exchange meetings)
            try:
                result = self._get(self._msg_url(message_id),
                                   params={"$select": "id,subject,meetingMessageType,internetMessageId"})
            except Exception:
                # meetingMessageType may fail on non-eventMessage items — retry without it
                result = self._get(self._msg_url(message_id),
                                   params={"$select": "id,subject,internetMessageId"})
            odata_type = result.get("@odata.type", "")

            # Native Exchange/Outlook meeting messages
            if "eventMessage" in str(odata_type):
                mtype = result.get("meetingMessageType", "")
                if mtype == "meetingRequest":
                    return "request"
                elif mtype in ("meetingCancelled", "meetingDeclined"):
                    return "cancellation"
                elif mtype in ("meetingAccepted", "meetingTentativelyAccepted"):
                    return "response"
                elif mtype == "":
                    # meetingMessageType is empty — this happens for:
                    # 1. Recurring meeting instances
                    # 2. Meetings where attendee was added after creation
                    # 3. Some Exchange/O365 edge cases
                    # Use the event navigation property on eventMessage (most reliable),
                    # then fall back to subject-based search
                    try:
                        # Method 1: Direct event navigation (works for most eventMessages)
                        event_data = None
                        try:
                            safe_mid = urlquote(message_id, safe="")
                            event_data = self._get(
                                f"{GRAPH_BASE}/me/messages/{safe_mid}/microsoft.graph.eventMessage/event",
                                params={"$select": "id,subject,responseStatus"})
                        except Exception:
                            pass

                        # Method 2: Fall back to subject search if navigation failed
                        if not event_data:
                            subject = result.get("subject", "")
                            clean_subj = subject
                            for prefix in ["Accepted: ", "Declined: ", "Tentative: ",
                                           "Canceled: ", "Cancelled: ", "Updated: "]:
                                if clean_subj.startswith(prefix):
                                    clean_subj = clean_subj[len(prefix):]
                            event_data = self._find_event_by_subject(clean_subj)

                        if event_data:
                            resp = event_data.get("responseStatus", {})
                            resp_type = resp.get("response", "none")
                            if resp_type in ("none", "notResponded", "organizer"):
                                return "request"
                            elif resp_type in ("accepted", "tentativelyAccepted"):
                                return "response"
                            elif resp_type == "declined":
                                return "cancellation"
                            else:
                                return "request"  # Default to actionable
                    except Exception:
                        pass
                    # If we can't determine status, default to request (show Accept)
                    # since it's an eventMessage — better to show Accept than hide it
                    return "request"
                else:
                    return "response"

            # Google Calendar invites — not native eventMessage but still meeting invites
            subject = result.get("subject", "")
            msg_id = result.get("internetMessageId", "")
            is_gcal = msg_id.startswith("<calendar-")
            if is_gcal:
                subj_lower = subject.lower()
                if "canceled" in subj_lower or "cancelled" in subj_lower:
                    return "cancellation"
                elif any(p in subj_lower for p in
                         ["invitation:", "updated invitation", "invite:", "new event:"]):
                    return "request"
                elif "@" in subject and ("am " in subject or "pm " in subject or
                                         "AM " in subject or "PM " in subject):
                    # Subject contains date/time pattern like "@ Mon Mar 2, 2026 2pm"
                    return "request"

            # Fallback: check for .ics attachment or calendar-like subject from non-Google senders
            if not is_gcal:
                subj_lower = (subject or "").lower()
                # Check subject patterns that indicate meeting invites from other systems
                if any(p in subj_lower for p in ["invitation:", "invite:", "meeting request:"]):
                    if "@" in subject or "am " in subj_lower or "pm " in subj_lower:
                        return "request"

            return False
        except Exception:
            return False

    def get_event_times(self, message_id):
        """Get start/end times for a meeting request email by matching the calendar event."""
        utc_header = {"Prefer": 'outlook.timezone="UTC"'}
        try:
            msg = self._get(self._msg_url(message_id),
                            params={"$select": "id,subject,receivedDateTime"})
            subject = msg.get("subject", "")
            received = msg.get("receivedDateTime", "")
            for prefix in ["Accepted: ", "Declined: ", "Tentative: ", "Canceled: ",
                           "Cancelled: ", "Updated: ", "RE: ", "FW: "]:
                if subject.startswith(prefix):
                    subject = subject[len(prefix):]

            # Handle Google Calendar subject format
            for gcal_prefix in ["Updated invitation with note: ", "Updated invitation: ",
                                "Invitation: ", "New event: "]:
                if subject.startswith(gcal_prefix):
                    subject = subject[len(gcal_prefix):]
                    break
            at_idx = subject.find(" @ ")
            if at_idx > 0:
                subject = subject[:at_idx].strip()

            # Strip common date prefixes like "2/9 " or "02/09 " from subject
            import re as _re
            subject_clean = _re.sub(r'^\d{1,2}/\d{1,2}\s+', '', subject).strip()

            # Method 1: Use the event navigation property (most reliable, handles [brackets])
            event_list = []
            try:
                safe_mid = urlquote(message_id, safe="")
                event_data = self._get(
                    f"{GRAPH_BASE}/me/messages/{safe_mid}/microsoft.graph.eventMessage/event",
                    params={"$select": "id,subject,start,end,isAllDay"},
                    extra_headers=utc_header)
                if event_data and event_data.get("id"):
                    event_list = [event_data]
            except Exception:
                pass

            # Method 2: Fall back to robust subject search
            if not event_list:
                safe_subject = subject.replace("'", "''")
                safe_subject_clean = subject_clean.replace("'", "''")

                # Try exact match on original subject
                try:
                    events = self._get(f"{GRAPH_BASE}/me/events",
                                       params={"$filter": f"subject eq '{safe_subject}'",
                                               "$top": 5,
                                               "$select": "id,subject,start,end,isAllDay"},
                                       extra_headers=utc_header)
                    event_list = events.get("value", [])
                except Exception:
                    pass

                # Try exact match on cleaned subject (without date prefix)
                if not event_list and safe_subject_clean != safe_subject:
                    try:
                        events = self._get(f"{GRAPH_BASE}/me/events",
                                           params={"$filter": f"subject eq '{safe_subject_clean}'",
                                                   "$top": 5,
                                                   "$select": "id,subject,start,end,isAllDay"},
                                           extra_headers=utc_header)
                        event_list = events.get("value", [])
                    except Exception:
                        pass

                # Try startsWith for partial match
                if not event_list:
                    try:
                        events = self._get(f"{GRAPH_BASE}/me/events",
                                           params={"$filter": f"startsWith(subject, '{safe_subject_clean}')",
                                                   "$top": 5,
                                                   "$select": "id,subject,start,end,isAllDay"},
                                           extra_headers=utc_header)
                        event_list = events.get("value", [])
                    except Exception:
                        pass

                # Try calendarView around received date (broad search)
                if not event_list and received:
                    try:
                        recv_dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
                        start_range = (recv_dt - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%S.0000000")
                        end_range = (recv_dt + timedelta(days=90)).strftime("%Y-%m-%dT%H:%M:%S.0000000")
                        events = self._get(f"{GRAPH_BASE}/me/calendarView",
                                           params={"startDateTime": start_range,
                                                   "endDateTime": end_range,
                                                   "$top": 100,
                                                   "$select": "id,subject,start,end,isAllDay"},
                                           extra_headers=utc_header)
                        # Find best match by subject similarity
                        subj_lower = subject_clean.lower()
                        for ev in events.get("value", []):
                            ev_subj = (ev.get("subject") or "").lower()
                            if subj_lower in ev_subj or ev_subj in subj_lower:
                                event_list = [ev]
                                break
                    except Exception:
                        pass

            if event_list:
                ev = event_list[0]
                return {
                    "start": ev.get("start", {}),
                    "end": ev.get("end", {}),
                    "isAllDay": ev.get("isAllDay", False),
                }
        except Exception:
            pass
        return None

    def mark_as_read(self, message_id):
        self._patch(self._msg_url(message_id), {"isRead": True})

    def archive_email(self, message_id):
        archive_id = self._get_or_create_folder("Archive")
        try:
            # Mark as read before moving to archive
            try:
                self._patch(self._msg_url(message_id), {"isRead": True})
            except Exception:
                pass  # best effort — still archive even if mark-read fails
            self._post(self._msg_url(message_id, "move"),
                       {"destinationId": archive_id})
        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code == 404:
                return  # already moved/deleted
            raise

    def delete_email(self, message_id):
        try:
            self._post(self._msg_url(message_id, "move"),
                       {"destinationId": "deleteditems"})
        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code == 404:
                return  # already moved/deleted
            raise

    def reply_to_email(self, message_id, reply_body_html, extra_cc=None,
                       subject=None, to_recipients=None):
        msg = {"body": {"contentType": "HTML", "content": reply_body_html}}
        if extra_cc:
            msg["ccRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in extra_cc if a.strip()]
        if subject:
            msg["subject"] = subject
        if to_recipients:
            msg["toRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in to_recipients if a.strip()]
        self._post(self._msg_url(message_id, "reply"), {"message": msg})

    def reply_all_to_email(self, message_id, reply_body_html, extra_cc=None,
                           subject=None, to_recipients=None):
        msg = {"body": {"contentType": "HTML", "content": reply_body_html}}
        if extra_cc:
            msg["ccRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in extra_cc if a.strip()]
        if subject:
            msg["subject"] = subject
        if to_recipients:
            msg["toRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in to_recipients if a.strip()]
        self._post(self._msg_url(message_id, "replyAll"), {"message": msg})

    def create_reply_draft(self, message_id, body_html, extra_cc=None,
                           subject=None, to_recipients=None):
        """Create a reply draft with original body quoted. Returns draft message.
        Creates draft first (preserving quoted thread), then prepends reply text."""
        msg = {}
        if extra_cc:
            msg["ccRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in extra_cc if a.strip()]
        if subject:
            msg["subject"] = subject
        if to_recipients:
            msg["toRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in to_recipients if a.strip()]
        # Create reply draft — Graph auto-includes quoted thread in body
        r = self._post(self._msg_url(message_id, "createReply"), {"message": msg} if msg else {})
        draft = r.json()
        # Prepend our reply text above the quoted thread
        draft_id = draft.get("id")
        if draft_id and body_html:
            existing_body = draft.get("body", {}).get("content", "")
            # Insert reply before the existing quoted content
            combined = self._prepend_reply_to_body(body_html, existing_body)
            self._patch(self._msg_url(draft_id),
                        {"body": {"contentType": "HTML", "content": combined}})
            draft["body"] = {"contentType": "HTML", "content": combined}
        return draft

    def create_reply_all_draft(self, message_id, body_html, extra_cc=None,
                               subject=None, to_recipients=None):
        """Create a reply-all draft with original body quoted. Returns draft message.
        Creates draft first (preserving quoted thread), then prepends reply text."""
        msg = {}
        if extra_cc:
            msg["ccRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in extra_cc if a.strip()]
        if subject:
            msg["subject"] = subject
        if to_recipients:
            msg["toRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in to_recipients if a.strip()]
        # Create reply-all draft — Graph auto-includes quoted thread in body
        r = self._post(self._msg_url(message_id, "createReplyAll"), {"message": msg} if msg else {})
        draft = r.json()
        # Prepend our reply text above the quoted thread
        draft_id = draft.get("id")
        if draft_id and body_html:
            existing_body = draft.get("body", {}).get("content", "")
            combined = self._prepend_reply_to_body(body_html, existing_body)
            self._patch(self._msg_url(draft_id),
                        {"body": {"contentType": "HTML", "content": combined}})
            draft["body"] = {"contentType": "HTML", "content": combined}
        return draft

    def _prepend_reply_to_body(self, reply_html, existing_body):
        """Insert reply HTML at the top of the existing body content."""
        import re
        # Try to insert after <body> tag
        m = re.search(r'(<body[^>]*>)', existing_body, re.IGNORECASE)
        if m:
            insert_pos = m.end()
            return existing_body[:insert_pos] + reply_html + existing_body[insert_pos:]
        # Fallback: prepend before existing content
        return reply_html + existing_body

    def create_forward_draft(self, message_id, to_addresses, comment_html=""):
        """Create a forward draft with original body + attachments. Returns draft message.
        Creates draft first (preserving original body), then prepends comment if any."""
        if isinstance(to_addresses, str):
            to_addresses = [to_addresses]
        recipients = [{"emailAddress": {"address": a.strip()}}
                      for a in to_addresses if a.strip() and "@" in a.strip()]
        # Create forward draft WITHOUT body — preserves original content + attachments
        payload = {"toRecipients": recipients}
        r = self._post(self._msg_url(message_id, "createForward"), payload)
        draft = r.json()

        # If there's a comment/signature, prepend it to the draft body
        if comment_html and comment_html.strip():
            draft_id = draft.get("id", "")
            if draft_id:
                existing_body = draft.get("body", {}).get("content", "")
                new_body = (
                    f'<div style="font-family:Aptos,Aptos_MSFontService,-apple-system,Roboto,Arial,'
                    f'Helvetica,sans-serif;font-size:12pt;color:rgb(33,33,33)">'
                    f'{comment_html}</div><br>{existing_body}'
                )
                self._patch(self._msg_url(draft_id),
                           {"body": {"contentType": "HTML", "content": new_body}})
                draft["body"] = {"contentType": "HTML", "content": new_body}
        return draft

    def forward_email(self, message_id, to_addresses, comment_html=""):
        """Forward an email to one or more addresses with optional comment."""
        if isinstance(to_addresses, str):
            to_addresses = [to_addresses]
        recipients = [{"emailAddress": {"address": a.strip()}}
                      for a in to_addresses if a.strip() and "@" in a.strip()]
        self._post(self._msg_url(message_id, "forward"), {
            "comment": comment_html,
            "toRecipients": recipients
        })

    def get_sent_count_for_conversation(self, conversation_id):
        """Check if user has already sent messages in this conversation thread."""
        try:
            result = self._get(f"{GRAPH_BASE}/me/mailFolders/sentitems/messages",
                               params={"$filter": f"conversationId eq '{conversation_id}'",
                                       "$select": "id",
                                       "$top": 1,
                                       "$count": "true"})
            return len(result.get("value", []))
        except Exception:
            return 0

    def get_mail_folders(self):
        result = self._get(f"{GRAPH_BASE}/me/mailFolders", params={"$top": 50})
        return result.get("value", [])

    def get_attachments(self, message_id):
        """Get list of attachments for an email. Returns list of {id, name, size, contentType, contentBytes}."""
        result = self._get(self._msg_url(message_id, "attachments"),
                           params={"$top": 50})
        return result.get("value", [])

    def _find_event_by_subject(self, subject):
        """Find a calendar event by subject, handling special characters like [, ].
        Returns the first matching event dict or None."""
        if not subject:
            return None
        safe_subj = subject.replace("'", "''")

        # Method 1: exact match with $filter
        try:
            events = self._get(f"{GRAPH_BASE}/me/events",
                               params={"$filter": f"subject eq '{safe_subj}'",
                                       "$top": 5,
                                       "$select": "id,subject,responseStatus"})
            if events.get("value"):
                return events["value"][0]
        except Exception:
            pass

        # Method 2: startsWith (handles partial matches)
        try:
            events = self._get(f"{GRAPH_BASE}/me/events",
                               params={"$filter": f"startsWith(subject, '{safe_subj}')",
                                       "$top": 5,
                                       "$select": "id,subject,responseStatus"})
            if events.get("value"):
                return events["value"][0]
        except Exception:
            pass

        # Method 3: $search (full-text, handles special chars like brackets)
        # Extract meaningful words from subject, skip punctuation
        import re as _re
        words = _re.findall(r'[A-Za-z0-9]+', subject)
        if words:
            search_q = " ".join(words[:4])  # Use first 4 words max
            try:
                events = self._get(f"{GRAPH_BASE}/me/events",
                                   params={"$search": f'"{search_q}"',
                                           "$top": 5,
                                           "$select": "id,subject,responseStatus"})
                # Verify it's actually a match (search is fuzzy)
                for ev in events.get("value", []):
                    ev_subj = ev.get("subject", "").lower()
                    if subject.lower() in ev_subj or ev_subj in subject.lower():
                        return ev
                    # Check if key words match
                    ev_words = set(_re.findall(r'[a-z0-9]+', ev_subj))
                    subj_words = set(_re.findall(r'[a-z0-9]+', subject.lower()))
                    if len(subj_words & ev_words) >= len(subj_words) * 0.7:
                        return ev
            except Exception:
                pass

        # Method 4: calendarView (broad search around now for recently received invites)
        try:
            now = datetime.now(timezone.utc)
            start = (now - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
            end = (now + timedelta(days=90)).strftime("%Y-%m-%dT%H:%M:%SZ")
            events = self._get(f"{GRAPH_BASE}/me/calendarView",
                               params={"startDateTime": start,
                                       "endDateTime": end,
                                       "$top": 100,
                                       "$select": "id,subject,responseStatus"})
            for ev in events.get("value", []):
                if ev.get("subject", "").strip().lower() == subject.strip().lower():
                    return ev
        except Exception:
            pass

        return None

    def accept_event(self, message_id):
        """Accept a meeting invite. Uses multiple strategies to find and accept the correct event.
        For recurring meetings, uses /events/{seriesMasterId}/instances to get proper instance IDs
        that work with the /accept endpoint (calendarView IDs often don't work)."""
        def _dbg(msg):
            log.debug("[accept_event] %s", msg)

        # Get message details
        try:
            msg = self._get(self._msg_url(message_id),
                            params={"$select": "id,subject,meetingMessageType,internetMessageId,receivedDateTime"})
        except Exception:
            msg = self._get(self._msg_url(message_id),
                            params={"$select": "id,subject,internetMessageId,receivedDateTime"})
        subject = msg.get("subject", "")
        mtype = msg.get("meetingMessageType", "")
        msg_id = msg.get("internetMessageId", "")
        received = msg.get("receivedDateTime", "")
        odata_type = msg.get("@odata.type", "")

        _dbg(f"Subject: {subject}")
        _dbg(f"meetingMessageType: '{mtype}', @odata.type: '{odata_type}', internetMessageId: '{msg_id[:40]}...'")

        # Don't try to accept cancellations or responses
        if mtype and mtype not in ("meetingRequest", ""):
            raise Exception(f"Cannot accept this meeting message (type: {mtype}). "
                          f"It may have been cancelled or is a response notification.")

        # Strip common prefixes from subject for event matching
        clean_subject = subject
        for prefix in ["Accepted: ", "Declined: ", "Tentative: ", "Canceled: ",
                       "Cancelled: ", "Updated: "]:
            if clean_subject.startswith(prefix):
                clean_subject = clean_subject[len(prefix):]

        # Handle Google Calendar subject format
        is_gcal = msg_id.startswith("<calendar-")
        if is_gcal:
            for gcal_prefix in ["Updated invitation with note: ", "Updated invitation: ",
                                "Invitation: ", "New event: "]:
                if clean_subject.startswith(gcal_prefix):
                    clean_subject = clean_subject[len(gcal_prefix):]
                    break
            at_idx = clean_subject.find(" @ ")
            if at_idx > 0:
                clean_subject = clean_subject[:at_idx].strip()

        is_event_message = "eventMessage" in str(odata_type)
        _dbg(f"Entering strategies: clean_subject='{clean_subject}', is_event_message={is_event_message}")

        # === Strategy 1: Navigate from eventMessage → event → accept ===
        # This is the most reliable path per Microsoft docs
        if is_event_message:
            try:
                safe_mid = urlquote(message_id, safe="")
                event_data = self._get(
                    f"{GRAPH_BASE}/me/messages/{safe_mid}/microsoft.graph.eventMessage/event",
                    params={"$select": "id,subject,type,seriesMasterId"})
                event_id = event_data.get("id")
                event_type = event_data.get("type", "singleInstance")
                master_id = event_data.get("seriesMasterId")
                _dbg(f"Strategy 1: event_id={event_id}, type={event_type}, master={master_id}")

                if event_id:
                    safe_eid = urlquote(event_id, safe="")
                    try:
                        self._post(f"{GRAPH_BASE}/me/events/{safe_eid}/accept",
                                   {"sendResponse": True})
                        _dbg("Strategy 1: Accept succeeded directly")
                        return
                    except requests.exceptions.HTTPError as e:
                        status = e.response.status_code if e.response else "?"
                        _dbg(f"Strategy 1: Direct accept failed ({status}), trying alternatives")

                        if status == 400:
                            # For recurring meetings: use /instances endpoint on series master
                            resolve_master = master_id or (event_id if event_type == "seriesMaster" else None)
                            if resolve_master:
                                if self._accept_via_instances(resolve_master, _dbg):
                                    return
                            # Try accepting via eventMessage navigation path directly
                            try:
                                self._post(
                                    f"{GRAPH_BASE}/me/messages/{safe_mid}/microsoft.graph.eventMessage/event/accept",
                                    {"sendResponse": True})
                                _dbg("Strategy 1: Accept via eventMessage/event/accept succeeded")
                                return
                            except Exception as e2:
                                _dbg(f"Strategy 1: eventMessage/event/accept failed: {e2}")
                        else:
                            raise
            except requests.exceptions.HTTPError as he:
                # Only re-raise if NOT a 400 — 400s should fall through to Strategy 2
                if he.response is not None and he.response.status_code != 400:
                    raise
                _dbg(f"Strategy 1: eventMessage/event GET returned 400, falling through to Strategy 2")
            except Exception as e:
                _dbg(f"Strategy 1 failed entirely: {e}")

        # === Strategy 2: Find event by subject via /me/events, use /instances for recurring ===
        try:
            safe_subj = clean_subject.replace("'", "''")
            events = self._get(f"{GRAPH_BASE}/me/events",
                               params={"$filter": f"subject eq '{safe_subj}'",
                                       "$top": 5,
                                       "$select": "id,subject,type,seriesMasterId"})
            for ev in events.get("value", []):
                ev_id = ev.get("id")
                ev_type = ev.get("type", "singleInstance")
                _dbg(f"Strategy 2: Found event {ev_id}, type={ev_type}")

                if ev_type == "seriesMaster":
                    # Use /instances to get proper IDs that work with /accept
                    if self._accept_via_instances(ev_id, _dbg):
                        return
                else:
                    # Single instance — accept directly
                    safe_eid = urlquote(ev_id, safe="")
                    try:
                        self._post(f"{GRAPH_BASE}/me/events/{safe_eid}/accept",
                                   {"sendResponse": True})
                        _dbg("Strategy 2: Accept succeeded on single instance")
                        return
                    except requests.exceptions.HTTPError as e:
                        if e.response and e.response.status_code == 400:
                            _dbg(f"Strategy 2: 400 on single instance, continuing")
                            continue
                        raise

            # Try startsWith match too
            if not events.get("value"):
                events = self._get(f"{GRAPH_BASE}/me/events",
                                   params={"$filter": f"startsWith(subject, '{safe_subj}')",
                                           "$top": 5,
                                           "$select": "id,subject,type,seriesMasterId"})
                for ev in events.get("value", []):
                    ev_id = ev.get("id")
                    ev_type = ev.get("type", "singleInstance")
                    if ev_type == "seriesMaster":
                        if self._accept_via_instances(ev_id, _dbg):
                            return
                    else:
                        safe_eid = urlquote(ev_id, safe="")
                        try:
                            self._post(f"{GRAPH_BASE}/me/events/{safe_eid}/accept",
                                       {"sendResponse": True})
                            _dbg("Strategy 2: Accept via startsWith match")
                            return
                        except Exception:
                            continue
        except requests.exceptions.HTTPError as he:
            _dbg(f"Strategy 2: HTTPError {he.response.status_code if he.response else '?'}, falling through")
        except Exception as e:
            _dbg(f"Strategy 2 failed: {e}")

        # === Strategy 3: calendarView search as last resort ===
        try:
            if received:
                recv_dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
            else:
                recv_dt = datetime.now(timezone.utc)
            start = (recv_dt - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
            end = (recv_dt + timedelta(days=90)).strftime("%Y-%m-%dT%H:%M:%SZ")
            events = self._get(f"{GRAPH_BASE}/me/calendarView",
                               params={"startDateTime": start, "endDateTime": end,
                                       "$top": 200,
                                       "$select": "id,subject,type,seriesMasterId,responseStatus"})
            for ev in events.get("value", []):
                if ev.get("subject", "").strip().lower() == clean_subject.strip().lower():
                    ev_id = ev.get("id")
                    master_id = ev.get("seriesMasterId")
                    _dbg(f"Strategy 3: calendarView match, id={ev_id}, master={master_id}")

                    # For recurring: use /instances on the series master (calendarView IDs often fail)
                    if master_id:
                        if self._accept_via_instances(master_id, _dbg):
                            return

                    # Try direct accept (works for single instances)
                    safe_eid = urlquote(ev_id, safe="")
                    try:
                        self._post(f"{GRAPH_BASE}/me/events/{safe_eid}/accept",
                                   {"sendResponse": True})
                        _dbg("Strategy 3: Accept succeeded")
                        return
                    except Exception as e:
                        _dbg(f"Strategy 3: Accept failed on {ev_id}: {e}")
                        continue
        except requests.exceptions.HTTPError as he:
            _dbg(f"Strategy 3: HTTPError {he.response.status_code if he.response else '?'}")
        except Exception as e:
            _dbg(f"Strategy 3 failed: {e}")

        _dbg("ALL strategies failed")
        raise Exception(
            f"Could not accept '{clean_subject}'. This may be a meeting added directly "
            f"to your calendar without a formal invite. Try accepting in Outlook directly.")

    def _accept_via_instances(self, series_master_id, _dbg=None):
        """Accept the next upcoming unaccepted instance of a recurring event.
        Uses /events/{id}/instances which returns proper instance IDs that work with /accept.
        Returns True on success, False on failure."""
        try:
            safe_master = urlquote(series_master_id, safe="")
            now = datetime.now(timezone.utc)
            start = (now - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
            end = (now + timedelta(days=90)).strftime("%Y-%m-%dT%H:%M:%SZ")
            instances = self._get(
                f"{GRAPH_BASE}/me/events/{safe_master}/instances",
                params={"startDateTime": start, "endDateTime": end,
                        "$top": 20,
                        "$select": "id,subject,start,responseStatus"})
            if _dbg:
                _dbg(f"_accept_via_instances: Got {len(instances.get('value', []))} instances for master {series_master_id}")
            for inst in instances.get("value", []):
                resp_status = inst.get("responseStatus", {}).get("response", "none")
                inst_id = inst.get("id")
                if _dbg:
                    _dbg(f"  Instance {inst_id[:40]}... response={resp_status}")
                if resp_status in ("none", "notResponded", "tentativelyAccepted"):
                    safe_iid = urlquote(inst_id, safe="")
                    try:
                        self._post(f"{GRAPH_BASE}/me/events/{safe_iid}/accept",
                                   {"sendResponse": True})
                        if _dbg:
                            _dbg(f"  Accepted instance {inst_id[:40]}!")
                        return True
                    except Exception as e:
                        if _dbg:
                            _dbg(f"  Accept failed on instance: {e}")
                        continue
            # If all instances are already accepted, accept the series master directly
            try:
                self._post(f"{GRAPH_BASE}/me/events/{safe_master}/accept",
                           {"sendResponse": True})
                if _dbg:
                    _dbg("  Accepted series master directly")
                return True
            except Exception as e:
                if _dbg:
                    _dbg(f"  Series master accept also failed: {e}")
        except Exception as e:
            if _dbg:
                _dbg(f"_accept_via_instances failed: {e}")
        return False

    def _get_or_create_folder(self, name):
        key = name.lower()
        if key in self._folder_cache:
            return self._folder_cache[key]
        for f in self.get_mail_folders():
            self._folder_cache[f["displayName"].lower()] = f["id"]
        if key in self._folder_cache:
            return self._folder_cache[key]
        r = self._post(f"{GRAPH_BASE}/me/mailFolders", {"displayName": name})
        fid = r.json()["id"]
        self._folder_cache[key] = fid
        return fid

    def snooze_email(self, message_id, folder_name="Future Action"):
        """Move email to Future Action folder (visible in Outlook but hidden from this app)."""
        folder_id = self._get_or_create_folder(folder_name)
        self._post(self._msg_url(message_id, "move"),
                   {"destinationId": folder_id})

    def move_to_inbox(self, message_id):
        """Move email back to inbox (un-snooze)."""
        self._post(self._msg_url(message_id, "move"),
                   {"destinationId": "inbox"})

    def get_snoozed_emails(self, folder_name="Future Action"):
        """Get emails in the Future Action folder."""
        try:
            folder_id = self._get_or_create_folder(folder_name)
            result = self._get(
                f"{GRAPH_BASE}/me/mailFolders/{folder_id}/messages",
                params={"$top": 100, "$orderby": "receivedDateTime desc",
                        "$select": "id,subject,from,receivedDateTime,isRead,flag,bodyPreview,hasAttachments,toRecipients,ccRecipients,categories"})
            return result.get("value", [])
        except Exception:
            return []

    def set_email_categories(self, message_id, categories):
        """Set categories on an email (used to store snooze metadata)."""
        self._patch(self._msg_url(message_id),
                    {"categories": categories})

    def create_draft(self, subject, body_html, to_addresses, cc_addresses=None):
        """Create a draft email. Returns the draft message object."""
        msg = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": a.strip()}}
                             for a in to_addresses if a.strip() and "@" in a.strip()]
        }
        if cc_addresses:
            msg["ccRecipients"] = [{"emailAddress": {"address": a.strip()}}
                                    for a in cc_addresses if a.strip() and "@" in a.strip()]
        r = self._post(f"{GRAPH_BASE}/me/messages", msg)
        return r.json()

    def move_to_send_queue(self, message_id):
        """Move a draft to the Send Queue folder."""
        folder_id = self._get_or_create_folder("Send Queue")
        self._post(self._msg_url(message_id, "move"),
                   {"destinationId": folder_id})

    def get_send_queue(self):
        """Get all messages in the Send Queue folder."""
        try:
            folder_id = self._get_or_create_folder("Send Queue")
            result = self._get(
                f"{GRAPH_BASE}/me/mailFolders/{folder_id}/messages",
                params={"$top": 50, "$orderby": "createdDateTime desc",
                        "$select": "id,subject,from,toRecipients,categories,createdDateTime"})
            return result.get("value", [])
        except Exception:
            return []

    def send_draft(self, message_id):
        """Send an existing draft message."""
        self._post(self._msg_url(message_id, "send"))


