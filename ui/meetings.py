import os, base64, json, re, sys, tempfile, threading, webbrowser
import wx
from ._wx_common import (
    _hex, _font, askstring, showerror, showinfo, askyesno,
    FONT, FONT_BOLD, CONFIG_DIR, CONFIG_FILE, TOKEN_CACHE_FILE,
    ADDRESS_BOOK_FILE, OFFLINE_QUEUE_FILE, SCORING_RULES_FILE,
    _APP_ICON_B64, DEFAULT_CONFIG, ensure_config_dir, load_config,
    save_config, load_address_book_cache, save_address_book_cache,
    C, P_ICON, CAT_ICON, TIMEZONES, DEFAULT_TZ, log,
    OutlookAuth, GraphClient, OfflineQueue, is_network_error,
    DEFAULT_SCORING_RULES, load_scoring_rules, save_scoring_rules,
    _deep_merge, strip_html, strip_outlook_banners, _build_word_pattern,
    keyword_in_text, html_to_readable_text, extract_text,
    extract_latest_reply, EmailIntelligence, SpellChecker,
    EmailAutocomplete, detect_renderer, EmailWebView, RENDERER,
    _HAS_GOOGLE_MODULE, msal, requests,
    datetime, timedelta, timezone, unescape,
)
try:
    from ..core.google_client import (
        GoogleAuth, GmailClient, GoogleCalendarClient,
        is_google_available, get_google_import_error, GOOGLE_CREDS_FILE,
    )
except ImportError:
    pass


class MeetingsMixin:
    """Meetings and Images"""

    def _apply_meeting_button(self, meeting_type):
        """Swap the archive/accept button based on meeting message type.
        meeting_type: False, 'request', 'cancellation', 'response', or True (legacy)."""
        if meeting_type == 'request' or meeting_type is True:
            self._archive_accept_btn.SetLabel('✅ Accept')
            self._archive_accept_btn.Bind(wx.EVT_BUTTON, lambda e: self._accept_invite())
            self._archive_accept_btn.SetBackgroundColour(_hex(C.get('green', '#16A34A')))
            self._archive_accept_btn.SetForegroundColour(wx.WHITE)
            self._current_action_is_accept = True
        else:
            self._archive_accept_btn.SetLabel('📦 Archive')
            self._archive_accept_btn.Bind(wx.EVT_BUTTON, lambda e: self._archive())
            self._archive_accept_btn.SetBackgroundColour(wx.NullColour)
            self._archive_accept_btn.SetForegroundColour(wx.NullColour)
            self._current_action_is_accept = False

    def _show_event_time(self, event_times):
        """Display event time bar for meeting requests."""
        time_str = self._fmt_event_time(event_times)
        if time_str:
            if self._is_event_past(event_times):
                self.d_event_time.SetLabel(f'{time_str}  ⚠ PAST EVENT')
                self.d_event_time.SetForegroundColour(_hex('#DC2626'))
                self._event_time_frame.SetBackgroundColour(_hex('#FEF2F2'))
            else:
                self.d_event_time.SetLabel(time_str)
                self.d_event_time.SetForegroundColour(_hex('#1E40AF'))
                self._event_time_frame.SetBackgroundColour(_hex('#EFF6FF'))
            self._event_time_frame.Show()
            self._detail_panel.Layout()

    def _reset_html_frame(self):
        """Recreate the EmailWebView for a clean state."""
        try:
            self.body.Destroy()
        except Exception:
            pass
        try:
            self.body = EmailWebView(self._body_container)
            sizer = self._body_container.GetSizer()
            if sizer:
                sizer.Clear()
                sizer.Add(self.body, 1, wx.EXPAND)
            else:
                s = wx.BoxSizer(wx.VERTICAL)
                s.Add(self.body, 1, wx.EXPAND)
                self._body_container.SetSizer(s)
            self._body_container.Layout()
        except Exception:
            pass

    def _load_images(self):
        """Un-block remote images in the current email via JS — no full re-render needed."""
        self._load_images_frame.Hide()
        try:
            self.body.load_images()
        except Exception:
            # Fallback: full re-render without blocking
            if getattr(self, '_last_raw_content', None):
                self._render_email_body(self._last_raw_content,
                                        self._last_content_type,
                                        block_images=False)

    def _safe_sender_images(self):
        """Add current sender to safe image sender list and reload images."""
        if not self.selected_email_id:
            return
        em = next((e for e in self.emails if e.get("id") == self.selected_email_id), None)
        if not em:
            return
        sender = em.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        sender_name = em.get("from", {}).get("emailAddress", {}).get("name", sender)
        if not sender:
            return
        rules = self.intelligence.rules
        sis = rules.setdefault("safe_image_senders", {"enabled": True, "entries": []})
        entries = sis.setdefault("entries", [])
        if sender not in entries:
            entries.append(sender)
            save_scoring_rules(rules)
        self._load_images()
        self._set_status(f"Images always loaded for {sender_name}")

    def _is_safe_image_sender(self, email_dict):
        """Check if sender is on the safe image sender list or is a VIP sender."""
        rules = self.intelligence.rules
        sender = email_dict.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        # VIP senders always get images
        vip_entries = [v.lower() for v in rules.get("vip_senders", {}).get("entries", [])]
        if any(v in sender for v in vip_entries):
            return True
        # Explicit safe image sender list
        sis = rules.get("safe_image_senders", {})
        if not sis.get("enabled", True):
            return False
        entries = [s.lower() for s in sis.get("entries", [])]
        return sender in entries

    # ── Filters / Search / Sort ───────────────────────────────

    def _on_search(self, event=None):
        q = self.search_ctrl.GetValue().strip()
        self.search_query = q if q and q!="Search emails..." else None
        self._refresh()

    def _clear_search(self):
        """Clear search and return to normal inbox view."""
        self.search_query = None
        self.search_ctrl.SetValue("")
        self.search_ctrl.Clear()
        self.search_ctrl.SetValue("")
        self.search_entry.selection_clear()
        self.root.SetFocus()
        self._full_refresh()

    def _on_folder_change(self):
        selected = self.folder_var.GetStringSelection()
        # Save current folder's emails to cache before switching
        if self.emails:
            self._folder_email_cache[self.current_folder] = list(self.emails)
        # Determine new folder
        if selected == "snoozed / reminded":
            new_folder = "_snoozed"
        elif selected == "send queue":
            new_folder = "_send_queue"
        else:
            new_folder = selected
        self.current_folder = new_folder
        # Restore cached emails for new folder, or start empty
        cached = self._folder_email_cache.get(new_folder, [])
        self.emails = list(cached)
        self.current_skip = len(self.emails)
        self._all_loaded = False
        # Render cached content immediately (may be empty)
        self.list_inner_sizer.Clear(delete_windows=True)
        self._card_refs = {}
        if self.emails:
            self._render_list(self.emails, append=False)
        self._update_stats()
        # Try network refresh (will replace cached data on success, preserve on failure)
        self._full_refresh()

    def _clear_filters(self):
        self.after_entry.SetValue(""); self.before_entry.SetValue("")
        self.after_entry.SetValue("YYYY-MM-DD")
        self.before_entry.SetValue("YYYY-MM-DD")
        self.search_query = None; self.search_ctrl.SetValue("")
        self._full_refresh()

    def _toggle_sort(self):
        self.sort_by_priority = not self.sort_by_priority
        self.sort_btn.SetLabel("Sort: Priority" if self.sort_by_priority else "Sort: Date")
        if self.sort_by_priority:
            self.emails.sort(key=lambda e: (e.get("_intel",{}).get("score",0),
                                             e.get("receivedDateTime","")), reverse=True)
        else:
            self.emails.sort(key=lambda e: e.get("receivedDateTime",""), reverse=True)
        self.list_inner_sizer.Clear(delete_windows=True)
        self._render_list()

    def _show_about(self):
        """Show the About dialog."""
        win = wx.Dialog(self.root, title="About Email Intelligence Dashboard",
                        style=wx.DEFAULT_DIALOG_STYLE, size=(460, 340))
        win.SetBackgroundColour(wx.Colour(255, 255, 255))

        outer = wx.BoxSizer(wx.VERTICAL)

        # Title
        title = wx.StaticText(win, label="Email Intelligence Dashboard")
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        outer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 12)

        version = wx.StaticText(win, label="Version 0.34")
        version.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        version.SetForegroundColour(wx.Colour(107, 114, 128))
        outer.Add(version, 0, wx.BOTTOM | wx.ALIGN_CENTER, 8)

        outer.Add(wx.StaticLine(win), 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 20)

        legal_text = (
            "Copyright \u00a9 2026 Email Intelligence Scorecard Contributors.\n"
            "All rights reserved.\n\n"
            "Licensed under the MIT License.\n"
            "See LICENSE file for details.\n\n"
            "Provided \"as is\" without warranty of any kind,\n"
            "express or implied."
        )
        legal = wx.StaticText(win, label=legal_text, style=wx.ALIGN_CENTER)
        legal.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        legal.SetForegroundColour(wx.Colour(107, 114, 128))
        outer.Add(legal, 0, wx.ALL | wx.ALIGN_CENTER, 12)

        outer.Add(wx.StaticLine(win), 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 20)

        author = wx.StaticText(win, label="Email Intelligence Scorecard")
        author.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL))
        outer.Add(author, 0, wx.ALL | wx.ALIGN_CENTER, 8)

        close_btn = wx.Button(win, label="Close")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: win.Destroy())
        outer.Add(close_btn, 0, wx.BOTTOM | wx.ALIGN_CENTER, 12)

        win.SetSizer(outer)
        win.Layout()
        win.Centre()
        win.ShowModal()
        win.Destroy()

    def _sign_out(self):
        if askyesno("Sign Out", "Sign out and clear credentials?"):
            if self.auth: self.auth.logout()
            if self.google_auth:
                try:
                    self.google_auth.logout()
                except Exception:
                    pass
            self.google_auth = None
            self.google_client = None
            self._google_email = ""
            self._google_name = ""
            self._google_emails = []
            self.config["client_id"] = ""
            self.config["google_enabled"] = False
            save_config(self.config)
            self.root.Destroy()

    def _auto_archive_past_events(self):
        """Background: check meeting emails for past events and auto-archive them."""
        if self.current_folder != "inbox":
            return
        # Gather candidates — emails that look like meeting invites
        # Skip Google emails since their Calendar API doesn't return event times
        candidates = [e for e in self.emails
                      if e.get("_provider") != "google"
                      and (e.get("_intel", {}).get("category") in ("meeting_invite", "meeting")
                      or any(kw in (e.get("subject") or "").lower()
                             for kw in ["accepted:", "declined:", "tentative:", "canceled:", "updated:"]))]
        if not candidates:
            return

        log.info("[auto-archive-events] Checking %d meeting candidate(s)", len(candidates))

        # Pre-resolve API clients while emails are still in self.emails
        candidate_tasks = [(e, self._api_for(e.get("id"))) for e in candidates]

        def run():
            archived_ids = []
            for em, client in candidate_tasks:
                eid = em.get("id")
                subj = (em.get("subject") or "")[:50]
                if not eid:
                    continue
                try:
                    # Check if it's actually a meeting message
                    if "_is_meeting_request" not in em:
                        em["_is_meeting_request"] = client.is_meeting_request(eid)
                    if not em["_is_meeting_request"]:
                        continue
                    # Get event times
                    if "_event_times" not in em or em["_event_times"] is None:
                        em["_event_times"] = client.get_event_times(eid)
                    if self._is_event_past(em.get("_event_times")):
                        client.archive_email(eid)
                        archived_ids.append(eid)
                        log.info("[auto-archive-events] Archived past event: subj='%s'", subj)
                except Exception as exc:
                    log.error("[auto-archive-events] Error checking: subj='%s' err=%s", subj, exc)
                    continue

            if archived_ids:
                log.info("[auto-archive-events] Archived %d past meeting(s)", len(archived_ids))
                def update_ui():
                    self.emails = [e for e in self.emails if e.get("id") not in archived_ids]
                    self._render_list()
                    self._update_stats()
                    self._list_scroll.Scroll(0, 0)
                    wx.CallAfter(lambda: wx.CallLater(50, self._update_scroll_region))
                    self._set_status(
                        f"Auto-archived {len(archived_ids)} past meeting invite(s)")
                wx.CallAfter(update_ui)

        threading.Thread(target=run, daemon=True).start()

        threading.Thread(target=run, daemon=True).start()

    # ── Meeting Alert System ──────────────────────────────────

    def _start_meeting_alert_timer(self):
        """Check every 30 seconds for upcoming meetings within 60 seconds."""
        def check():
            try:
                self._check_upcoming_meetings()
            except Exception:
                pass
            # Re-schedule every 30 seconds
            wx.CallLater(30000, check)
        # First check after 5 seconds (let calendar load)
        wx.CallLater(5000, check)

    def _check_upcoming_meetings(self):
        """Check if any meeting starts within 60 seconds and show alert."""
        if not getattr(self, '_todays_events', None):
            return
        now = datetime.now(timezone.utc)
        for event in self._todays_events:
            eid = event.get("id", "")
            if eid in self._alerted_events:
                continue
            if event.get("isCancelled"):
                continue
            if event.get("isAllDay"):
                continue
            start = event.get("start", {})
            start_str = start.get("dateTime", "")
            if not start_str:
                continue
            try:
                # Times are in UTC (requested via Prefer header)
                start_dt = datetime.fromisoformat(start_str.replace("Z", ""))
                start_dt = start_dt.replace(tzinfo=timezone.utc)

                diff = (start_dt - now).total_seconds()
                if 0 <= diff <= 60:
                    self._alerted_events.AddPage(eid)
                    wx.CallAfter(lambda ev=event, d=int(diff): self._show_meeting_alert(ev, d))
            except Exception:
                continue

    def _show_meeting_alert(self, event, seconds_until):
        """Show a popup alert for an upcoming meeting."""
        win = wx.Dialog(self.root, style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER)
        # win.title("📅 Meeting Starting Soon")
        # win.configure(bg="#FFFFFF")
        # win.resizable(False, False)
        win.attributes("-topmost", True)
        

        # Try to flash/bring to attention
        try:
            win.bell()
            self.root.SetFocus()
            win.SetFocus()
        except Exception:
            pass

        container = wx.Panel(win)
        container.Show()

        # Alert icon and title
        wx.StaticText(container, label="⏰ Meeting Starting Now!")
        # Separator
        wx.Panel(container).Show()

        # Meeting subject
        subject = event.get("subject", "(No subject)")
        _make_static_text(container, subject, FONT_BOLD, 12)

        # Time
        start = event.get("start", {})
        end = event.get("end", {})
        start_str = start.get("dateTime", "")
        end_str = end.get("dateTime", "")
        time_text = ""
        try:
            # Times are UTC, convert to user's selected timezone
            s_utc = datetime.fromisoformat(start_str.replace("Z", "")).replace(tzinfo=timezone.utc)
            e_utc = datetime.fromisoformat(end_str.replace("Z", "")).replace(tzinfo=timezone.utc)
            tz_offset = self._get_tz_offset()
            s_local = s_utc + tz_offset
            e_local = e_utc + tz_offset
            tz_label = self.tz_var.GetStringSelection()
            time_text = f"{s_local.strftime('%I:%M %p')} – {e_local.strftime('%I:%M %p')} {tz_label}"
        except Exception:
            pass
        if time_text:
            _make_static_text(container, time_text, FONT, 11)
        # Location
        location = event.get("location", {}).get("displayName", "")
        if location:
            _make_static_text(container, f"📍 {location}", FONT, 10)
        # Separator
        wx.Panel(container).Show()

        # Attendees
        attendees = event.get("attendees", [])
        if attendees:
            _make_static_text(container, "Attendees:", FONT_BOLD, 10)
            att_frame = wx.Panel(container)
            att_frame.Show()

            for i, att in enumerate(attendees[:12]):  # Show up to 12
                ea = att.get("emailAddress", {})
                name = ea.get("name", "")
                email = ea.get("address", "")
                status = att.get("status", {}).get("response", "")
                status_icon = {"accepted": "✅", "declined": "❌", "tentative": "❔",
                               "organizer": "👑"}.get(status, "⬜")
                att_type = att.get("type", "")
                if att_type == "resource":
                    continue
                display = f"{status_icon} {name}" if name else f"{status_icon} {email}"
                _make_static_text(att_frame, display, FONT, 9)

            remaining = len(attendees) - 12
            if remaining > 0:
                _stxt = wx.StaticText(att_frame, label=f"  +{remaining} more")
                _stxt.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL, faceName=FONT))

        # Organizer
        organizer = event.get("organizer", {}).get("emailAddress", {})
        org_name = organizer.get("name", "")
        if org_name:
            wx.StaticText(container, label=f"Organized by: {org_name}")
        # Dismiss button
        dismiss_btn = wx.Button(container, label="Dismiss")
        dismiss_btn.Bind(wx.EVT_BUTTON, lambda e: win.Destroy())

        # Auto-dismiss after 90 seconds
        wx.CallLater(90000, lambda: win.EndModal(wx.ID_OK) if win else None)

        # Center on screen
        win.Update()
        w, h = win.GetSize()
        sw, sh = wx.GetDisplaySize()
        x = (sw - w) // 2
        y = (sh - h) // 2
        # win.geometry(f"+{x}+{y}")

    # ── Utilities ─────────────────────────────────────────────

