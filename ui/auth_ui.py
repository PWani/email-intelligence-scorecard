import os, base64, json, re, sys, tempfile, threading, webbrowser
import wx
from ._wx_common import (
    _hex, _font, askstring, showerror, showinfo, askyesno, _wx_menu_item,
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


class AuthUIMixin:
    """Authentication and Accounts"""

    def _start_auth(self):
        if not self.config.get("client_id"):
            cid = askstring(
                "Azure App Setup", "Enter your Azure Application (Client) ID:\n\n"
                "See README.md for setup instructions.", self.root)
            if not cid:
                showerror("Error", "Client ID is required.")
                self.root.Destroy(); return
            self.config["client_id"] = cid.strip()
            save_config(self.config)

        if not self.config.get("user_name"):
            uname = askstring(
                "User Setup", "Enter your full name (e.g. John Smith):\n\n"
                "This is used to detect when someone addresses\n"
                "you by name in an email.", self.root)
            if uname and uname.strip():
                self.config["user_name"] = uname.strip()
                save_config(self.config)

        self._set_status("Authenticating with Microsoft...")
        self.progress.start(15)

        def run():
            try:
                self.auth = OutlookAuth(
                    self.config["client_id"], self.config["scopes"],
                    self.config["redirect_uri"], self.config["authority"])
                self.auth.get_token()
                self.graph = GraphClient(self.auth.get_token)
                info = self.graph.get_me()
                email = info.get("mail") or info.get("userPrincipalName", "")
                name = info.get("displayName", "User")
                # Store full profile for signature generation
                self._user_profile = {
                    "displayName": name,
                    "mail": email,
                    "jobTitle": info.get("jobTitle", ""),
                    "businessPhones": info.get("businessPhones", []),
                    "mobilePhone": info.get("mobilePhone", ""),
                }
                # Use config name if set, otherwise fall back to Graph displayName
                user_name = self.config.get("user_name") or name
                self.intelligence = EmailIntelligence(user_email=email, user_name=user_name)
                wx.CallAfter(lambda: self._auth_ok(name, email))
            except Exception as e:
                err_msg = str(e)
                wx.CallAfter(lambda: self._auth_fail(err_msg))

        threading.Thread(target=run, daemon=True).start()

    def _auth_ok(self, name, email):
        self.progress.stop()
        self._ms_name = name
        self._ms_email = email
        self._update_account_label()
        self._set_status(f"Signed in as {email}")
        log.info("Signed in as %s", email)

        # Load cached address book immediately for instant autocomplete
        self._address_book = load_address_book_cache()

        # Refresh address book from API in background, then save
        def load_contacts():
            try:
                fresh = self.graph.get_address_book()
                if fresh:
                    self._address_book = fresh
                    save_address_book_cache(fresh)
            except Exception:
                pass
        threading.Thread(target=load_contacts, daemon=True).start()

        # Load today's calendar in background and start meeting alert timer
        self._todays_events = []
        self._alerted_events = set()  # event IDs already alerted
        def load_calendar():
            try:
                self._todays_events = self.graph.get_todays_events()
            except Exception:
                pass
        threading.Thread(target=load_calendar, daemon=True).start()
        self._start_meeting_alert_timer()
        self._start_auto_refresh_timer()

        # Recover any snoozed/reminded emails and scheduled sends from before app restart
        self._recover_pending_actions()

        self._refresh()

        # Try to connect Google account in background (non-blocking)
        if self.config.get("google_enabled") and _HAS_GOOGLE_MODULE:
            wx.CallLater(500, self._start_google_auth)

    # ── Google Account Integration ─────────────────────────────

    def _start_google_auth(self):
        """Attempt to authenticate Google account in background."""
        if not _HAS_GOOGLE_MODULE:
            return
        if not is_google_available():
            self._set_status(f"Google libs missing: {get_google_import_error()}")
            return

        creds_file = self.config.get("google_credentials_file") or GOOGLE_CREDS_FILE
        if not os.path.exists(creds_file):
            # No credentials file — Google not configured
            return

        self._set_status("Connecting Google account...")

        def run():
            log.info("[google] auth attempt — creds_file=%s exists=%s", creds_file, os.path.exists(creds_file))
            try:
                self.google_auth = GoogleAuth(credentials_file=creds_file)
                log.info("[google] GoogleAuth created")
                self.google_auth.get_credentials()
                log.info("[google] Credentials obtained")
                self.google_client = GmailClient(self.google_auth)
                log.info("[google] GmailClient created")
                info = self.google_client.get_me()
                log.info("[google] get_me returned: %s", info.get('mail', 'NO MAIL'))
                self._google_email = info.get("mail", "")
                self._google_name = info.get("displayName", "")

                # Merge Google contacts into address book
                try:
                    g_contacts = self.google_client.get_address_book()
                    if g_contacts:
                        existing = {c["email"].lower() for c in self._address_book}
                        for c in g_contacts:
                            if c["email"].lower() not in existing:
                                self._address_book.append(c)
                except Exception:
                    pass

                # Merge Google calendar events
                try:
                    g_events = self.google_client.get_todays_events()
                    if g_events:
                        self._todays_events.extend(g_events)
                except Exception:
                    pass

                wx.CallAfter(self._google_auth_ok)
            except FileNotFoundError as e:
                log.warning("[google] FileNotFoundError: %s", e)
                wx.CallAfter(lambda: self._set_status(str(e)[:100]))
            except Exception as e:
                import traceback
                log.error("[google] %s: %s\n%s", type(e).__name__, e, traceback.format_exc())
                wx.CallAfter(lambda: self._set_status(
                    f"Google auth skipped: {str(e)[:80]}"))

        threading.Thread(target=run, daemon=True).start()

    def _google_auth_ok(self):
        """Called after successful Google authentication."""
        self._update_account_label()
        self._set_status(
            f"Connected: {self._ms_email} + {self._google_email}")
        # Fetch Google emails and merge into list
        self._refresh_google_emails()

    def _refresh_google_emails(self):
        """Fetch emails from Google and merge into the main email list."""
        if not self.google_client:
            return

        _after_val = self.after_entry.GetValue()
        _before_val = self.before_entry.GetValue()

        def run():
            try:
                log.info("[google] Fetching emails...")
                after_dt = self._parse_date(_after_val)
                before_dt = self._parse_date(_before_val)
                params = {"top": 50, "folder": self.current_folder,
                          "before": before_dt, "after": after_dt}
                if self.search_query:
                    params["search"] = self.search_query
                result = self.google_client.get_emails(**params)
                fetched = result.get("value", [])
                log.info("[google] Fetched %d emails", len(fetched))
                if fetched:
                    if not self.intelligence:
                        log.warning("[google] intelligence not ready, skipping scoring")
                        self._google_emails = fetched
                    else:
                        enriched = self.intelligence.process_emails(fetched)
                        log.info("[google] Running auto-archive on %d emails", len(enriched))
                        enriched = self._apply_auto_archive(enriched)
                        log.info("[google] After auto-archive: %d emails remaining", len(enriched))
                        self._google_emails = enriched
                    wx.CallAfter(self._merge_google_and_render)
            except Exception as e:
                log.warning("Google refresh error: %s", e)

        threading.Thread(target=run, daemon=True).start()

    def _merge_google_and_render(self):
        """Merge Google emails into the main list, inserting/removing only changed cards."""
        if not self._google_emails:
            return
        log.info("[google] Merging %d Google emails into list", len(self._google_emails))

        # Build the new combined + sorted list
        ms_emails = [e for e in self.emails if e.get("_provider") != "google"]
        combined = ms_emails + self._google_emails
        if self.sort_by_priority:
            combined.sort(
                key=lambda e: (e.get("_intel", {}).get("score", 0),
                               e.get("receivedDateTime", "")),
                reverse=True)
        else:
            combined.sort(key=lambda e: e.get("receivedDateTime", ""), reverse=True)

        # Work out what's new vs what was already rendered
        existing_ids = set(self._card_refs.keys())
        new_ids = {e.get("id") for e in self._google_emails}
        truly_new = [e for e in self._google_emails if e.get("id") not in existing_ids]
        # Cards that were previously rendered but are no longer in Google emails
        # (auto-archived or removed) — destroy them
        removed_ids = [eid for eid in list(self._card_refs.keys())
                       if eid not in {e.get("id") for e in combined}]

        self.emails = combined

        if not truly_new and not removed_ids:
            # Nothing changed — just update stats
            self._update_stats()
            return

        self._list_scroll.Freeze()
        try:
            # Remove stale cards first
            for eid in removed_ids:
                if eid in self._card_refs:
                    card = self._card_refs[eid][0]
                    self.list_inner_sizer.Detach(card)
                    self._panel_to_id.pop(id(card), None)
                    card.Destroy()
                    del self._card_refs[eid]

            if not truly_new and not removed_ids:
                pass  # nothing else to do
            elif set(e.get("id") for e in combined) == set(self._card_refs.keys()):
                # Same cards after removals, potentially different order — rebuild
                self._rebuild_list(self._get_split_emails())
            else:
                # Insert new cards at correct positions
                for em in truly_new:
                    eid = em.get("id")
                    target_idx = next(
                        (i for i, e in enumerate(combined) if e.get("id") == eid), None)
                    if target_idx is not None:
                        self._render_card(em, insert_at=target_idx)

            self._list_scroll.FitInside()
            # If cards were removed from top, scroll to top to close the gap
            if removed_ids:
                self._list_scroll.Scroll(0, 0)
        finally:
            self._list_scroll.Thaw()

        self._update_stats()
        if self.selected_email_id and self.selected_email_id in self._card_refs:
            self._highlight_card(self.selected_email_id)

    def _get_client_for_email(self, email_dict):
        """Return the appropriate API client (Graph or Gmail) for an email."""
        if email_dict and email_dict.get("_provider") == "google":
            return self.google_client
        return self.graph

    def _api_for(self, eid):
        """Return the correct API client for an email ID.
        Falls back to ID format detection when email not in self.emails
        (e.g. removed during auto-archive batch or offline queue replay)."""
        for e in self.emails:
            if e.get("id") == eid:
                if e.get("_provider") == "google" and self.google_client:
                    return self.google_client
                return self.graph
        # Not found in self.emails — check Google email cache
        if self.google_client:
            for e in getattr(self, '_google_emails', []):
                if e.get("id") == eid:
                    return self.google_client
            # Fallback: Gmail IDs are short hex strings (typically 16 chars)
            if eid and len(eid) <= 16:
                try:
                    int(eid, 16)
                    return self.google_client
                except (ValueError, TypeError):
                    pass
        return self.graph

    def _show_account_menu(self):
        """Show dropdown menu for account filter: Microsoft, Google, All, Settings."""
        menu = wx.Menu()

        # Check mark for current selection
        def _label(name, key):
            return f"  ✓  {name}" if self._account_filter == key else f"      {name}"

        _wx_menu_item(menu, _label("Microsoft", "microsoft"), lambda: self._set_account_filter("microsoft"))
        _wx_menu_item(menu, _label("Google", "google"), lambda: self._set_account_filter("google"))
        _wx_menu_item(menu, _label("All", "all"), lambda: self._set_account_filter("all"))
        menu.AppendSeparator()
        _wx_menu_item(menu, "      Settings...", self._show_account_manager)

        # Position below the button
        try:
            x = self._acct_btn.GetScreenPosition().x
            y = self._acct_btn.GetScreenPosition().y + self._acct_btn.GetSize().height
            self.root.PopupMenu(menu)
        except Exception:
            self.root.PopupMenu(menu)

    def _set_account_filter(self, mode):
        """Switch account filter and re-render."""
        self._account_filter = mode
        labels = {"microsoft": "Microsoft", "google": "Google", "all": "All"}
        self._acct_btn.SetLabel(f"👥 {labels[mode]} ▾")
        # Re-filter and re-render the email list
        self._apply_account_filter_and_render()

    def _apply_account_filter_and_render(self):
        """Re-render the email list with current account filter applied."""
        self._render_list()
        self._update_stats()

    def _show_account_menu(self):
        """Show dropdown menu for account filter: Microsoft, Google, All, Settings."""
        menu = wx.Menu()
        chk = lambda key: "  \u2713  " if self._account_filter == key else "      "
        _wx_menu_item(menu, f"{chk('microsoft')}Microsoft", lambda: self._set_account_filter("microsoft"))
        _wx_menu_item(menu, f"{chk('google')}Google", lambda: self._set_account_filter("google"))
        _wx_menu_item(menu, f"{chk('all')}All", lambda: self._set_account_filter("all"))
        menu.AppendSeparator()
        _wx_menu_item(menu, "      Settings...", self._show_account_manager)
        try:
            x = self._acct_btn.GetScreenPosition().x
            y = self._acct_btn.GetScreenPosition().y + self._acct_btn.GetSize().height
            self.root.PopupMenu(menu)
        except Exception:
            self.root.PopupMenu(menu)

    def _set_account_filter(self, mode):
        """Switch account filter and re-render."""
        self._account_filter = mode
        labels = {"microsoft": "Microsoft", "google": "Google", "all": "All"}
        self._acct_btn.SetLabel(f"\U0001f465 {labels[mode]} \u25be")
        self._apply_account_filter_and_render()

    def _apply_account_filter_and_render(self):
        """Filter the current email list by account and re-render."""
        if self._account_filter == "microsoft":
            filtered = [e for e in self.emails if e.get("_provider") != "google"]
        elif self._account_filter == "google":
            filtered = [e for e in self.emails if e.get("_provider") == "google"]
        else:
            filtered = self.emails
        self._render_list(filtered)
        self._update_stats(filtered)

    def _update_account_label(self):
        """Update the user label in the topbar to show connected accounts."""
        parts = []
        if self._ms_name:
            parts.append(self._ms_name)
        if self._google_name:
            parts.append(f"G: {self._google_name}")
        label = " | ".join(parts) if parts else "Not signed in"
        self.user_label.SetLabel(f"\U0001f464 {label}")

    def _show_account_manager(self):
        """Show account management dialog for adding/removing Google account."""
        win = wx.Dialog(self.root, style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER)
        # win.title("Account Manager")
        # win.geometry("480x360")
        # win.resizable(False, False)
        # win.configure(bg=C["bg"])
        
        

        wx.StaticText(win, label="Connected Accounts")
        # Microsoft account
        ms_frame = wx.Panel(win)
        ms_frame.Show()
        wx.StaticText(ms_frame, label="Microsoft 365")
        ms_status = self._ms_email if self._ms_email else "Not connected"
        wx.StaticText(ms_frame, label=ms_status)

        # Google account
        g_frame = wx.Panel(win)
        g_frame.Show()
        _stxt = wx.StaticText(g_frame, label="Google (Gmail + Calendar)", bg=C["bg_card"])
        _stxt.SetFont(wx.Font(11, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, faceName=FONT_BOLD))

        g_status_text = self._google_email if self._google_email else "Not connected"
        g_status = wx.StaticText(g_frame, label=g_status_text)
        g_status.Show()

        g_btn_frame = wx.Panel(g_frame)
        g_btn_frame.Show()

        def _connect_google():
            if not _HAS_GOOGLE_MODULE:
                showerror("Missing Libraries", "Google API libraries not installed.\n\n"
                    "Run in terminal:\n"
                    "pip install google-auth google-auth-oauthlib google-api-python-client",
                    parent=win)
                return
            if not is_google_available():
                showerror("Missing Libraries", f"Google import error:\n{get_google_import_error()}",
                    parent=win)
                return

            # Check for credentials file
            creds_file = self.config.get("google_credentials_file") or GOOGLE_CREDS_FILE
            if not os.path.exists(creds_file):
                # Ask user to provide it
                dlg = wx.FileDialog(win, "Select Google OAuth credentials JSON",
                    wildcard="JSON files (*.json)|*.json", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
                path = dlg.GetPath() if dlg.ShowModal() == wx.ID_OK else None
                dlg.Destroy()

        close_btn = wx.Button(win, label="Close")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: win.Destroy())

    def _start_auto_refresh_timer(self):
        """Smart refresh inbox every 5 minutes. Also refreshes today's calendar."""
        def auto_refresh():
            try:
                if self.graph and not getattr(self, '_loading_more', False):
                    self._is_auto_refresh = True
                    self._refresh()
                    # Also refresh calendar events
                    def reload_cal():
                        try:
                            self._todays_events = self.graph.get_todays_events()
                            # Also refresh Google calendar if connected
                            if self.google_client:
                                try:
                                    g_events = self.google_client.get_todays_events()
                                    if g_events:
                                        self._todays_events.extend(g_events)
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    threading.Thread(target=reload_cal, daemon=True).start()
                    # Also refresh Google emails
                    if self.google_client:
                        self._refresh_google_emails()
            except Exception:
                pass
            wx.CallLater(300000, auto_refresh)  # 5 minutes = 300,000ms
        wx.CallLater(300000, auto_refresh)


    def _auth_fail(self, err):
        self.progress.stop()
        self._set_status(f"Auth error: {err}")
        if askyesno("Auth Failed", f"{err}\n\nRetry? (Yes to retry, No to quit)"):
            self.config["client_id"] = ""; save_config(self.config); self._start_auth()
        else:
            self.root.Destroy()
