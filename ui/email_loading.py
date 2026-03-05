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


class EmailLoadingMixin:
    """Email Loading"""

    def _refresh(self):
        """Smart refresh: if we have emails, do incremental fetch for new ones only.
        Falls back to full reload on folder change, search, or first load."""
        log.info("[refresh] folder=%s emails=%d search=%s",
                 self.current_folder, len(self.emails) if self.emails else 0,
                 bool(self.search_query))
        if self.search_query or not self.emails or not self.graph:
            # Full reload needed for search, empty state, or no auth
            self._full_refresh()
            return

        # Incremental: fetch only emails newer than our newest
        self._incremental_refresh()

    def _full_refresh(self):
        """Complete reload — fetches from scratch. Preserves existing data if offline."""
        has_cached = bool(self.emails)
        if not has_cached:
            # Only wipe UI if nothing to show
            self.list_inner_sizer.Clear(delete_windows=True)
            self._card_refs = {}
        self.current_skip = 0
        self._all_loaded = False
        self._load_more_btn.Show()
        if self.current_folder == "_snoozed":
            self._load_snoozed_emails()
        elif self.current_folder == "_send_queue":
            self._load_send_queue_emails()
        elif self.search_query:
            self._load_emails()
        else:
            self._load_emails_batch(self._initial_load_target)
        # Also refresh Google emails if connected
        if self.google_client and self.current_folder not in ("_snoozed", "_send_queue"):
            self._refresh_google_emails()

    def _load_snoozed_emails(self):
        """Load snoozed/reminded emails from Future Action folder."""
        self._set_status("Loading snoozed emails...")
        self._loading_more = True
        self.progress.start(15)

        def run():
            try:
                snoozed = self.graph.get_snoozed_emails()
                # Tag each email with its snooze/remind info from categories
                for em in snoozed:
                    cats = em.get("categories", [])
                    snooze_info = ""
                    for cat in cats:
                        if cat.startswith("snooze:"):
                            try:
                                dt = datetime.fromisoformat(cat[7:])
                                local = self._utc_to_local(dt)
                                snooze_info = f"⏰ Returns {local.strftime('%b %d %I:%M %p')}"
                            except Exception:
                                snooze_info = "⏰ Snoozed"
                        elif cat.startswith("remind:"):
                            try:
                                dt = datetime.fromisoformat(cat[7:])
                                local = self._utc_to_local(dt)
                                snooze_info = f"🔔 Remind {local.strftime('%b %d %I:%M %p')}"
                            except Exception:
                                snooze_info = "🔔 Reminded"
                    em["_snooze_info"] = snooze_info

                enriched = self.intelligence.process_emails(snoozed)
                self.emails = enriched
                self._all_loaded = True
                self._folder_email_cache["_snoozed"] = list(self.emails)
                wx.CallAfter(lambda: self._render_list(enriched, append=False))
                wx.CallAfter(self._update_load_more_btn)
            except Exception as e:
                err_msg = str(e)
                wx.CallAfter(lambda: self._load_err(err_msg))
            finally:
                self._loading_more = False

        threading.Thread(target=run, daemon=True).start()

    def _load_send_queue_emails(self):
        """Load scheduled/queued emails from Send Queue folder."""
        self._set_status("Loading send queue...")
        self._loading_more = True
        self.progress.start(15)

        def run():
            try:
                queued = self.graph.get_send_queue()
                for em in queued:
                    cats = em.get("categories", [])
                    for cat in cats:
                        if cat.startswith("send_at:"):
                            try:
                                dt = datetime.fromisoformat(cat.split(":", 1)[1])
                                local = self._utc_to_local(dt)
                                em["_schedule_info"] = f"📤 Sends {local.strftime('%b %d %I:%M %p')}"
                            except Exception:
                                em["_schedule_info"] = "📤 Queued"

                enriched = self.intelligence.process_emails(queued)
                self.emails = enriched
                self._all_loaded = True
                self._folder_email_cache["_send_queue"] = list(self.emails)
                wx.CallAfter(lambda: self._render_list(enriched, append=False))
                wx.CallAfter(self._update_load_more_btn)
            except Exception as e:
                err_msg = str(e)
                wx.CallAfter(lambda: self._load_err(err_msg))
            finally:
                self._loading_more = False

        threading.Thread(target=run, daemon=True).start()

    def _incremental_refresh(self):
        """Fetch only new emails since the most recent one we have, and prepend them."""
        # For special folders, always do full refresh
        if self.current_folder in ("_snoozed", "_send_queue"):
            self._full_refresh()
            return
        self._set_status("Checking for new emails...")
        self.progress.start(15)

        # Find the newest receivedDateTime we already have
        newest_dt = None
        for em in self.emails:
            rd = em.get("receivedDateTime", "")
            if rd and (newest_dt is None or rd > newest_dt):
                newest_dt = rd

        existing_ids = {e.get("id") for e in self.emails}

        _after_val = self.after_entry.GetValue()
        _before_val = self.before_entry.GetValue()

        def run():
            try:
                after_dt = self._parse_date(_after_val)
                before_dt = self._parse_date(_before_val)

                # Fetch recent emails — use receivedDateTime filter if we have a reference point
                params = {"top": 50, "skip": 0, "folder": self.current_folder,
                          "before": before_dt, "after": after_dt}
                result = self.graph.get_emails(**params)
                fetched = result.get("value", [])

                # Filter to only truly new emails we don't already have
                new_emails = [e for e in fetched if e.get("id") not in existing_ids]

                if not new_emails:
                    # No new emails — just update read status of existing ones
                    self._sync_read_status(fetched, existing_ids)
                    wx.CallAfter(lambda: self.progress.stop())
                    wx.CallAfter(lambda: self._set_status(f"Up to date — {len(self.emails)} emails"))
                    return

                # Process and enrich new emails
                enriched = self.intelligence.process_emails(new_emails)
                enriched = self._apply_auto_archive(enriched)

                if not enriched:
                    wx.CallAfter(lambda: self.progress.stop())
                    wx.CallAfter(lambda: self._set_status(f"Up to date — {len(self.emails)} emails"))
                    return

                # Insert new emails into the list in correct sort order
                if self.sort_by_priority:
                    self.emails = enriched + self.emails
                    # Re-sort entire list by priority to maintain correct order
                    self.emails.sort(
                        key=lambda e: (e.get("_intel",{}).get("score",0), e.get("receivedDateTime","")),
                        reverse=True)
                else:
                    # Date sort: sort new emails by date descending, then merge
                    enriched.sort(key=lambda e: e.get("receivedDateTime", ""), reverse=True)
                    # Merge into correct positions (both lists are date-sorted descending)
                    merged = []
                    ei, oi = 0, 0
                    while ei < len(enriched) and oi < len(self.emails):
                        e_dt = enriched[ei].get("receivedDateTime", "")
                        o_dt = self.emails[oi].get("receivedDateTime", "")
                        if e_dt >= o_dt:
                            merged.append(enriched[ei])
                            ei += 1
                        else:
                            merged.append(self.emails[oi])
                            oi += 1
                    merged.extend(enriched[ei:])
                    merged.extend(self.emails[oi:])
                    self.emails = merged
                self.current_skip += len(enriched)

                # Also sync read status for emails that appeared in both fetched and existing
                self._sync_read_status(fetched, existing_ids)

                # Re-render: insert new cards at top without destroying existing ones
                wx.CallAfter(lambda: self._prepend_cards(enriched))
                wx.CallAfter(lambda: self._update_stats())
                wx.CallAfter(lambda: self.progress.stop())
                wx.CallAfter(lambda: self._set_status(
                    f"{len(enriched)} new email{'s' if len(enriched)!=1 else ''} — {len(self.emails)} total"))
                wx.CallAfter(lambda: wx.CallLater(500, self._auto_archive_past_events))

            except Exception as e:
                err_msg = str(e)
                is_network = is_network_error(err_msg)
                wx.CallAfter(lambda: self.progress.stop())
                if is_network:
                    wx.CallAfter(lambda: self._set_status("⚠ Offline — will retry automatically"))
                elif self.emails:
                    # Have existing data, don't blow it away with full refresh on transient errors
                    wx.CallAfter(lambda: self._set_status(f"⚠ Refresh failed — will retry"))
                else:
                    # No data yet, try full refresh
                    wx.CallAfter(lambda: self._set_status("Refresh failed, reloading..."))
                    wx.CallAfter(self._full_refresh)
            finally:
                self._loading_more = False

        self._loading_more = True
        threading.Thread(target=run, daemon=True).start()

    def _sync_read_status(self, fetched, existing_ids):
        """Update read status of existing emails based on freshly fetched data."""
        fetched_map = {e.get("id"): e for e in fetched if e.get("id") in existing_ids}
        for em in self.emails:
            eid = em.get("id")
            if eid in fetched_map:
                fresh = fetched_map[eid]
                old_read = em.get("isRead", True)
                new_read = fresh.get("isRead", True)
                if old_read != new_read:
                    em["isRead"] = new_read
                    # Update card styling if visible
                    wx.CallAfter(lambda e=eid: self._update_card_read(e))

    def _prepend_cards(self, new_emails):
        """Insert new email cards at the correct position without destroying existing cards."""
        self._list_scroll.Freeze()
        try:
            for em in new_emails:
                eid = em.get("id")
                insert_idx = next(
                    (i for i, e in enumerate(self.emails) if e.get("id") == eid), 0)
                self._render_card(em, insert_at=insert_idx)
            self._list_scroll.FitInside()
        finally:
            self._list_scroll.Thaw()

        if self.selected_email_id and self.selected_email_id in self._card_refs:
            self._highlight_card(self.selected_email_id)

    def _load_more(self):
        if self.graph and not self._loading_more and not self.search_query:
            if self.current_folder in ("_snoozed", "_send_queue"):
                return  # Virtual folders don't paginate
            self._load_emails()

    def _auto_load_more(self):
        """Auto-load more emails when near end of list. Capped to prevent runaway loading."""
        if self.graph and not self._loading_more and not self.search_query and len(self.emails) < 500:
            if self.current_folder in ("_snoozed", "_send_queue"):
                return  # Virtual folders don't paginate
            self._loading_more = True
            self._load_emails()

    def _load_emails(self):
        """Load a single page of emails."""
        self._set_status("Loading emails...")
        self._loading_more = True
        self.progress.start(15)

        _after_val = self.after_entry.GetValue()
        _before_val = self.before_entry.GetValue()

        def run():
            try:
                after_dt = self._parse_date(_after_val)
                before_dt = self._parse_date(_before_val)
                result = self.graph.get_emails(
                    top=self.config["emails_per_page"], skip=self.current_skip,
                    folder=self.current_folder, after=after_dt, before=before_dt,
                    search=self.search_query)
                new = result.get("value", [])
                if len(new) < self.config["emails_per_page"]:
                    self._all_loaded = True
                enriched = self.intelligence.process_emails(new)
                # Auto-archive matching senders
                enriched = self._apply_auto_archive(enriched)
                self.emails.extend(enriched)
                self.current_skip += len(new)
                wx.CallAfter(lambda: self._render_list(enriched, append=True))
                wx.CallAfter(self._update_load_more_btn)
            except Exception as e:
                err_msg = str(e)
                wx.CallAfter(lambda: self._load_err(err_msg))
            finally:
                self._loading_more = False

        threading.Thread(target=run, daemon=True).start()

    def _load_emails_batch(self, target_count):
        """Load emails in background batches, rendering each page as it arrives."""
        self._set_status(f"Loading emails...")
        self._loading_more = True
        self.progress.start(15)
        page_size = 50  # max per Graph API call
        first_batch = [True]  # use list so closure can mutate

        _after_val = self.after_entry.GetValue()
        _before_val = self.before_entry.GetValue()

        def run():
            try:
                after_dt = self._parse_date(_after_val)
                before_dt = self._parse_date(_before_val)
                loaded = 0
                hit_end = False
                while loaded < target_count:
                    result = self.graph.get_emails(
                        top=page_size, skip=loaded,
                        folder=self.current_folder, after=after_dt, before=before_dt,
                        search=self.search_query)
                    new = result.get("value", [])
                    if not new:
                        hit_end = True
                        break
                    enriched = self.intelligence.process_emails(new)
                    enriched = self._apply_auto_archive(enriched)

                    if first_batch[0]:
                        # First page: replace any placeholder / empty state and render immediately
                        first_batch[0] = False
                        self.emails = list(enriched)
                        self.current_skip = len(new)
                        wx.CallAfter(lambda batch=enriched: self._render_list(batch, append=False))
                    else:
                        # Subsequent pages: append without wiping
                        self.emails.extend(enriched)
                        self.current_skip += len(new)
                        wx.CallAfter(lambda batch=enriched: self._render_list(batch, append=True))

                    loaded += len(new)
                    count = loaded
                    wx.CallAfter(lambda c=count: self._set_status(f"Loading emails ({c})..."))

                    if len(new) < page_size:
                        hit_end = True
                        break

                self._all_loaded = hit_end
                self._folder_email_cache[self.current_folder] = list(self.emails)
                wx.CallAfter(lambda: self._set_status(f"Loaded {len(self.emails)} emails"))
                wx.CallAfter(lambda: self.progress.stop())
                wx.CallAfter(self._update_load_more_btn)
                wx.CallAfter(lambda: wx.CallLater(500, self._auto_archive_past_events))
            except Exception as e:
                err_msg = str(e)
                wx.CallAfter(lambda: self._load_err(err_msg))
            finally:
                self._loading_more = False

        threading.Thread(target=run, daemon=True).start()

    def _load_err(self, err):
        self.progress.stop()
        is_auto = getattr(self, '_is_auto_refresh', False)
        self._is_auto_refresh = False
        # Network errors — show in status bar only, don't popup
        if is_network_error(err):
            self._set_status(f"⚠ Offline — will retry automatically")
            return
        # Auto-refresh errors — always silent
        if is_auto:
            self._set_status(f"⚠ Refresh failed — will retry")
            return
        self._set_status(f"Error: {err}")
        showerror("Error", err)

    # ── List Rendering ────────────────────────────────────────
