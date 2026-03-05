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


class ActionsMixin:
    """Email Actions"""

    def _on_archive_accept(self):
        """Dispatch to archive or accept based on current email type.
        If meeting status is unknown, check first to avoid accidentally archiving invites."""
        log.info("[action] archive/accept dispatch: eid=%s", (self.selected_email_id or "")[:40])
        if self._current_action_is_accept:
            self._accept_invite()
            return
        # Double-check: if we haven't determined meeting status, do so before archiving
        if self.selected_email_id:
            em = next((e for e in self.emails if e.get("id") == self.selected_email_id), None)
            if em and "_is_meeting_request" not in em:
                self._set_status("Checking email type...")
                def check_then_act():
                    try:
                        _c = self._api_for(self.selected_email_id)
                        is_meeting = _c.is_meeting_request(self.selected_email_id)
                        em["_is_meeting_request"] = is_meeting
                        if is_meeting:
                            em["_event_times"] = _c.get_event_times(self.selected_email_id)
                            wx.CallAfter(lambda: (
                                self._apply_meeting_button(True),
                                self._accept_invite()))
                        else:
                            wx.CallAfter(self._archive)
                    except Exception:
                        wx.CallAfter(self._archive)
                threading.Thread(target=check_then_act, daemon=True).start()
                return
        self._archive()

    def _accept_invite(self):
        """Accept a meeting invitation."""
        if not self.selected_email_id: return
        eid = self.selected_email_id
        em = next((e for e in self.emails if e.get("id") == eid), None)
        subj = (em.get("subject") or "")[:50] if em else "?"
        prov = em.get("_provider", "ms") if em else "?"
        log.info("[accept] Meeting: subj='%s' provider=%s eid=%s", subj, prov, eid[:40])
        # Resolve the correct API client BEFORE removing from self.emails
        client = self._api_for(eid)
        next_id = self._find_next_email_id(eid)
        # Update UI immediately
        self.emails = [e for e in self.emails if e.get("id")!=eid]
        self._after_action(next_id, removed_id=eid)
        self._set_status("Accepting meeting...")
        # API call in background
        def run():
            try:
                client.accept_event(eid)
                try: client.archive_email(eid)
                except Exception: pass
                wx.CallAfter(lambda: self._set_status("Meeting accepted ✓"))
            except Exception as e:
                # Log full traceback to debug file
                import traceback
                log.error("[_accept_invite] %s: %s\n%s", type(e).__name__, e, traceback.format_exc())
                if is_network_error(str(e)):
                    self._offline_queue.enqueue("accept_invite", eid=eid)
                    self._offline_queue.enqueue("archive", eid=eid)
                    wx.CallAfter(lambda: self._set_status("📴 Offline — accept queued"))
                else:
                    # Accept failed (e.g. cancelled meeting) — still archive it
                    try: client.archive_email(eid)
                    except Exception: pass
                    short_err = str(e)[:80]
                    wx.CallAfter(lambda: self._set_status(f"⚠ {short_err}"))
        threading.Thread(target=run, daemon=True).start()

    def _archive(self):
        if not self.selected_email_id: return
        eid = self.selected_email_id
        # Resolve the correct API client BEFORE removing from self.emails
        # (once removed, _api_for can't find the email's provider tag)
        client = self._api_for(eid)
        em = next((e for e in self.emails if e.get("id") == eid), None)
        subj = (em.get("subject") or "")[:50] if em else ""
        prov = em.get("_provider", "ms") if em else "?"
        log.info("[archive] Manual: subj='%s' provider=%s eid=%s", subj, prov, eid[:40])
        next_id = self._find_next_email_id(eid)
        # Update UI immediately (optimistic)
        self.emails = [e for e in self.emails if e.get("id")!=eid]
        self._after_action(next_id, removed_id=eid)
        # API call in background
        def run():
            try:
                client.archive_email(eid)
                log.info("[archive] OK: subj='%s' provider=%s", subj, prov)
            except Exception as e:
                log.error("[archive] FAILED: subj='%s' err=%s", subj, e)
                if is_network_error(str(e)):
                    self._offline_queue.enqueue("archive", eid=eid)
                    wx.CallAfter(lambda: self._set_status("📴 Offline — archive queued"))
                else:
                    # Restore email to list since server rejected the archive
                    if em:
                        def restore():
                            self.emails.append(em)
                            self._render_list()
                            self._set_status(f"Archive failed: {subj}")
                        wx.CallAfter(restore)
        threading.Thread(target=run, daemon=True).start()

    def _delete(self):
        if not self.selected_email_id: return
        eid = self.selected_email_id
        em = next((e for e in self.emails if e.get("id") == eid), None)
        subj = (em.get("subject") or "")[:50] if em else "?"
        prov = em.get("_provider", "ms") if em else "?"
        log.info("[delete] Manual: subj='%s' provider=%s eid=%s", subj, prov, eid[:40])
        # Resolve the correct API client BEFORE removing from self.emails
        client = self._api_for(eid)
        next_id = self._find_next_email_id(eid)
        # Update UI immediately
        self.emails = [e for e in self.emails if e.get("id")!=eid]
        self._after_action(next_id, removed_id=eid)
        # API call in background
        def run():
            try:
                client.delete_email(eid)
            except Exception as e:
                if is_network_error(str(e)):
                    self._offline_queue.enqueue("delete", eid=eid)
                    wx.CallAfter(lambda: self._set_status("📴 Offline — delete queued"))
                else:
                    err_msg = str(e)
                    wx.CallAfter(lambda: showerror("Delete Error", err_msg))
        threading.Thread(target=run, daemon=True).start()

    def _auto_archive_sender(self):
        """Add the selected email's sender to auto-archive list, archive this email,
        and archive all other visible emails from the same sender."""
        if not self.selected_email_id:
            return
        # Find the sender
        email = next((e for e in self.emails if e.get("id") == self.selected_email_id), None)
        if not email:
            return
        sender_addr = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        sender_name = email.get("from", {}).get("emailAddress", {}).get("name", sender_addr)
        # Fallback: try _sender_addr cached during intel enrichment
        if not sender_addr:
            sender_addr = email.get("_sender_addr", "").lower()
        if not sender_addr:
            log.warning("[auto-archive] no sender address found for eid=%s", (self.selected_email_id or "")[:40])
            self._set_status("Cannot auto-archive: sender address unknown")
            return

        if not askyesno("Auto Archive Sender", f"Automatically archive all emails from:\n\n"
                f"  {sender_name}\n  ({sender_addr})\n\n"
                f"This sender will be added to your Auto Archive rules.\n"
                f"All current and future emails from them will be archived.", self.root):
            return

        # Add to scoring rules
        rules = self.intelligence.rules
        aa = rules.get("auto_archive_senders", {"enabled": True, "entries": []})
        entries = aa.get("entries", [])
        if sender_addr not in entries:
            entries.append(sender_addr)
            aa["entries"] = entries
            rules["auto_archive_senders"] = aa
            save_scoring_rules(rules)
            self.intelligence.reload_rules()

        # Archive all visible emails from this sender
        matching = [e for e in self.emails
                    if e.get("from", {}).get("emailAddress", {}).get("address", "").lower() == sender_addr]
        log.info("[auto-archive] sender=%s matched %d emails", sender_addr, len(matching))
        if not matching:
            self._set_status(f"No emails found from {sender_addr}")
            return

        # Find next email below current position that's NOT from this sender
        visible = self._get_split_emails()
        next_id = None
        found_current = False
        for e in visible:
            eid_check = e.get("id")
            if eid_check == self.selected_email_id:
                found_current = True
                continue
            if found_current and e.get("from", {}).get("emailAddress", {}).get("address", "").lower() != sender_addr:
                next_id = eid_check
                break
        # If nothing below, try above
        if not next_id:
            for e in reversed(visible):
                eid_check = e.get("id")
                if eid_check == self.selected_email_id:
                    break
                if e.get("from", {}).get("emailAddress", {}).get("address", "").lower() != sender_addr:
                    next_id = eid_check

        # Resolve API clients BEFORE removing from self.emails
        # (_api_for searches self.emails and won't find them after removal)
        archive_tasks = []
        for e in matching:
            eid = e.get("id")
            archive_tasks.append((self._api_for(eid), eid))

        # Remove from UI
        match_ids = {e.get("id") for e in matching}
        self.emails = [e for e in self.emails if e.get("id") not in match_ids]
        self._list_scroll.Freeze()
        for mid in match_ids:
            if mid in self._card_refs:
                try:
                    card = self._card_refs[mid][0]
                    self.list_inner_sizer.Detach(card)
                    self._panel_to_id.pop(id(card), None)
                    card.Destroy()
                except Exception:
                    pass
                del self._card_refs[mid]
        self._list_scroll.FitInside()
        self._list_scroll.Thaw()
        self._update_stats()

        if next_id and next_id in self._card_refs:
            self._select(next_id)
        elif self.emails:
            self._select(self.emails[0].get("id"))
        else:
            self.selected_email_id = None

        count = len(matching)
        self._set_status(f"Auto-archiving {count} email(s) from {sender_name}")

        # Archive in background using pre-resolved clients
        def run():
            ok = 0
            for client, eid in archive_tasks:
                try:
                    client.archive_email(eid)
                    ok += 1
                except Exception as exc:
                    log.error("[auto-archive-sender] FAILED: eid=%s err=%s", eid[:40] if eid else "None", exc)
            log.info("[auto-archive-sender] Added sender=%s — archived %d/%d", sender_addr, ok, count)
            wx.CallAfter(lambda: self._set_status(
                f"Auto-archived {count} email(s) from {sender_name} ✓"))
        threading.Thread(target=run, daemon=True).start()

    def _apply_auto_archive(self, emails):
        """Check emails against auto-archive sender list and archive matches.
        Returns list of emails that should be displayed (non-auto-archived)."""
        rules = self.intelligence.rules
        aa = rules.get("auto_archive_senders", {})
        if not aa.get("enabled", True) or not aa.get("entries"):
            return emails
        aa_set = {s.lower() for s in aa.get("entries", [])}
        keep = []
        to_archive = []
        for e in emails:
            sender = e.get("from", {}).get("emailAddress", {}).get("address", "").lower()
            if sender in aa_set:
                to_archive.append(e)
            else:
                keep.append(e)
        # Archive in background — resolve the correct client NOW while the
        # email dicts are still available (they may not be in self.emails yet,
        # e.g. during Google email refresh before merge).
        if to_archive:
            for e in to_archive:
                sender = e.get("from", {}).get("emailAddress", {}).get("address", "")
                subj = (e.get("subject") or "")[:50]
                prov = e.get("_provider", "microsoft")
                log.info("[auto-archive] Queuing: sender=%s subj='%s' provider=%s", sender, subj, prov)
            # Build (client, id) pairs so the background thread doesn't need
            # to call _api_for (which searches self.emails and may miss them).
            archive_tasks = []
            for e in to_archive:
                eid = e.get("id")
                if e.get("_provider") == "google" and self.google_client:
                    archive_tasks.append((self.google_client, eid, "google"))
                elif self.graph:
                    archive_tasks.append((self.graph, eid, "microsoft"))
            if archive_tasks:
                def run():
                    ok = 0
                    for client, eid, prov in archive_tasks:
                        try:
                            client.archive_email(eid)
                            ok += 1
                        except Exception as exc:
                            log.error("[auto-archive] FAILED: eid=%s provider=%s err=%s",
                                      eid[:40] if eid else "None", prov, exc)
                    log.info("[auto-archive] Completed: %d/%d archived", ok, len(archive_tasks))
                    wx.CallAfter(lambda c=len(archive_tasks):
                        self._set_status(f"Auto-archived {c} email(s)"))
                threading.Thread(target=run, daemon=True).start()
        return keep

    def _mark_read(self):
        if not self.selected_email_id: return
        log.info("[mark-read] eid=%s", self.selected_email_id[:40])
        threading.Thread(target=lambda: self._mark_read_bg(self.selected_email_id), daemon=True).start()
        self._set_status("Marked as read")

    def _update_action_buttons_for_folder(self, em=None):
        """Swap action buttons based on current folder context."""
        # Hide all special buttons first
        for btn in self._sq_btns + self._snz_btns:
            btn.Hide()

        if self.current_folder == "_send_queue":
            for btn in self._normal_action_btns:
                btn.Hide()
            self._attach_btn.Hide()
            self._sq_cancel_btn.Show()
            self._sq_sendnow_btn.Show()
        elif self.current_folder == "_snoozed":
            for btn in self._normal_action_btns:
                btn.Hide()
            self._attach_btn.Hide()
            self._snz_return_btn.Show()
            self._snz_reschedule_btn.Show()
            self._snz_delete_btn.Show()
        else:
            # Re-pack normal buttons in order
            for btn in self._normal_action_btns:
                if btn in (self._delete_btn, self._auto_archive_btn):
                    btn.Show()
                else:
                    btn.Show()

    def _cancel_queued_email(self):
        log.info("[send-queue] Cancel: eid=%s", (self.selected_email_id or "")[:40])
        """Cancel a queued send by deleting the draft from Send Queue."""
        if not self.selected_email_id or not self.graph:
            return
        eid = self.selected_email_id
        em = next((e for e in self.emails if e.get("id") == eid), None)
        subj = em.get("subject", "(no subject)") if em else "(no subject)"
        if not askyesno("Cancel Send", f"Cancel this queued email?\n\n{subj}\n\nThe draft will be permanently deleted.", self.root):
            return
        next_id = self._find_next_email_id(eid)

        def run():
            try:
                self._api_for(eid).delete_email(eid)
                wx.CallAfter(lambda: self._after_queue_action(eid, next_id, f"↩ Send cancelled: {subj[:40]}"))
            except Exception as e:
                err = str(e)
                if "404" in err or "Not Found" in err or "ErrorItemNotFound" in err:
                    wx.CallAfter(lambda: self._after_queue_action(eid, next_id, f"⚠ Already sent: {subj[:40]}"))
                else:
                    wx.CallAfter(lambda: self._set_status(f"Cancel failed — check connection"))
        threading.Thread(target=run, daemon=True).start()

    def _send_queued_now(self):
        """Send a queued email immediately instead of waiting."""
        if not self.selected_email_id or not self.graph:
            return
        eid = self.selected_email_id
        em = next((e for e in self.emails if e.get("id") == eid), None)
        subj = em.get("subject", "(no subject)") if em else "(no subject)"
        next_id = self._find_next_email_id(eid)

        def run():
            try:
                log.info("[send-now] Sending immediately: subj='%s' eid=%s", subj[:50], eid[:40])
                self._api_for(eid).set_email_categories(eid, [])
                self._api_for(eid).send_draft(eid)
                log.info("[send-now] Sent OK: subj='%s'", subj[:50])
                wx.CallAfter(lambda: self._after_queue_action(eid, next_id, f"📤 Sent now: {subj[:40]}"))
            except Exception as e:
                err = str(e)
                if "404" in err or "Not Found" in err or "ErrorItemNotFound" in err:
                    log.info("[send-now] Already sent/gone: subj='%s'", subj[:50])
                    wx.CallAfter(lambda: self._after_queue_action(eid, next_id, f"✓ Already sent: {subj[:40]}"))
                else:
                    log.error("[send-now] Failed: subj='%s' err=%s", subj[:50], err[:200])
                    wx.CallAfter(lambda: self._set_status(f"Send failed: {subj[:30]}… — check connection"))
        threading.Thread(target=run, daemon=True).start()

    def _return_to_inbox(self):
        log.info("[snooze] Return to inbox: eid=%s", (self.selected_email_id or "")[:40])
        """Move snoozed/reminded email back to inbox immediately."""
        if not self.selected_email_id or not self.graph:
            return
        eid = self.selected_email_id
        next_id = self._find_next_email_id(eid)

        def run():
            try:
                self._api_for(eid).move_to_inbox(eid)
                self._api_for(eid).set_email_categories(eid, [])
                wx.CallAfter(lambda: self._after_queue_action(eid, next_id, "📥 Returned to inbox"))
            except Exception as e:
                err = str(e)
                wx.CallAfter(lambda: self._set_status(f"Failed: {err}"))
        threading.Thread(target=run, daemon=True).start()

    def _show_reschedule_menu(self):
        """Show reschedule options for a snoozed/reminded email."""
        if not self.selected_email_id:
            return
        menu = wx.Menu()
        # Build from snooze config options
        for opt in self.config.get("snooze_options", []):
            label = opt.get("label", "")
            if opt.get("hours"):
                h = opt["hours"]
                _wx_menu_item(menu, f"⏰ {label}", lambda hrs=h: self._reschedule_snooze_hours(hrs))
            elif opt.get("preset") == "tomorrow_morning":
                _wx_menu_item(menu, f"🌅 {label}", self._reschedule_tomorrow)
            elif opt.get("preset") == "next_week":
                _wx_menu_item(menu, f"📅 {label}", self._reschedule_next_week)
        try:
            x = self._snz_reschedule_btn.GetScreenPosition().x
            y = self._snz_reschedule_btn.GetScreenPosition().y + self._snz_reschedule_btn.GetSize().height
        except Exception:
            x, y = wx.GetMousePosition()
        self.root.PopupMenu(menu)

    def _reschedule_snooze_hours(self, hours):
        if not self.selected_email_id or not self.graph:
            return
        eid = self.selected_email_id
        new_time = datetime.now(timezone.utc) + timedelta(hours=hours)
        self._do_reschedule(eid, new_time, f"Rescheduled for {hours}h")

    def _reschedule_tomorrow(self):
        if not self.selected_email_id:
            return
        eid = self.selected_email_id
        tz_off = self._get_tz_offset()
        now_local = datetime.now(timezone.utc) + tz_off
        tomorrow_9am = now_local.replace(hour=9, minute=0, second=0, microsecond=0) + timedelta(days=1)
        new_time = tomorrow_9am - tz_off
        self._do_reschedule(eid, new_time, "Rescheduled to tomorrow 9 AM")

    def _reschedule_next_week(self):
        if not self.selected_email_id:
            return
        eid = self.selected_email_id
        tz_off = self._get_tz_offset()
        now_local = datetime.now(timezone.utc) + tz_off
        next_week = (now_local + timedelta(days=7)).replace(hour=9, minute=0, second=0, microsecond=0)
        new_time = next_week - tz_off
        self._do_reschedule(eid, new_time, "Rescheduled for next week")

    def _do_reschedule(self, eid, new_time_utc, status_msg):
        """Update the snooze/remind time on a snoozed email."""
        next_id = self._find_next_email_id(eid)
        em = next((e for e in self.emails if e.get("id") == eid), None)

        def run():
            try:
                # Determine if this was snooze or remind based on existing categories
                cats = em.get("categories", []) if em else []
                prefix = "snooze"
                for cat in cats:
                    if cat.startswith("remind:"):
                        prefix = "remind"
                        break
                new_cat = f"{prefix}:{new_time_utc.isoformat()}"
                self._api_for(eid).set_email_categories(eid, [new_cat])
                wx.CallAfter(lambda: (
                    self._set_status(f"⏰ {status_msg}"),
                    self._full_refresh()))
            except Exception as e:
                err = str(e)
                wx.CallAfter(lambda: self._set_status(f"Reschedule failed: {err}"))
        threading.Thread(target=run, daemon=True).start()

    def _delete_snoozed(self):
        log.info("[snooze] Delete snoozed: eid=%s", (self.selected_email_id or "")[:40])
        """Permanently delete a snoozed/reminded email."""
        if not self.selected_email_id or not self.graph:
            return
        eid = self.selected_email_id
        em = next((e for e in self.emails if e.get("id") == eid), None)
        subj = em.get("subject", "(no subject)") if em else "(no subject)"
        if not askyesno("Delete", f"Permanently delete this email?\n\n{subj}", self.root):
            return
        next_id = self._find_next_email_id(eid)

        def run():
            try:
                self._api_for(eid).delete_email(eid)
                wx.CallAfter(lambda: self._after_queue_action(eid, next_id, f"🗑 Deleted: {subj[:40]}"))
            except Exception as e:
                err = str(e)
                wx.CallAfter(lambda: self._set_status(f"Delete failed: {err}"))
        threading.Thread(target=run, daemon=True).start()

    def _after_queue_action(self, eid, next_id, status_msg):
        """Update UI after a send queue action (cancel/send now)."""
        self._dismiss_undo_bar()
        self._undo_countdown_remaining = 0
        self.emails = [e for e in self.emails if e.get("id") != eid]
        if eid in self._card_refs:
            try:
                card = self._card_refs[eid][0]
                self.list_inner_sizer.Detach(card)
                self._panel_to_id.pop(id(card), None)
                card.Destroy()
            except Exception:
                pass
            del self._card_refs[eid]
        self._list_scroll.FitInside()
        self._update_stats()
        if next_id and next_id in self._card_refs:
            self._select(next_id)
        else:
            visible = self._get_split_emails()
            if visible:
                self._select(visible[0].get("id"))
            else:
                self.selected_email_id = None
        self._set_status(status_msg)

    # ── Attachments ───────────────────────────────────────────

