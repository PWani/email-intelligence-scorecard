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
    datetime, timedelta, timezone, unescape)
try:
    from ..core.google_client import (
        GoogleAuth, GmailClient, GoogleCalendarClient,
        is_google_available, get_google_import_error, GOOGLE_CREDS_FILE)
except ImportError:
    pass


class UtilsMixin:
    """Utilities and Scheduling"""

    def _set_status(self, t): self.status.SetLabel(t)

    def _parse_date(self, t):
        t = t.strip()
        if not t or t=="YYYY-MM-DD": return None
        # Try multiple formats
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y",
                     "%Y-%m-%d", "%d/%m/%Y", "%Y.%m.%d"):
            try: return datetime.strptime(t, fmt).replace(tzinfo=timezone.utc)
            except ValueError: continue
        # Last resort: try dateutil-style loose parse
        try:
            parts = re.split(r'[/\-.]', t)
            if len(parts) == 3:
                # Assume YYYY first if first part is 4 digits
                if len(parts[0]) == 4:
                    return datetime(int(parts[0]), int(parts[1]), int(parts[2]),
                                    tzinfo=timezone.utc)
                # Otherwise assume M/D/YYYY
                return datetime(int(parts[2]), int(parts[0]), int(parts[1]),
                                tzinfo=timezone.utc)
        except Exception: pass
        return None

    def _get_tz_offset(self):
        """Get current timezone offset as timedelta."""
        hours = TIMEZONES.get(self.tz_var.GetStringSelection(), -8)
        return timedelta(hours=hours)

    def _utc_to_local(self, dt):
        """Convert a UTC-aware datetime to the user's selected timezone."""
        return dt + self._get_tz_offset()

    def _fmt_date(self, iso):
        if not iso: return ""
        try:
            dt = datetime.fromisoformat(iso.replace("Z","+00:00"))
            local = self._utc_to_local(dt)
            now_local = self._utc_to_local(datetime.now(timezone.utc))
            d = (now_local.date() - local.date()).days
            if d==0: return local.strftime("%I:%M %p")
            if d==1: return "Yesterday"
            if d<7:  return local.strftime("%a")
            if d<365:return local.strftime("%b %d")
            return local.strftime("%b %d, %Y")
        except: return iso[:10]

    def _fmt_date_full(self, iso):
        if not iso: return ""
        try:
            dt = datetime.fromisoformat(iso.replace("Z","+00:00"))
            local = self._utc_to_local(dt)
            tz_name = self.tz_var.GetStringSelection()
            return local.strftime(f"%B %d, %Y at %I:%M %p") + f" {tz_name}"
        except: return iso

    def _fmt_event_time(self, event_times):
        """Format event start/end for display."""
        if not event_times: return ""
        try:
            start = event_times.get("start", {})
            end = event_times.get("end", {})
            is_all_day = event_times.get("isAllDay", False)

            # Graph returns start/end as {"dateTime": "...", "timeZone": "UTC"}
            start_iso = start.get("dateTime", "")
            end_iso = end.get("dateTime", "")

            if not start_iso: return ""

            # Parse — Graph event times may not have Z suffix
            if not start_iso.endswith("Z") and "+" not in start_iso:
                start_iso += "Z"
            if end_iso and not end_iso.endswith("Z") and "+" not in end_iso:
                end_iso += "Z"

            start_dt = datetime.fromisoformat(start_iso.replace("Z", "+00:00"))
            start_local = self._utc_to_local(start_dt)

            if is_all_day:
                return f"📅 All day — {start_local.strftime('%A, %B %d, %Y')}"

            time_str = f"📅 {start_local.strftime('%A, %B %d at %I:%M %p')}"
            if end_iso:
                end_dt = datetime.fromisoformat(end_iso.replace("Z", "+00:00"))
                end_local = self._utc_to_local(end_dt)
                time_str += f" – {end_local.strftime('%I:%M %p')}"
            time_str += f" {self.tz_var.GetStringSelection()}"
            return time_str
        except Exception:
            return ""

    def _is_event_past(self, event_times):
        """Check if an event's start time is in the past."""
        if not event_times: return False
        try:
            start = event_times.get("start", {})
            start_iso = start.get("dateTime", "")
            if not start_iso: return False
            if not start_iso.endswith("Z") and "+" not in start_iso:
                start_iso += "Z"
            start_dt = datetime.fromisoformat(start_iso.replace("Z", "+00:00"))
            return start_dt < datetime.now(timezone.utc)
        except Exception:
            return False

    # ═══════════════════════════════════════════════════════════════
    # SNOOZE / REMIND ME
    # ═══════════════════════════════════════════════════════════════

    def _on_key_snooze(self, event=None):
        if not self._is_typing(event):
            self._dismiss_popup()
            self._show_snooze_menu()

    def _on_key_remind(self, event=None):
        if not self._is_typing(event):
            self._dismiss_popup()
            self._show_remind_menu()

    def _show_snooze_menu(self):
        """Show snooze time picker popup with numbered accelerators, reading from config."""
        if not self.selected_email_id:
            return
        options = self.config.get("snooze_options", DEFAULT_CONFIG["snooze_options"])
        menu = wx.Menu()

        for idx, opt in enumerate(options):
            num = idx + 1
            label = opt.get("label", f"Option {num}")
            _wx_menu_item(menu, f"({num})  ⏰  {label}", lambda o=opt: self._execute_snooze_option(o))

        # Position below the snooze button
        try:
            x = self._snooze_btn.GetScreenPosition().x
            y = self._snooze_btn.GetScreenPosition().y + self._snooze_btn.GetSize().height
        except Exception:
            x, y = wx.GetMousePosition()
        self._active_popup_menu = menu
        self.root.PopupMenu(menu)
        self._active_popup_menu = None
        menu.Destroy()

    def _execute_snooze_option(self, opt):
        """Execute a snooze option from config."""
        if opt.get("hours"):
            self._snooze(opt["hours"])
        elif opt.get("preset") == "tomorrow_morning":
            self._snooze_until_morning()
        elif opt.get("preset") == "next_week":
            self._snooze_next_week()
        elif opt.get("preset") == "next_monday":
            self._snooze_next_week()  # treat as 7 days

    def _show_remind_menu(self):
        """Show remind me popup with numbered accelerators, reading from config."""
        if not self.selected_email_id:
            return
        options = self.config.get("remind_options", DEFAULT_CONFIG["remind_options"])
        menu = wx.Menu()

        for idx, opt in enumerate(options):
            num = idx + 1
            label = opt.get("label", f"Option {num}")
            _wx_menu_item(menu, f"({num})  🔔  {label}", lambda o=opt: self._execute_remind_option(o))

        # Position below the remind button
        try:
            x = self._remind_btn.GetScreenPosition().x
            y = self._remind_btn.GetScreenPosition().y + self._remind_btn.GetSize().height
        except Exception:
            x, y = wx.GetMousePosition()
        self._active_popup_menu = menu
        self.root.PopupMenu(menu)
        self._active_popup_menu = None
        menu.Destroy()

    def _execute_remind_option(self, opt):
        """Execute a remind option from config."""
        days = opt.get("days", 3)
        self._remind_me(days)

    def _snooze(self, hours):
        log.info("[snooze] %d hours: eid=%s", hours, (self.selected_email_id or "")[:40])
        """Snooze selected email for N hours."""
        if not self.selected_email_id or not self.graph:
            return
        eid = self.selected_email_id
        snooze_until = datetime.now(timezone.utc) + timedelta(hours=hours)
        self._do_snooze(eid, snooze_until, f"Snoozed for {hours}h")

    def _snooze_until_morning(self):
        """Snooze until tomorrow 9 AM in user's timezone."""
        if not self.selected_email_id:
            return
        eid = self.selected_email_id
        tz_off = self._get_tz_offset()
        now_local = datetime.now(timezone.utc) + tz_off
        tomorrow_9am = now_local.replace(hour=9, minute=0, second=0, microsecond=0) + timedelta(days=1)
        snooze_until = tomorrow_9am - tz_off  # back to UTC
        self._do_snooze(eid, snooze_until, "Snoozed until tomorrow 9 AM")

    def _snooze_next_week(self):
        """Snooze for 7 days from now, 9 AM."""
        if not self.selected_email_id:
            return
        eid = self.selected_email_id
        tz_off = self._get_tz_offset()
        now_local = datetime.now(timezone.utc) + tz_off
        next_week_9am = (now_local + timedelta(days=7)).replace(
            hour=9, minute=0, second=0, microsecond=0)
        snooze_until = next_week_9am - tz_off
        self._do_snooze(eid, snooze_until, "Snoozed for 7 days")

    def _do_snooze(self, eid, snooze_until_utc, status_msg):
        """Move email to Future Action folder. Polling loop handles un-snooze."""
        next_id = self._find_next_email_id(eid)

        def run():
            try:
                # Store snooze time as a category so we can read it back
                cat = f"snooze:{snooze_until_utc.isoformat()}"
                self._api_for(eid).set_email_categories(eid, [cat])
                self._api_for(eid).snooze_email(eid)
                wx.CallAfter(lambda: self._after_snooze(eid, next_id, status_msg))
            except Exception as e:
                if is_network_error(str(e)):
                    self._offline_queue.enqueue("snooze", eid=eid,
                        until=snooze_until_utc.isoformat())
                    wx.CallAfter(lambda: self._after_snooze(eid, next_id,
                        f"{status_msg} (queued offline)"))
                else:
                    err = str(e)
                    wx.CallAfter(lambda: self._set_status(f"Snooze failed: {err}"))
        threading.Thread(target=run, daemon=True).start()

    def _after_snooze(self, eid, next_id, msg):
        self.emails = [e for e in self.emails if e.get("id") != eid]
        self._after_action(select_next_id=next_id, removed_id=eid)
        self._set_status(f"💤 {msg}")

    def _un_snooze(self, eid):
        """Move email back to inbox (called by polling loop)."""
        def run():
            try:
                self._api_for(eid).move_to_inbox(eid)
                self._api_for(eid).set_email_categories(eid, [])
                wx.CallAfter(lambda: (self._set_status("⏰ Snoozed email returned to inbox"),
                                             self._refresh()))
            except Exception:
                pass
        threading.Thread(target=run, daemon=True).start()

    def _remind_me(self, days):
        log.info("[remind] %d days: eid=%s", days, (self.selected_email_id or "")[:40])
        """Flag email for follow-up and snooze. If no reply in N days, it returns."""
        if not self.selected_email_id or not self.graph:
            return
        eid = self.selected_email_id
        remind_at = datetime.now(timezone.utc) + timedelta(days=days)
        next_id = self._find_next_email_id(eid)

        def run():
            try:
                # Flag the email (Graph-specific; Gmail uses stars instead)
                client = self._api_for(eid)
                if hasattr(client, '_patch'):
                    client._patch(client._msg_url(eid),
                                   {"flag": {"flagStatus": "flagged"}})
                # Set remind category and move to future action
                cat = f"remind:{remind_at.isoformat()}"
                self._api_for(eid).set_email_categories(eid, [cat])
                self._api_for(eid).snooze_email(eid)
                wx.CallAfter(lambda: self._after_snooze(eid, next_id,
                    f"Remind me in {days} days if no reply"))
            except Exception as e:
                err = str(e)
                wx.CallAfter(lambda: self._set_status(f"Remind failed: {err}"))
        threading.Thread(target=run, daemon=True).start()

    # ═══════════════════════════════════════════════════════════════
    # UNDO SEND (Send Queue based — all sends go through Send Queue folder)
    # ═══════════════════════════════════════════════════════════════

    def _tick_undo_countdown(self):
        """Update countdown label every second."""
        if not hasattr(self, '_undo_countdown_remaining'):
            return
        if not getattr(self, '_undo_active_draft', None):
            return
        if self._undo_countdown_remaining <= 0:
            # Countdown finished — dismiss bar, polling loop will handle the actual send
            self._dismiss_undo_bar()
            self._undo_active_draft = None
            desc = getattr(self, '_undo_countdown_desc', 'Email')
            log.info("[send] Undo window expired for '%s' — waiting for dispatch", desc)
            self._set_status(f"📤 {desc} — queued for delivery")
            return
        try:
            self._undo_countdown_label.SetLabel(f"📤 {self._undo_countdown_desc} — sending in {self._undo_countdown_remaining}s...")
            self._undo_countdown_remaining -= 1
            wx.CallLater(1000, self._tick_undo_countdown)
        except Exception:
            pass

    def _dismiss_undo_bar(self):
        if hasattr(self, '_undo_bar') and self._undo_bar:
            try:
                self._undo_bar.Hide()
                self._detail_panel.Layout()
            except Exception:
                pass
        self._undo_active_draft = None

    # ═══════════════════════════════════════════════════════════════
    # SPLIT INBOX
    # ═══════════════════════════════════════════════════════════════

    def _set_split(self, mode):
        self._split_mode = mode
        self._update_split_tabs()
        self._render_list()

    def _update_split_tabs(self):
        """Update split tab visual state."""
        for m, btn in self._split_btns.items():
            if m == self._split_mode:
                btn.SetBackgroundColour(_hex(C["accent"]))
                btn.SetForegroundColour(_hex("white"))
            else:
                btn.SetBackgroundColour(_hex(C["bg_card"]))
                btn.SetForegroundColour(_hex(C["text2"]))

    def _get_split_emails(self):
        """Filter emails based on current split mode and account filter."""
        # Apply account filter first
        if self._account_filter == "microsoft":
            pool = [e for e in self.emails if e.get("_provider") != "google"]
        elif self._account_filter == "google":
            pool = [e for e in self.emails if e.get("_provider") == "google"]
        else:
            pool = self.emails
        if self._split_mode == "all":
            return pool
        result = []
        for em in pool:
            intel = em.get("_intel", {})
            pri = intel.get("priority", "normal")
            cat = intel.get("category", "general")
            signals = intel.get("signals", [])
            is_vip = any("VIP" in s for s in signals)
            from_addr = em.get("from", {}).get("emailAddress", {}).get("address", "").lower()

            if self._split_mode == "vip" and is_vip:
                result.append(em)
            elif self._split_mode == "team":
                # Team = emails from same domain (company emails), excluding VIPs
                user_email = getattr(self.intelligence, 'user_email', '') if self.intelligence else ''
                user_domain = user_email.split("@")[-1].lower() if "@" in user_email else ""
                from_domain = from_addr.split("@")[-1].lower() if "@" in from_addr else ""
                if user_domain and from_domain == user_domain and not is_vip:
                    result.append(em)
            elif self._split_mode == "newsletters":
                # Other = not VIP, not team, not low
                user_email = getattr(self.intelligence, 'user_email', '') if self.intelligence else ''
                user_domain = user_email.split("@")[-1].lower() if "@" in user_email else ""
                from_domain = from_addr.split("@")[-1].lower() if "@" in from_addr else ""
                is_team = user_domain and from_domain == user_domain
                if not is_vip and not is_team and pri not in ("low"):
                    result.append(em)
            elif self._split_mode == "low" and pri == "low":
                result.append(em)
        return result

    # ═══════════════════════════════════════════════════════════════
    # COMMAND PALETTE (Ctrl+K)
    # ═══════════════════════════════════════════════════════════════

    def _on_command_palette(self, event=None):
        if self._is_typing(event):
            # Allow Ctrl+K even when typing
            pass
        self._show_command_palette()
        return "break"

    def _show_command_palette(self):
        """Show command palette — fuzzy search for actions and emails."""
        win = wx.Dialog(self.root, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        win.SetWindowStyle(win.GetWindowStyle() & ~wx.CAPTION)

        # Center on screen
        w, h = 550, 420
        x = self.root.GetScreenPosition().x + (self.root.GetSize().width - w) // 2
        y = self.root.GetScreenPosition().y + 80
        win.SetSize(w, h)
        win.SetPosition((x, y))

        outer_sizer = wx.BoxSizer(wx.VERTICAL)
        inner = wx.Panel(win)
        inner.SetBackgroundColour(_hex(C["bg_card"]))
        inner_sizer = wx.BoxSizer(wx.VERTICAL)

        # Search entry
        entry = wx.TextCtrl(inner, style=wx.TE_PROCESS_ENTER)
        entry.SetValue("")
        entry.SetFocus()
        inner_sizer.Add(entry, 0, wx.EXPAND | wx.ALL, 6)

        # Results list — proper wx.ScrolledWindow
        results_scroll = wx.ScrolledWindow(inner, style=wx.VSCROLL)
        results_scroll.SetScrollRate(0, 20)
        results_scroll.SetBackgroundColour(_hex(C["bg_card"]))
        results_sizer = wx.BoxSizer(wx.VERTICAL)
        results_scroll.SetSizer(results_sizer)
        inner_sizer.Add(results_scroll, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

        inner.SetSizer(inner_sizer)
        outer_sizer.Add(inner, 1, wx.EXPAND)
        win.SetSizer(outer_sizer)

        sel_idx = [0]
        items = []

        def get_commands():
            """Build list of available commands."""
            cmds = [
                {"label": "📦 Archive", "desc": "Archive selected email (A)", "action": lambda: self._on_archive_accept()},
                {"label": "↩ Reply", "desc": "Reply to selected email (R)", "action": lambda: self._reply()},
                {"label": "↩↩ Reply All", "desc": "Reply all (Shift+R)", "action": lambda: self._reply_all()},
                {"label": "→ Forward", "desc": "Forward email (F)", "action": lambda: self._forward()},
                {"label": "🗑 Delete", "desc": "Delete selected email (Del)", "action": lambda: self._delete()},
                {"label": "⏰ Snooze", "desc": "Snooze email (S)", "action": lambda: self._show_snooze_menu()},
                {"label": "🔔 Remind me in 3 days", "desc": "Remind if no reply (M)", "action": lambda: self._remind_me(3)},
                {"label": "🔔 Remind me in 7 days", "desc": "Remind if no reply", "action": lambda: self._remind_me(7)},
                {"label": "🔄 Refresh", "desc": "Check for new emails", "action": lambda: self._refresh()},
                {"label": "📊 Sort by Priority", "desc": "Toggle sort mode", "action": lambda: self._toggle_sort()},
                {"label": "📊 Sort by Date", "desc": "Toggle sort mode", "action": lambda: self._toggle_sort()},
                {"label": "⭐ Show VIP only", "desc": "Split inbox: VIP", "action": lambda: self._set_split("vip")},
                {"label": "👥 Show Team only", "desc": "Split inbox: Team", "action": lambda: self._set_split("team")},
                {"label": "📋 Show All", "desc": "Split inbox: All", "action": lambda: self._set_split("all")},
                {"label": "⚙ Scoring Rules", "desc": "Edit email scoring rules", "action": lambda: self._open_scoring_settings()},
                {"label": "ℹ About", "desc": "About this app", "action": lambda: self._show_about()},
                {"label": "👥 Accounts", "desc": "Switch account filter or manage accounts", "action": lambda: self._show_account_menu()},
                {"label": "🚪 Sign Out", "desc": "Sign out of account", "action": lambda: self._sign_out()},
            ]
            # Add folder switches
            for folder in ["inbox", "sentitems", "drafts", "archive", "deleteditems"]:
                cmds.append({
                    "label": f"📂 Go to {folder}",
                    "desc": f"Switch to {folder}",
                    "action": lambda f=folder: (self.folder_var.SetStringSelection(f), self._on_folder_change())
                })
            cmds.append({
                "label": "📂 Go to Snoozed / Reminded",
                "desc": "View snoozed and reminded emails",
                "action": lambda: (self.folder_var.SetStringSelection("snoozed / reminded"), self._on_folder_change())
            })
            cmds.append({
                "label": "📂 Go to Send Queue",
                "desc": "View queued and scheduled emails",
                "action": lambda: (self.folder_var.SetStringSelection("send queue"), self._on_folder_change())
            })
            return cmds

        def render_results(query=""):
            # Clear existing result rows
            for w in results_scroll.GetChildren():
                w.Destroy()
            results_sizer.Clear(delete_windows=False)
            items.clear()
            sel_idx[0] = 0

            q = query.lower().strip()
            cmds = get_commands()

            # Filter commands
            if q:
                matches = [c for c in cmds if q in c["label"].lower() or q in c["desc"].lower()]
            else:
                matches = cmds[:10]  # Show top commands by default

            # Also search emails if query is 3+ chars
            if len(q) >= 3:
                for em in self.emails[:50]:
                    subj = em.get("subject", "")
                    sender = em.get("from", {}).get("emailAddress", {}).get("name", "")
                    if q in subj.lower() or q in sender.lower():
                        eid = em.get("id")
                        matches.append({
                            "label": f"✉ {subj[:60]}",
                            "desc": f"From: {sender}",
                            "action": lambda i=eid: self._select(i)
                        })

            for i, item in enumerate(matches[:15]):
                row = wx.Panel(results_scroll)
                row_sizer = wx.BoxSizer(wx.VERTICAL)
                lbl = wx.StaticText(row, label=item["label"])
                desc = wx.StaticText(row, label=item["desc"])
                row_sizer.Add(lbl, 0, wx.LEFT | wx.TOP, 4)
                row_sizer.Add(desc, 0, wx.LEFT | wx.BOTTOM, 4)
                row.SetSizer(row_sizer)
                results_sizer.Add(row, 0, wx.EXPAND)
                for widget in (row, lbl, desc):
                    widget.Bind(wx.EVT_LEFT_DOWN, lambda e, idx=i: execute(idx))
                    widget.Bind(wx.EVT_ENTER_WINDOW, lambda e, idx=i: highlight(idx))
                items.append((row, lbl, desc, item["action"]))

            results_scroll.FitInside()
            results_scroll.Layout()
            highlight(0)

        def highlight(idx):
            for i, (row, lbl, desc, _) in enumerate(items):
                if i == idx:
                    row.SetBackgroundColour(_hex(C["accent"]))
                    lbl.SetBackgroundColour(_hex(C["accent"]))
                    desc.SetBackgroundColour(_hex(C["accent"]))
                else:
                    row.SetBackgroundColour(_hex(C["bg_card"]))
                    lbl.SetBackgroundColour(_hex(C["bg_card"]))
                    lbl.SetForegroundColour(_hex(C["text"]))
                    desc.SetBackgroundColour(_hex(C["bg_card"]))
                    desc.SetForegroundColour(_hex(C["muted"]))
                row.Refresh(); lbl.Refresh(); desc.Refresh()
            sel_idx[0] = idx

        def execute(idx=None):
            if idx is None:
                idx = sel_idx[0]
            if 0 <= idx < len(items):
                action = items[idx][3]
                win.EndModal(wx.ID_OK)
                action()

        def on_key(event):
            kc = event.GetKeyCode()
            if kc == wx.WXK_DOWN:
                highlight(min(sel_idx[0] + 1, len(items) - 1))
            elif kc == wx.WXK_UP:
                highlight(max(sel_idx[0] - 1, 0))
            elif kc in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
                execute()
            elif kc == wx.WXK_ESCAPE:
                win.EndModal(wx.ID_CANCEL)
            else:
                event.Skip()

        def on_text(event):
            render_results(entry.GetValue())
            event.Skip()

        entry.Bind(wx.EVT_TEXT, on_text)
        entry.Bind(wx.EVT_KEY_DOWN, on_key)
        win.Bind(wx.EVT_KEY_DOWN, lambda e: win.EndModal(wx.ID_CANCEL) if e.GetKeyCode() == wx.WXK_ESCAPE else e.Skip())

        render_results()
        win.ShowModal()
        win.Destroy()

    # ═══════════════════════════════════════════════════════════════
    # SEND LATER
    # ═══════════════════════════════════════════════════════════════

    def _show_send_later_menu(self):
        """Show send later time picker from config options."""
        menu = wx.Menu()
        _wx_menu_item(menu, "📤  Send now", self._on_send)
        menu.AppendSeparator()
        for opt in self.config.get("send_later_options", []):
            label = opt.get("label", "")
            if opt.get("hours"):
                h = opt["hours"]
                _wx_menu_item(menu, f"⏰  {label}", lambda hrs=h: self._schedule_send(hrs))
            elif opt.get("preset") == "tomorrow_morning":
                _wx_menu_item(menu, f"🌅  {label}", self._schedule_send_morning)
            elif opt.get("preset") == "next_monday":
                _wx_menu_item(menu, f"📅  {label}", self._schedule_send_monday)
        try:
            x = self._send_btn.GetScreenPosition().x
            y = self._send_btn.GetScreenPosition().y - 30 * (len(self.config.get("send_later_options", [])) + 2)
        except Exception:
            x, y = wx.GetMousePosition()
        self.root.PopupMenu(menu)

    def _schedule_send(self, hours):
        """Schedule sending the current reply/forward in N hours."""
        send_at = datetime.now(timezone.utc) + timedelta(hours=hours)
        self._do_scheduled_send(send_at, f"Scheduled to send in {hours}h")

    def _schedule_send_morning(self):
        tz_off = self._get_tz_offset()
        now_local = datetime.now(timezone.utc) + tz_off
        tomorrow_8am = now_local.replace(hour=8, minute=0, second=0, microsecond=0) + timedelta(days=1)
        send_at = tomorrow_8am - tz_off
        self._do_scheduled_send(send_at, "Scheduled for tomorrow 8 AM")

    def _schedule_send_monday(self):
        tz_off = self._get_tz_offset()
        now_local = datetime.now(timezone.utc) + tz_off
        days_ahead = (7 - now_local.weekday()) % 7
        if days_ahead == 0:
            days_ahead = 7
        monday_8am = (now_local + timedelta(days=days_ahead)).replace(
            hour=8, minute=0, second=0, microsecond=0)
        send_at = monday_8am - tz_off
        self._do_scheduled_send(send_at, "Scheduled for Monday 8 AM")

    def _do_scheduled_send(self, send_at_utc, status_msg):
        """Queue email for later sending via Send Queue folder.
        Uses createReply/createReplyAll/createForward for proper threading + attachments."""
        is_fwd = getattr(self, '_is_forward', False)

        if is_fwd:
            self._send_forward()
            fwd_data = dict(self._pending_fwd_data) if self._pending_fwd_data else None
            if not fwd_data:
                return
        else:
            self._send_reply()
            reply_data = dict(self._pending_reply_data) if self._pending_reply_data else None
            if not reply_data:
                return

        self._cancel_reply()

        # Create proper draft and move to Send Queue in background
        def _queue():
            try:
                if is_fwd:
                    draft = self._api_for(fwd_data["eid"]).create_forward_draft(
                        fwd_data["eid"], fwd_data["to_addrs"], fwd_data["comment_html"])
                else:
                    d = reply_data
                    if d["is_all"]:
                        draft = self._api_for(d["eid"]).create_reply_all_draft(
                            d["eid"], d["html"], extra_cc=d.get("extra_cc"),
                            subject=d.get("edited_subject"),
                            to_recipients=d.get("edited_to"))
                    else:
                        draft = self._api_for(d["eid"]).create_reply_draft(
                            d["eid"], d["html"], extra_cc=d.get("extra_cc"),
                            subject=d.get("edited_subject"),
                            to_recipients=d.get("edited_to"))

                draft_id = draft.get("id")
                if not draft_id:
                    log.error("[scheduled-send] Draft creation returned no ID")
                    wx.CallAfter(lambda: self._set_status("❌ Failed to create draft"))
                    return
                cats = [f"send_at:{send_at_utc.isoformat()}"]
                self._api_for(draft_id).set_email_categories(draft_id, cats)
                self._api_for(draft_id).move_to_send_queue(draft_id)
                log.info("[scheduled-send] Queued: %s | draft_id=%s | send_at=%s",
                         status_msg, draft_id[:40], send_at_utc.isoformat())
                wx.CallAfter(lambda: self._set_status(
                    f"⏰ {status_msg} — view/edit/cancel in Send Queue folder"))
            except Exception as e:
                err = str(e)
                log.error("[scheduled-send] Queue failed: %s", err[:200])
                wx.CallAfter(lambda: self._set_status(f"❌ Schedule error: {err}"))

        threading.Thread(target=_queue, daemon=True).start()

    def _recover_pending_actions(self):
        """On startup, run one immediate check then start the polling loop."""
        log.info("[dispatch] Running startup check for pending actions")
        self._check_pending_actions()
        self._start_pending_actions_timer()

    def _start_pending_actions_timer(self):
        """Poll every 15 seconds for snoozed/reminded/queued items that are due."""
        log.info("[dispatch] Starting 15s polling timer")
        def tick():
            self._check_pending_actions()
            self.root.after(15_000, tick)
        self.root.after(15_000, tick)

    def _check_pending_actions(self):
        """Scan Future Action folder and Send Queue for items whose time has passed.
        This runs every 15s and also on startup — all persistence is server-side."""
        if not self.graph:
            return

        def run():
            now = datetime.now(timezone.utc)
            recovered_snooze = 0
            recovered_send = 0

            # ── 1. Check snoozed/reminded emails ──
            try:
                snoozed = self.graph.get_snoozed_emails()
                for em in snoozed:
                    eid = em.get("id")
                    if not eid:
                        continue
                    cats = em.get("categories", [])
                    target_dt = None
                    for cat in cats:
                        if cat.startswith("snooze:") or cat.startswith("remind:"):
                            try:
                                dt_str = cat.split(":", 1)[1]
                                target_dt = datetime.fromisoformat(dt_str)
                                if target_dt.tzinfo is None:
                                    target_dt = target_dt.replace(tzinfo=timezone.utc)
                            except Exception:
                                pass

                    if target_dt and target_dt <= now:
                        try:
                            self._api_for(eid).move_to_inbox(eid)
                            self._api_for(eid).set_email_categories(eid, [])
                            recovered_snooze += 1
                        except Exception:
                            pass
            except Exception:
                pass

            # ── 2. Check Send Queue for due emails ──
            try:
                queued = self.graph.get_send_queue()
                if queued:
                    log.info("[dispatch] Polling Send Queue: %d item(s) found", len(queued))
                for draft in queued:
                    did = draft.get("id")
                    subj = draft.get("subject", "(no subject)")[:50]
                    if not did:
                        continue
                    cats = draft.get("categories", [])
                    target_dt = None
                    for cat in cats:
                        if cat.startswith("send_at:"):
                            try:
                                dt_str = cat.split(":", 1)[1]
                                target_dt = datetime.fromisoformat(dt_str)
                                if target_dt.tzinfo is None:
                                    target_dt = target_dt.replace(tzinfo=timezone.utc)
                            except Exception as parse_err:
                                log.warning("[dispatch] Bad send_at parse: cat='%s' err=%s", cat, parse_err)

                    if target_dt and target_dt <= now:
                        log.info("[dispatch] Due for send: subj='%s' send_at=%s (%.0fs ago) did=%s",
                                 subj, target_dt.isoformat(), (now - target_dt).total_seconds(), did[:40])
                        try:
                            self._api_for(did).set_email_categories(did, [])
                            self._api_for(did).send_draft(did)
                            recovered_send += 1
                            log.info("[dispatch] Sent OK: subj='%s'", subj)
                        except Exception as send_err:
                            log.error("[dispatch] Send FAILED: subj='%s' err=%s", subj, send_err)
                    elif target_dt:
                        secs_left = (target_dt - now).total_seconds()
                        log.info("[dispatch] Not yet due: subj='%s' send_at=%s (%.0fs remaining)",
                                  subj, target_dt.isoformat(), secs_left)
                    else:
                        log.warning("[dispatch] No send_at found: subj='%s' cats=%s — removing from queue", subj, cats)
                        try:
                            # Move draft out of Send Queue folder back to drafts
                            self._api_for(did).set_email_categories(did, [])
                        except Exception:
                            pass
            except Exception as sq_err:
                log.error("[dispatch] Send Queue poll error: %s", sq_err)

            # Report only when something happened
            if recovered_snooze or recovered_send:
                parts = []
                if recovered_snooze:
                    parts.append(f"{recovered_snooze} snoozed email{'s' if recovered_snooze>1 else ''} returned")
                if recovered_send:
                    parts.append(f"{recovered_send} email{'s' if recovered_send>1 else ''} sent")
                msg = "🔄 " + ", ".join(parts)
                log.info("[dispatch] %s", msg)
                wx.CallAfter(lambda: (self._set_status(msg), self._refresh()))

            # Replay offline queue
            self._replay_offline_queue()

        threading.Thread(target=run, daemon=True).start()

    def _replay_offline_queue(self):
        """Attempt to replay queued offline actions. Runs in background thread."""
        if not self.graph or self._offline_queue.is_empty():
            return

        items = self._offline_queue.peek_all()
        replayed = 0

        for item in items:
            action = item.get("action")
            eid = item.get("eid", "")
            try:
                if action == "archive":
                    self._api_for(eid).archive_email(eid)
                elif action == "delete":
                    self._api_for(eid).delete_email(eid)
                elif action == "mark_read":
                    self._api_for(eid).mark_as_read(eid)
                elif action == "accept_invite":
                    try:
                        self._api_for(eid).accept_event(eid)
                    except Exception:
                        pass  # Event may be gone
                elif action == "snooze":
                    cat = f"snooze:{item.get('until', '')}"
                    self._api_for(eid).set_email_categories(eid, [cat])
                    self._api_for(eid).snooze_email(eid)
                elif action == "reply":
                    d = item
                    if d.get("is_all"):
                        draft = self._api_for(d["eid"]).create_reply_all_draft(
                            d["eid"], d["html"], extra_cc=d.get("extra_cc"),
                            subject=d.get("edited_subject"),
                            to_recipients=d.get("edited_to"))
                    else:
                        draft = self._api_for(d["eid"]).create_reply_draft(
                            d["eid"], d["html"], extra_cc=d.get("extra_cc"),
                            subject=d.get("edited_subject"),
                            to_recipients=d.get("edited_to"))
                    draft_id = draft.get("id")
                    if draft_id:
                        self._api_for(item["eid"]).send_draft(draft_id)
                elif action == "forward":
                    draft = self._api_for(item["eid"]).create_forward_draft(
                        item["eid"], item["to_addrs"], item.get("comment_html", ""))
                    draft_id = draft.get("id")
                    if draft_id:
                        self._api_for(item["eid"]).send_draft(draft_id)
                else:
                    pass  # Unknown action — skip

                replayed += 1
            except Exception as e:
                if is_network_error(str(e)):
                    # Still offline — stop trying, keep remaining items
                    break
                else:
                    # Non-network error (404, 400, etc.) — item is stale, skip it
                    replayed += 1

        if replayed > 0:
            self._offline_queue.remove_completed(replayed)
            remaining = self._offline_queue.count
            if remaining == 0:
                wx.CallAfter(lambda: self._set_status(
                    f"🔄 Synced {replayed} offline action{'s' if replayed!=1 else ''}"))
            else:
                wx.CallAfter(lambda: self._set_status(
                    f"🔄 Synced {replayed}, {remaining} still pending"))

    def run(self):
        self._patch_root_after()
        self.root.Show()
        self._wx_app.MainLoop()
