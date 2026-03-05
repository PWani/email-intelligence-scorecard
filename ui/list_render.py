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


class ListRenderMixin:
    """List Rendering"""

    def _render_list(self, new_emails=None, append=False):
        self.progress.stop()

        if not append:
            # Full re-render — sort, filter, then do a diff against what's already on screen.
            if self.sort_by_priority:
                self.emails.sort(
                    key=lambda e: (e.get("_intel", {}).get("score", 0),
                                   e.get("receivedDateTime", "")),
                    reverse=True)
            else:
                self.emails.sort(key=lambda e: e.get("receivedDateTime", ""), reverse=True)

            items = self._get_split_emails()
            wanted_ids = [e.get("id") for e in items]
            current_ids = list(self._card_refs.keys())

            # Fast path: same set and same order → just update in-place
            if wanted_ids == current_ids:
                self._update_cards_in_place(items)
            else:
                self._rebuild_list(items)
        else:
            # Append path — just add new cards to the bottom
            items = new_emails or []
            self._list_scroll.Freeze()
            try:
                for em in items:
                    self._render_card(em)
                self._list_scroll.FitInside()
            finally:
                self._list_scroll.Thaw()

        # Re-highlight selected card
        if self.selected_email_id and self.selected_email_id in self._card_refs:
            self._highlight_card(self.selected_email_id)

        self._update_stats()
        split_count = len(self._get_split_emails())
        total = len(self.emails)
        if self._split_mode != "all":
            self._set_status(f"Showing {split_count} of {total} emails")
        else:
            self._set_status(f"Loaded {total} emails")

        if not self.selected_email_id:
            first = self._get_split_emails()
            if first:
                first_id = first[0].get("id")
                if first_id:
                    wx.CallLater(100, lambda: self._select(first_id))

    def _rebuild_list(self, items):
        """Full wipe and rebuild — used only when card set or order actually changed."""
        self._list_scroll.Freeze()
        try:
            self.list_inner_sizer.Clear(delete_windows=True)
            self._card_refs = {}
            self._panel_to_id = {}
            for em in items:
                self._render_card(em)
            self._list_scroll.FitInside()
        finally:
            self._list_scroll.Thaw()

    def _update_cards_in_place(self, items):
        """Update existing card labels without destroying/recreating widgets.

        Called when the ordered set of visible emails hasn't changed — avoids
        all widget destruction and re-creation, eliminating flicker entirely.
        Only the text labels and colours that may have changed (score, read
        state, flag) are touched.
        """
        for em in items:
            eid = em.get("id")
            if eid not in self._card_refs:
                continue
            entry = self._card_refs[eid]
            card, orig_bg = entry[0], entry[1]
            accent = entry[2] if len(entry) > 2 else None
            content = entry[3] if len(entry) > 3 else None
            intel = em.get("_intel", {})
            pri = intel.get("priority", "normal")
            score = intel.get("score", 50)
            bg_hex = C.get(f"{pri}_bg", C["bg_card"])
            wx_bg = _hex(bg_hex)

            unread = not em.get("isRead", True)
            flagged = em.get("flag", {}).get("flagStatus") == "flagged"

            # Update background colour if priority changed
            if wx_bg != orig_bg:
                card.SetBackgroundColour(wx_bg)
                if accent: accent.SetBackgroundColour(wx_bg)
                if content: content.SetBackgroundColour(wx_bg)
                self._card_refs[eid] = (card, wx_bg, accent, content)

            # Walk content panel children by position:
            # content children: [0]=icon_lbl, [1]=pri_lbl, [2]=score_lbl, [3]=subj_lbl, [4]=sender_lbl, [5]=date_lbl
            kids = list(content.GetChildren()) if content else list(card.GetChildren())

            # [0]=icon_lbl, [1]=score_lbl, [2]=subj_lbl, [3]=sender_lbl, [4]=date_lbl
            if len(kids) >= 2:
                score_fg = (C["red"] if score >= 75 else C["orange"] if score >= 55
                            else C["blue"] if score >= 35 else C["muted"])
                kids[1].SetLabel(f"{score:3d}")
                kids[1].SetForegroundColour(_hex(score_fg))
                kids[1].SetBackgroundColour(wx_bg)

            if len(kids) >= 3:
                subj = em.get("subject") or "(no subject)"
                subj_display = ("🚩 " if flagged else "") + subj
                max_subj = max(int(self._list_width / 7), 30)
                if len(subj_display) > max_subj:
                    subj_display = subj_display[:max_subj] + "…"
                kids[2].SetLabel(subj_display)
                kids[2].SetFont(_font(FONT_BOLD if unread else FONT, 8))
                kids[2].SetForegroundColour(_hex(C["text"] if unread else C["text2"]))
                kids[2].SetBackgroundColour(wx_bg)

            for child in kids:
                child.SetBackgroundColour(wx_bg)

    def _update_stats(self, email_list=None):
        """Update the email counter bar and status bar from current emails."""
        pool = email_list if email_list is not None else self.emails
        u = sum(1 for e in pool if e.get("_intel",{}).get("priority")=="urgent")
        i = sum(1 for e in pool if e.get("_intel",{}).get("priority")=="important")
        n = sum(1 for e in pool if not e.get("isRead",True))
        s = f"{len(pool)} emails"
        if n: s += f"  •  {n} unread"
        if u: s += f"  •  🔴 {u} urgent"
        if i: s += f"  •  🟠 {i} important"
        self.stats_label.SetLabel(s)
        # Keep bottom status bar in sync, show offline queue if pending
        status = f"{len(self.emails)} emails loaded"
        pending = self._offline_queue.count
        if pending:
            status += f"  •  📴 {pending} action{'s' if pending!=1 else ''} queued offline"
        self._set_status(status)

    def _rescore_and_rerender(self):
        """Re-score all loaded emails with current rules and update cards in place."""
        if not self.intelligence or not self.emails:
            return
        self.intelligence.rules = load_scoring_rules()
        count = 0
        for em in self.emails:
            new_intel = self.intelligence.score_email(em)
            em["_intel"] = new_intel
            count += 1

        # Use in-place update — scores changed but IDs and order likely haven't
        items = self._get_split_emails()
        wanted_ids = [e.get("id") for e in items]
        current_ids = list(self._card_refs.keys())
        if wanted_ids == current_ids:
            self._update_cards_in_place(items)
        else:
            self._rebuild_list(items)

        if self.selected_email_id and self.selected_email_id in self._card_refs:
            self._highlight_card(self.selected_email_id)

        if self.selected_email_id:
            sel = next((e for e in self.emails if e.get("id") == self.selected_email_id), None)
            if sel and sel.get("_intel"):
                intel = sel["_intel"]
                self.d_priority.SetLabel(intel['priority'].upper())
                self.d_score.SetLabel(f"Score: {intel['score']}/100")
                self.d_signals.SetLabel("  •  ".join(intel.get("signals", [])[:5]))
        self._set_status(f"✓ Rescored {count} emails with updated rules")

    def _update_load_more_btn(self):
        """Show or hide the Load More button based on whether all emails are loaded."""
        if self._all_loaded or self.search_query:
            self._load_more_btn.Hide()
        else:
            self._load_more_btn.Show()
        self._load_more_btn.GetParent().Layout()

    def _render_card(self, email, insert_at=None):
        intel = email.get("_intel", {})
        pri = intel.get("priority", "normal")
        score = intel.get("score", 50)
        summary = intel.get("summary", "")
        cat = intel.get("category", "general")

        bg_hex = C.get(f"{pri}_bg", C["bg_card"])
        wx_bg = _hex(bg_hex)

        # Card panel
        card = wx.Panel(self._list_scroll)
        card.SetBackgroundColour(wx_bg)
        card.SetMinSize((-1, 90))
        card.SetCursor(wx.Cursor(wx.CURSOR_HAND))

        # Outer horizontal sizer: accent strip | content
        outer_h = wx.BoxSizer(wx.HORIZONTAL)
        accent_strip = wx.Panel(card, size=(4, -1))
        accent_strip.SetBackgroundColour(wx_bg)  # invisible until selected
        outer_h.Add(accent_strip, 0, wx.EXPAND)
        card.SetSizer(outer_h)

        # Inner content panel
        content = wx.Panel(card)
        content.SetBackgroundColour(wx_bg)
        content.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        outer_h.Add(content, 1, wx.EXPAND)

        card_sizer = wx.BoxSizer(wx.VERTICAL)

        # Row 1: priority icon + category + subject + score
        r1 = wx.BoxSizer(wx.HORIZONTAL)

        icon_lbl = wx.StaticText(content, label=CAT_ICON.get(cat, ''))
        icon_lbl.SetFont(_font(FONT, 9))
        icon_lbl.SetBackgroundColour(wx_bg)
        r1.Add(icon_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)

        subj = email.get("subject") or "(no subject)"
        unread = not email.get("isRead", True)
        flagged = email.get("flag", {}).get("flagStatus") == "flagged"
        subj_display = ("🚩 " if flagged else "") + subj
        max_subj = max(int(self._list_width / 7), 30)
        if len(subj_display) > max_subj:
            subj_display = subj_display[:max_subj] + "…"

        score_fg = C["red"] if score>=75 else C["orange"] if score>=55 else C["blue"] if score>=35 else C["muted"]
        score_lbl = wx.StaticText(content, label=f"{score:3d}")
        score_lbl.SetFont(_font(FONT_BOLD, 9))
        score_lbl.SetForegroundColour(_hex(score_fg))
        score_lbl.SetBackgroundColour(wx_bg)
        r1.Add(score_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)

        subj_lbl = wx.StaticText(content, label=subj_display,
                                  style=wx.ST_ELLIPSIZE_END | wx.ST_NO_AUTORESIZE)
        subj_lbl.SetFont(_font(FONT_BOLD if unread else FONT, 9))
        subj_lbl.SetForegroundColour(_hex(C["text"] if unread else C["text2"]))
        subj_lbl.SetBackgroundColour(wx_bg)
        r1.Add(subj_lbl, 1, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 4)
        card_sizer.Add(r1, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 8)

        # Row 2: sender + date
        r2 = wx.BoxSizer(wx.HORIZONTAL)
        sender = email.get("from", {}).get("emailAddress", {}).get("name", "Unknown")
        provider_prefix = ""
        if self.google_client and email.get("_provider") == "google":
            provider_prefix = "G "
        sender_lbl = wx.StaticText(content, label=provider_prefix + sender,
                                   style=wx.ST_ELLIPSIZE_END | wx.ST_NO_AUTORESIZE)
        sender_lbl.SetFont(_font(FONT, 8))
        sender_lbl.SetForegroundColour(_hex(C["muted"]))
        sender_lbl.SetBackgroundColour(wx_bg)
        sender_lbl.SetMinSize((0, -1))
        r2.Add(sender_lbl, 1, wx.ALIGN_CENTER_VERTICAL)

        date_lbl = wx.StaticText(content,
                    label=self._fmt_date(email.get("receivedDateTime", "")))
        date_lbl.SetFont(_font(FONT, 8))
        date_lbl.SetForegroundColour(_hex(C["muted"]))
        date_lbl.SetBackgroundColour(wx_bg)
        r2.Add(date_lbl, 0, wx.ALIGN_CENTER_VERTICAL)

        # Snooze/schedule info
        snooze_info = email.get("_snooze_info", "")
        if snooze_info:
            si_lbl = wx.StaticText(content, label=snooze_info)
            si_lbl.SetFont(_font(FONT, 8))
            si_lbl.SetForegroundColour(_hex(C["purple"]))
            si_lbl.SetBackgroundColour(wx_bg)
            r2.Add(si_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 4)

        card_sizer.Add(r2, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 4)

        # Row 3: summary
        if summary:
            max_sum = max(int(self._list_width / 6), 30)
            sum_text = summary[:max_sum] + ("…" if len(summary) > max_sum else "")
            sum_lbl = wx.StaticText(content, label=sum_text)
            sum_lbl.SetFont(_font(FONT, 8))
            sum_lbl.SetForegroundColour(_hex(C["muted"]))
            sum_lbl.SetBackgroundColour(wx_bg)
            card_sizer.Add(sum_lbl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

        content.SetSizer(card_sizer)

        # Add to scroll panel sizer (caller is responsible for FitInside after the loop)
        if insert_at is not None and 0 <= insert_at < self.list_inner_sizer.GetItemCount():
            self.list_inner_sizer.Insert(insert_at, card, 0,
                                         wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)
        else:
            self.list_inner_sizer.Add(card, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

        # Store ref: (panel, original_bg_colour)
        eid = email.get("id")
        self._card_refs[eid] = (card, wx_bg, accent_strip, content)  # (card, orig_bg, accent, content)
        self._panel_to_id[id(card)] = eid

        # Click handler — bind recursively to all children
        def _on_click(e, i=eid):
            self._set_focus("list")
            self._select(i, auto_scroll=False)

        def _bind_recursive(w):
            w.Bind(wx.EVT_LEFT_DOWN, _on_click)
            for ch in w.GetChildren():
                _bind_recursive(ch)
        _bind_recursive(card)


    # ── Keyboard Navigation ───────────────────────────────────

