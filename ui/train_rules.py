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


class TrainRulesMixin:
    """Train Rules Dialog"""

    def _find_next_email_id(self, current_id):
        """Find the next email below the current one in the visible list.
        Uses the actual rendered card widget order as source of truth."""
        # Get card widgets in their actual visual pack order
        try:
            children = list(self._card_refs.values())
            # Each child is the border_frame; map to email IDs via _card_refs
            ref_to_id = {id(refs[0]): eid for eid, refs in self._card_refs.items()}
            visible_ids = [ref_to_id[id(w)] for w in children if id(w) in ref_to_id]
        except Exception:
            visible_ids = []

        if not visible_ids:
            # Fallback to split-filtered list order
            visible = self._get_split_emails()
            visible_ids = [e.get("id") for e in visible]

        if current_id not in visible_ids:
            return visible_ids[0] if visible_ids else None
        idx = visible_ids.index(current_id)
        # Next item below in list; if at bottom, go to item above
        if idx + 1 < len(visible_ids):
            return visible_ids[idx + 1]
        elif idx - 1 >= 0:
            return visible_ids[idx - 1]
        return None

    def _after_action(self, select_next_id=None, removed_id=None):
        log.info("[after-action] removed=%s next=%s", (removed_id or "")[:40], (select_next_id or "")[:40])
        """Fast UI update after archive/delete — removes one card, no full re-render."""
        self._cancel_reply()

        # Remove just the one card widget instead of re-rendering everything
        if removed_id and removed_id in self._card_refs:
            entry = self._card_refs.pop(removed_id); card_frame = entry[0]
            self._panel_to_id.pop(id(card_frame), None)
            # Detach from sizer BEFORE destroy — otherwise sizer keeps a null slot
            self.list_inner_sizer.Detach(card_frame)
            card_frame.Destroy()
            wx.CallAfter(lambda: self._update_scroll_region())

        # Update counters
        self._update_stats()

        # Auto-select next email and keep focus on list for continued arrow-key navigation
        if select_next_id and any(e.get("id") == select_next_id for e in self.emails):
            self.selected_email_id = None
            self._select(select_next_id)
            self._focus_pane = "list"
        else:
            self.selected_email_id = None
            # Clear detail pane
            self.d_subject.SetLabel("Select an email to view")
            self.d_from.SetLabel(""); self.d_date.SetLabel(""); self.d_to.SetLabel(""); self.d_to.Hide()
            self.d_priority.SetLabel(""); self.d_score.SetLabel(""); self.d_signals.SetLabel("")
            self._render_email_body("<html><body></body></html>", "html")

    def _update_scroll_region(self):
        """Update canvas scroll region after card removal."""
        self._list_scroll.FitInside()

    def _show_train_rules(self):
        """Show which scoring rules fired for the selected email and let user tweak them."""
        if not self.selected_email_id or not self.intelligence:
            return
        em = next((e for e in self.emails if e.get("id") == self.selected_email_id), None)
        if not em:
            return

        # Re-score with detailed breakdown
        breakdown = self._score_email_detailed(em)
        if not breakdown:
            return

        win = wx.Dialog(self.root, title="Train Rules — Score Breakdown",
                        style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        win.SetSize(560, 640)
        win.SetBackgroundColour(_hex(C["bg"]))

        outer_sizer = wx.BoxSizer(wx.VERTICAL)

        # Subject line
        subj = (em.get("subject") or "")[:60]
        subj_lbl = wx.StaticText(win, label=f"📧 {subj}")
        subj_lbl.SetFont(_font(size=10, bold=True))
        outer_sizer.Add(subj_lbl, 0, wx.ALL, 8)

        # Score display — updated live as user edits values
        current_score = em.get("_intel", {}).get("score", 0)
        score_label = wx.StaticText(win, label=f"Score: {current_score}/100")
        score_label.SetFont(_font(size=9))
        outer_sizer.Add(score_label, 0, wx.LEFT | wx.BOTTOM, 8)

        # ── Scrollable rules area ──────────────────────────────────
        scroll = wx.ScrolledWindow(win, style=wx.VSCROLL)
        scroll.SetScrollRate(0, 20)
        scroll.SetBackgroundColour(_hex(C["bg"]))
        scroll_sizer = wx.BoxSizer(wx.VERTICAL)

        # Header row
        hdr = wx.Panel(scroll)
        hdr.SetBackgroundColour(_hex(C["border"]))
        hdr_sizer = wx.BoxSizer(wx.HORIZONTAL)
        for col_label, proportion in [("Rule", 4), ("Pts", 1), ("New", 1)]:
            t = wx.StaticText(hdr, label=col_label)
            t.SetForegroundColour(_hex(C["text"]))
            t.SetFont(_font(size=9, bold=True))
            hdr_sizer.Add(t, proportion, wx.ALL, 3)
        hdr.SetSizer(hdr_sizer)
        scroll_sizer.Add(hdr, 0, wx.EXPAND | wx.BOTTOM, 2)

        # Build rows — store the actual wx.TextCtrl in row_vars
        row_vars = []
        base_score = self.intelligence.rules.get("base_score", 30)

        # Base score row (editable)
        bf = wx.Panel(scroll)
        bf_sizer = wx.BoxSizer(wx.HORIZONTAL)
        t = wx.StaticText(bf, label="Base score")
        t.SetForegroundColour(_hex(C["text"]))
        t.SetFont(_font(size=9, bold=True))
        bf_sizer.Add(t, 4, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 3)
        t2 = wx.StaticText(bf, label=str(base_score))
        t2.SetForegroundColour(_hex(C["blue"]))
        t2.SetFont(_font(size=9, bold=True))
        bf_sizer.Add(t2, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 3)
        base_entry = wx.TextCtrl(bf, value=str(base_score), size=(50, -1))
        bf_sizer.Add(base_entry, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 2)
        bf.SetSizer(bf_sizer)
        scroll_sizer.Add(bf, 0, wx.EXPAND)

        for item in breakdown:
            bg_col = _hex("#FFFFFF") if item["points"] >= 0 else _hex("#FFF5F5")
            rf = wx.Panel(scroll)
            rf.SetBackgroundColour(bg_col)
            rf_sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Rule name
            name_lbl = wx.StaticText(rf, label=item["name"][:40])
            name_lbl.SetBackgroundColour(bg_col)
            name_lbl.SetForegroundColour(_hex(C["text"]))
            name_lbl.SetFont(_font(size=9))
            rf_sizer.Add(name_lbl, 4, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 3)
            # Current points
            pts_str = f"+{item['points']}" if item["points"] > 0 else str(item["points"])
            fg_col = _hex("#16A34A") if item["points"] > 0 else _hex("#DC2626")
            pts_lbl = wx.StaticText(rf, label=pts_str)
            pts_lbl.SetBackgroundColour(bg_col)
            pts_lbl.SetForegroundColour(fg_col)
            pts_lbl.SetFont(_font(size=9, bold=True))
            rf_sizer.Add(pts_lbl, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 3)
            # Editable entry — store the TextCtrl directly
            entry = wx.TextCtrl(rf, value=str(item["points"]), size=(50, -1))
            rf_sizer.Add(entry, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 2)
            rf.SetSizer(rf_sizer)
            scroll_sizer.Add(rf, 0, wx.EXPAND)
            row_vars.append({"ctrl": entry, "item": item})

        # Summary separator
        scroll_sizer.Add(wx.StaticLine(scroll), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 6)

        # ── Quick Actions ────────────────────────────────────────
        sender_email = em.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        sender_name = em.get("from", {}).get("emailAddress", {}).get("name", "").lower()
        sender_domain = sender_email.split("@")[-1] if "@" in sender_email else ""
        R = self.intelligence.rules

        qa_hdr = wx.StaticText(scroll, label="⚡ Quick Actions")
        qa_hdr.SetFont(_font(size=9, bold=True))
        scroll_sizer.Add(qa_hdr, 0, wx.LEFT | wx.TOP, 6)

        _qa_applied = []  # track changes for save

        # Determine current state
        vip_entries = [v.lower() for v in R.get("vip_senders", {}).get("entries", [])]
        is_sender_vip = any(v in sender_email for v in vip_entries)
        is_domain_vip = sender_domain.lower() in vip_entries if sender_domain else False
        auto_ext = [v.lower() for v in R.get("automated_senders_extended", [])]
        is_automated = any(v in sender_email for v in auto_ext)
        crit_domains = [d.lower() for d in R.get("critical_sender_domains", {}).get("domains", [])]
        is_crit_domain = any(sender_domain.endswith(d.lstrip(".")) for d in crit_domains) if sender_domain else False
        crit_subj_pats = [p.lower() for p in R.get("critical_subject_keywords", {}).get("patterns", [])]

        def _make_action_btn(text, command):
            btn = wx.Button(scroll, label=text)
            btn.Bind(wx.EVT_BUTTON, lambda e, c=command: c())
            scroll_sizer.Add(btn, 0, wx.LEFT | wx.TOP, 4)
            return btn

        # ── VIP: sender ──
        def _toggle_sender_vip():
            entries = R.setdefault("vip_senders", {}).setdefault("entries", [])
            if is_sender_vip:
                R["vip_senders"]["entries"] = [v for v in entries if v.lower() != sender_email]
                _qa_applied.append("vip_sender_remove")
            else:
                entries.append(sender_email)
                _qa_applied.append("vip_sender_add")
            sender_vip_btn.SetLabel("  ✓ Done")
            sender_vip_btn.Enable(False)

        sender_vip_text = f"  ✕ Remove {sender_email} from VIP" if is_sender_vip else f"  ⭐ Add sender to VIP  ({sender_email})"
        sender_vip_btn = _make_action_btn(sender_vip_text, _toggle_sender_vip)

        # ── VIP: domain ──
        if sender_domain and not sender_domain.startswith("gmail") and not sender_domain.startswith("yahoo") and not sender_domain.startswith("outlook"):
            def _toggle_domain_vip():
                entries = R.setdefault("vip_senders", {}).setdefault("entries", [])
                if is_domain_vip:
                    R["vip_senders"]["entries"] = [v for v in entries if v.lower() != sender_domain.lower()]
                    _qa_applied.append("vip_domain_remove")
                else:
                    entries.append(sender_domain)
                    _qa_applied.append("vip_domain_add")
                domain_vip_btn.SetLabel("  ✓ Done")
                domain_vip_btn.Enable(False)

            domain_vip_text = f"  ✕ Remove @{sender_domain} from VIP" if is_domain_vip else f"  ⭐ Add domain to VIP  (@{sender_domain})"
            domain_vip_btn = _make_action_btn(domain_vip_text, _toggle_domain_vip)

        # ── Critical domain ──
        if sender_domain and not is_crit_domain:
            def _add_crit_domain():
                domains = R.setdefault("critical_sender_domains", {}).setdefault("domains", [])
                domains.append(f".{sender_domain}")
                _qa_applied.append("crit_domain_add")
                crit_domain_btn.SetLabel("  ✓ Done")
                crit_domain_btn.Enable(False)

            crit_domain_btn = _make_action_btn(
                f"  🔴 Add @{sender_domain} to critical domains  (+{R.get('critical_sender_domains',{}).get('score',40)} pts)",
                _add_crit_domain)

        # ── Keyword section ──
        scroll_sizer.Add(wx.StaticLine(scroll), 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 6)
        kw_hdr = wx.StaticText(scroll, label="📝 Add keyword to scoring list:")
        kw_hdr.SetFont(_font(size=9, bold=True))
        scroll_sizer.Add(kw_hdr, 0, wx.LEFT | wx.BOTTOM, 4)

        # Radio buttons to select target list
        kw_target = ["important_keywords"]
        kw_radio_frame = wx.Panel(scroll)
        kw_radio_sizer = wx.BoxSizer(wx.HORIZONTAL)
        first_radio = True
        for radio_label, key in [
            ("Critical Subject", "critical_subject_keywords"),
            ("Urgent", "urgent_keywords"),
            ("Important", "important_keywords"),
            ("Low Priority", "low_priority_keywords"),
        ]:
            style = wx.RB_GROUP if first_radio else 0
            rb = wx.RadioButton(kw_radio_frame, label=radio_label, style=style)
            rb.SetValue(key == "important_keywords")
            rb.Bind(wx.EVT_RADIOBUTTON, lambda e, k=key: kw_target.__setitem__(0, k))
            kw_radio_sizer.Add(rb, 0, wx.RIGHT, 8)
            first_radio = False
        kw_radio_frame.SetSizer(kw_radio_sizer)
        scroll_sizer.Add(kw_radio_frame, 0, wx.LEFT | wx.BOTTOM, 4)

        # Input + Add button
        kw_input_frame = wx.Panel(scroll)
        kw_input_sizer = wx.BoxSizer(wx.HORIZONTAL)
        kw_entry = wx.TextCtrl(kw_input_frame, size=(200, -1))
        kw_input_sizer.Add(kw_entry, 1, wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, 6)
        kw_status_lbl = wx.StaticText(scroll, label="")
        kw_status_lbl.SetFont(_font(size=9))

        def _add_keyword(event=None):
            kw = kw_entry.GetValue().strip().lower()
            list_key = kw_target[0]
            if not kw:
                return
            section = R.get(list_key, {})
            if list_key == "critical_subject_keywords":
                keywords = section.get("patterns", [])
            else:
                keywords = section.get("keywords", [])
            if kw not in [k.lower() for k in keywords]:
                keywords.append(kw)
                _qa_applied.append(("kw", list_key, kw))
                friendly = {"critical_subject_keywords": "Critical Subject",
                            "urgent_keywords": "Urgent", "important_keywords": "Important",
                            "low_priority_keywords": "Low Priority"}.get(list_key, list_key)
                kw_status_lbl.SetLabel(f"✓ Added \"{kw}\" → {friendly}")
            else:
                kw_status_lbl.SetLabel(f"Already in list: \"{kw}\"")
            kw_entry.SetValue("")

        def _on_kw_key(event):
            if event.GetKeyCode() == wx.WXK_RETURN:
                _add_keyword()
            else:
                event.Skip()

        add_kw_btn = wx.Button(kw_input_frame, label="➕ Add")
        add_kw_btn.Bind(wx.EVT_BUTTON, _add_keyword)
        kw_input_sizer.Add(add_kw_btn, 0, wx.ALIGN_CENTER_VERTICAL)
        kw_entry.Bind(wx.EVT_KEY_DOWN, _on_kw_key)
        kw_input_frame.SetSizer(kw_input_sizer)
        scroll_sizer.Add(kw_input_frame, 0, wx.LEFT | wx.BOTTOM, 4)
        scroll_sizer.Add(kw_status_lbl, 0, wx.LEFT | wx.BOTTOM, 4)

        scroll.SetSizer(scroll_sizer)
        scroll.FitInside()
        outer_sizer.Add(scroll, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)

        # ── Live score recalculation — bind EVT_TEXT on all entries ──
        def _recalculate(event=None):
            """Live recalculate score as user edits values."""
            try:
                new_base = int(base_entry.GetValue())
            except ValueError:
                new_base = base_score
            new_score = new_base
            for rv in row_vars:
                try:
                    new_score += int(rv["ctrl"].GetValue())
                except ValueError:
                    new_score += rv["item"]["points"]
            new_score = max(0, min(100, new_score))
            thresholds = self.intelligence.rules.get("priority_thresholds", {})
            if new_score >= thresholds.get("urgent", 75):
                pri = "🔴 URGENT"
            elif new_score >= thresholds.get("important", 55):
                pri = "🟠 IMPORTANT"
            elif new_score >= thresholds.get("normal", 35):
                pri = "🔵 NORMAL"
            else:
                pri = "⚪ LOW"
            score_label.SetLabel(f"Score: {new_score}/100  →  {pri}")

        base_entry.Bind(wx.EVT_TEXT, _recalculate)
        for rv in row_vars:
            rv["ctrl"].Bind(wx.EVT_TEXT, _recalculate)

        def _save_tweaks(event=None):
            """Apply tweaked scores and quick actions back to the actual rules config."""
            R = self.intelligence.rules
            changed = 0

            # Count quick action changes
            if _qa_applied:
                changed += len(_qa_applied)

            # Save base score if changed
            try:
                new_base = int(base_entry.GetValue())
                if new_base != base_score:
                    R["base_score"] = new_base
                    changed += 1
            except ValueError:
                pass

            for rv in row_vars:
                try:
                    new_val = int(rv["ctrl"].GetValue())
                except ValueError:
                    continue
                item = rv["item"]
                if new_val == item["points"]:
                    continue
                path = item.get("rule_path")
                if not path:
                    continue
                if len(path) == 2:
                    section = R.get(path[0], {})
                    if isinstance(section, dict):
                        target = section.get(path[1])
                        if isinstance(target, dict) and "score" in target:
                            target["score"] = new_val
                        else:
                            section[path[1]] = new_val
                        changed += 1
                elif len(path) == 1:
                    R[path[0]] = new_val
                    changed += 1
            if changed:
                save_scoring_rules(R)
            win.EndModal(wx.ID_OK)
            if changed:
                self._rescore_and_rerender()

        # ── Buttons at bottom ──────────────────────────────────────
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        save_btn = wx.Button(win, label="💾 Save Changes")
        save_btn.Bind(wx.EVT_BUTTON, _save_tweaks)
        close_btn2 = wx.Button(win, label="Close")
        close_btn2.Bind(wx.EVT_BUTTON, lambda e: win.EndModal(wx.ID_CANCEL))
        btn_sizer.Add(save_btn, 0, wx.RIGHT, 8)
        btn_sizer.Add(close_btn2, 0)
        outer_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        win.SetSizer(outer_sizer)
        win.Layout()
        win.ShowModal()
        win.Destroy()

    def _score_email_detailed(self, email):
        """Re-run scoring and return a detailed breakdown of each rule that fired."""
        R = self.intelligence.rules
        breakdown = []
        text = extract_text(email).lower()
        subject = (email.get("subject") or "").lower()
        preview = (email.get("bodyPreview") or "").lower()
        latest_reply = extract_latest_reply(email)
        # Use latest reply for keyword matching to avoid false positives from quoted thread history
        combined = f"{subject} {latest_reply}"

        sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "").lower()
        to_addrs = [r.get("emailAddress", {}).get("address", "").lower()
                    for r in email.get("toRecipients", [])]
        cc_addrs = [r.get("emailAddress", {}).get("address", "").lower()
                    for r in email.get("ccRecipients", [])]
        user_email = self.intelligence.user_email

        # Static signals
        ss = R.get("static_signals", {})
        if ss.get("unread", {}).get("enabled", True) and not email.get("isRead", True):
            pts = ss["unread"].get("score", 5)
            breakdown.append({"name": "Unread", "points": pts, "rule_path": ["static_signals", "unread"]})

        if ss.get("high_importance", {}).get("enabled", True):
            if email.get("importance", "normal").lower() == "high":
                pts = ss["high_importance"].get("score", 15)
                breakdown.append({"name": "Marked high importance", "points": pts,
                                  "rule_path": ["static_signals", "high_importance"]})

        if ss.get("low_importance", {}).get("enabled", True):
            if email.get("importance", "normal").lower() == "low":
                pts = ss["low_importance"].get("score", -10)
                breakdown.append({"name": "Marked low importance", "points": pts,
                                  "rule_path": ["static_signals", "low_importance"]})

        if ss.get("flagged", {}).get("enabled", True):
            if email.get("flag", {}).get("flagStatus") == "flagged":
                pts = ss["flagged"].get("score", 12)
                breakdown.append({"name": "Flagged", "points": pts, "rule_path": ["static_signals", "flagged"]})

        if ss.get("filtered_other", {}).get("enabled", True):
            if email.get("inferenceClassification") == "other":
                pts = ss["filtered_other"].get("score", -15)
                breakdown.append({"name": "Filtered to Other", "points": pts,
                                  "rule_path": ["static_signals", "filtered_other"]})

        if ss.get("direct_to", {}).get("enabled", True):
            if user_email and user_email in to_addrs:
                pts = ss["direct_to"].get("score", 8)
                breakdown.append({"name": "Direct recipient (To:)", "points": pts,
                                  "rule_path": ["static_signals", "direct_to"]})
            elif ss.get("cc_only", {}).get("enabled", True) and user_email and user_email in cc_addrs:
                pts = ss["cc_only"].get("score", -5)
                breakdown.append({"name": "CC'd only", "points": pts,
                                  "rule_path": ["static_signals", "cc_only"]})
        # Automated senders
        auto_cfg = R.get("automated_senders", {})
        if auto_cfg.get("enabled", True):
            if any(s in sender_email for s in auto_cfg.get("patterns", [])):
                pts = auto_cfg.get("score", -20)
                breakdown.append({"name": "Automated sender", "points": pts,
                                  "rule_path": ["automated_senders", "score"]})
        # VIP senders
        vip_cfg = R.get("vip_senders", {})
        is_vip = False
        if vip_cfg.get("enabled", True):
            vip_entries = vip_cfg.get("entries", [])
            if any(vip in sender_name or vip in sender_email for vip in vip_entries):
                pts = vip_cfg.get("sender_score", 35)
                breakdown.append({"name": f"VIP sender ({sender_name[:20]})", "points": pts,
                                  "rule_path": ["vip_senders", "sender_score"]})
                is_vip = True
            if not is_vip and vip_cfg.get("recipient_score", 0):
                all_r = to_addrs + cc_addrs
                all_n = [r.get("emailAddress", {}).get("name", "").lower()
                         for r in email.get("toRecipients", []) + email.get("ccRecipients", [])]
                if any(vip in addr or vip in name for addr in all_r for vip in vip_entries for name in all_n):
                    pts = vip_cfg.get("recipient_score", 25)
                    breakdown.append({"name": "VIP on thread", "points": pts,
                                      "rule_path": ["vip_senders", "recipient_score"]})
        # Critical subjects
        crit_subj = R.get("critical_subjects", {})
        if crit_subj.get("enabled", True):
            for cs in crit_subj.get("patterns", []):
                if keyword_in_text(cs, subject):
                    pts = crit_subj.get("score", 40)
                    breakdown.append({"name": f"Critical subject: {cs}", "points": pts,
                                      "rule_path": ["critical_subjects", "score"]})
                    break

        # Critical subject keywords
        crit_kw = R.get("critical_subject_keywords", {})
        if crit_kw.get("enabled", True):
            for ck in crit_kw.get("patterns", []):
                if keyword_in_text(ck, subject):
                    pts = crit_kw.get("score", 40)
                    breakdown.append({"name": f"Critical keyword: {ck}", "points": pts,
                                      "rule_path": ["critical_subject_keywords", "score"]})
                    break

        # Critical sender domains
        crit_dom = R.get("critical_sender_domains", {})
        if crit_dom.get("enabled", True):
            for dom in crit_dom.get("domains", []):
                if sender_email.endswith(dom):
                    pts = crit_dom.get("score", 40)
                    breakdown.append({"name": f"Gov/org domain: {dom}", "points": pts,
                                      "rule_path": ["critical_sender_domains", "score"]})
                    break

        # Conditional rules
        for i, rule in enumerate(R.get("conditional_rules", [])):
            if not rule.get("enabled", True):
                continue
            match = True
            if "sender_contains" in rule and rule["sender_contains"].lower() not in sender_email:
                match = False
            if rule.get("must_be_to_recipient") and not (user_email and user_email in to_addrs):
                match = False
            if match:
                pts = rule.get("score", 0)
                breakdown.append({"name": rule.get("name", f"Conditional #{i+1}"), "points": pts,
                                  "rule_path": None})

        # Name + question detection
        nq = R.get("name_question_detection", {})
        if nq.get("enabled", True) and self.intelligence.user_names:
            addressed = any(name in latest_reply for name in self.intelligence.user_names)
            auto_ext = R.get("automated_senders_extended", [])
            is_auto = any(s in sender_email for s in auto_ext)
            if addressed and not is_auto:
                q_pats = nq.get("question_patterns", [])
                has_q = "?" in latest_reply or any(re.search(p, latest_reply) for p in q_pats)
                if has_q:
                    pts = nq.get("score", 30)
                    breakdown.append({"name": "Addressed by name + question", "points": pts,
                                      "rule_path": ["name_question_detection", "score"]})
        # Urgent keywords
        urg = R.get("urgent_keywords", {})
        if urg.get("enabled", True):
            hits = [kw for kw in urg.get("keywords", []) if keyword_in_text(kw, combined)]
            if hits:
                pts = min(urg.get("max_score", 25), len(hits) * urg.get("per_hit", 10))
                breakdown.append({"name": f"Urgent: {', '.join(h.strip() for h in hits[:3])}", "points": pts,
                                  "rule_path": ["urgent_keywords", "max_score"]})
        # Important keywords
        imp = R.get("important_keywords", {})
        if imp.get("enabled", True):
            hits = [kw for kw in imp.get("keywords", []) if keyword_in_text(kw, combined)]
            if hits:
                pts = min(imp.get("max_score", 15), len(hits) * imp.get("per_hit", 5))
                breakdown.append({"name": f"Important: {', '.join(h.strip() for h in hits[:4])}", "points": pts,
                                  "rule_path": ["important_keywords", "max_score"]})
        # General questions
        gq = R.get("general_question_detection", {})
        if gq.get("enabled", True):
            q_pats = gq.get("patterns", [])
            if any(re.search(p, combined) for p in q_pats):
                pts = gq.get("score", 8)
                breakdown.append({"name": "Contains questions/requests", "points": pts,
                                  "rule_path": ["general_question_detection", "score"]})
        # Low priority keywords
        lp = R.get("low_priority_keywords", {})
        if lp.get("enabled", True):
            hits = [kw for kw in lp.get("keywords", []) if keyword_in_text(kw, combined)]
            if hits:
                pts = -min(lp.get("max_score", 25), len(hits) * lp.get("per_hit", 8))
                breakdown.append({"name": f"Low-priority: {', '.join(h.strip() for h in hits[:3])}", "points": pts,
                                  "rule_path": ["low_priority_keywords", "max_score"]})
        # Calendar keywords
        cal = R.get("calendar_keywords", {})
        if cal.get("enabled", True):
            if any(keyword_in_text(kw, combined) for kw in cal.get("keywords", [])):
                pts = cal.get("score", 5)
                breakdown.append({"name": "Meeting/calendar related", "points": pts,
                                  "rule_path": ["calendar_keywords", "score"]})
        # Recency
        rec = R.get("recency_scores", {})
        if rec.get("enabled", True):
            received = email.get("receivedDateTime", "")
            if received:
                try:
                    recv_dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
                    age_h = (datetime.now(timezone.utc) - recv_dt).total_seconds() / 3600
                    if age_h < 1:
                        pts = rec.get("under_1h", 10)
                        breakdown.append({"name": "< 1 hour ago", "points": pts,
                                          "rule_path": ["recency_scores", "under_1h"]})
                    elif age_h < 4:
                        pts = rec.get("under_4h", 6)
                        breakdown.append({"name": "< 4 hours ago", "points": pts,
                                          "rule_path": ["recency_scores", "under_4h"]})
                    elif age_h < 24:
                        pts = rec.get("under_24h", 3)
                        breakdown.append({"name": "< 24 hours ago", "points": pts,
                                          "rule_path": ["recency_scores", "under_24h"]})
                    elif age_h > 168:
                        pts = rec.get("over_7d", -5)
                        breakdown.append({"name": "> 7 days old", "points": pts,
                                          "rule_path": ["recency_scores", "over_7d"]})
                except Exception:
                    pass

        # Attachments
        if ss.get("has_attachments", {}).get("enabled", True) and email.get("hasAttachments"):
            pts = ss["has_attachments"].get("score", 3)
            breakdown.append({"name": "Has attachments", "points": pts,
                              "rule_path": ["static_signals", "has_attachments"]})
        # Short message
        sm = ss.get("short_message", {})
        if sm.get("enabled", True):
            min_l, max_l = sm.get("min_len", 10), sm.get("max_len", 200)
            if min_l < len(text) < max_l:
                pts = sm.get("score", 4)
                breakdown.append({"name": "Short (likely needs response)", "points": pts,
                                  "rule_path": ["static_signals", "short_message"]})
        return breakdown

