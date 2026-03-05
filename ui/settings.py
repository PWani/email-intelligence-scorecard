from ._wx_common import (
    wx, FONT, FONT_BOLD,
    DEFAULT_CONFIG, load_config, save_config,
    load_scoring_rules, save_scoring_rules,
    DEFAULT_SCORING_RULES,
    askyesno, showerror, showinfo, askstring,
    log,
)
import json
try:
    from ..core.google_client import (
        GoogleAuth, GmailClient, GoogleCalendarClient,
        is_google_available, get_google_import_error, GOOGLE_CREDS_FILE,
    )
    _HAS_GOOGLE_MODULE = True
except ImportError:
    _HAS_GOOGLE_MODULE = False


class SettingsMixin:
    """Settings Dialog"""

    def _open_scoring_settings(self):
        """Open the scoring rules editor dialog."""
        rules = load_scoring_rules()
        win = wx.Dialog(self.root, title="Scoring Rules",
                        size=(800, 650), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        win.SetBackgroundColour(wx.Colour(243, 244, 246))

        outer = wx.BoxSizer(wx.VERTICAL)

        # Title labels
        lbl1 = wx.StaticText(win, label="⚙ Email Scoring Rules")
        lbl1.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        outer.Add(lbl1, 0, wx.ALL | wx.ALIGN_CENTER, 8)
        lbl2 = wx.StaticText(win, label="Changes take effect on next refresh")
        lbl2.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        outer.Add(lbl2, 0, wx.BOTTOM | wx.ALIGN_CENTER, 8)

        # Notebook
        nb = wx.Notebook(win)
        outer.Add(nb, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 12)

        # ── State storage ──
        # We use plain dicts/lists instead of tk.StringVar/BooleanVar
        # Each "var" is just a dict with {"widget": wx_ctrl, "type": "bool"|"str"}
        vars_store = {}
        text_widgets = {}  # key -> wx.TextCtrl (multiline)

        def make_scroll_tab(label):
            """Create a scrollable panel tab, return the inner panel."""
            outer_panel = wx.Panel(nb)
            outer_panel.SetBackgroundColour(wx.Colour(255, 255, 255))
            nb.AddPage(outer_panel, label)
            scroll = wx.ScrolledWindow(outer_panel, style=wx.VSCROLL)
            scroll.SetScrollRate(0, 20)
            scroll.SetBackgroundColour(wx.Colour(255, 255, 255))
            sizer = wx.BoxSizer(wx.VERTICAL)
            outer_panel.SetSizer(sizer)
            sizer.Add(scroll, 1, wx.EXPAND)
            inner = wx.Panel(scroll)
            inner.SetBackgroundColour(wx.Colour(255, 255, 255))
            inner_sizer = wx.GridBagSizer(2, 4)
            inner.SetSizer(inner_sizer)
            scroll_sizer = wx.BoxSizer(wx.VERTICAL)
            scroll_sizer.Add(inner, 1, wx.EXPAND | wx.ALL, 4)
            scroll.SetSizer(scroll_sizer)
            return inner, inner_sizer

        def section(parent, gs, text, row):
            lbl = wx.StaticText(parent, label=text)
            lbl.SetFont(wx.Font(11, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            lbl.SetForegroundColour(wx.Colour(31, 41, 55))
            gs.Add(lbl, pos=(row, 0), span=(1, 4), flag=wx.EXPAND | wx.LEFT | wx.TOP, border=12)
            return row + 1

        def score_row(parent, gs, key, label, cfg, row, score_key="score"):
            """Enabled checkbox + label + score entry."""
            enabled_cb = wx.CheckBox(parent)
            enabled_cb.SetValue(cfg.get("enabled", True))
            gs.Add(enabled_cb, pos=(row, 0), flag=wx.LEFT | wx.ALIGN_CENTER_VERTICAL, border=12)

            lbl = wx.StaticText(parent, label=label)
            lbl.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
            gs.Add(lbl, pos=(row, 1), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

            score_lbl = wx.StaticText(parent, label="Score:")
            score_lbl.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
            gs.Add(score_lbl, pos=(row, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.LEFT, border=8)

            score_entry = wx.TextCtrl(parent, value=str(cfg.get(score_key, 0)), size=(50, -1))
            gs.Add(score_entry, pos=(row, 3), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)

            vars_store[key] = {"enabled": enabled_cb, "score": score_entry, "score_key": score_key}
            return row + 1

        def keyword_editor(parent, gs, key, label, items, row):
            """Multi-line text editor for a keyword list."""
            lbl = wx.StaticText(parent, label=label)
            lbl.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            gs.Add(lbl, pos=(row, 0), span=(1, 4), flag=wx.EXPAND | wx.LEFT | wx.TOP, border=12)
            row += 1

            hint = wx.StaticText(parent, label="One per line:")
            hint.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
            gs.Add(hint, pos=(row, 0), span=(1, 2), flag=wx.LEFT, border=12)
            row += 1

            h = max(3, min(8, len(items))) * 20 + 10
            txt = wx.TextCtrl(parent, style=wx.TE_MULTILINE | wx.TE_DONTWRAP,
                              size=(500, h))
            txt.SetValue("\n".join(items))
            gs.Add(txt, pos=(row, 0), span=(1, 4), flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=12)
            text_widgets[key] = txt
            return row + 1

        # ── Tab 1: VIP & Critical ──
        t1, gs1 = make_scroll_tab("  VIP & Critical  ")
        r = 0
        r = section(t1, gs1, "VIP Senders", r)
        vip = rules.get("vip_senders", {})
        r = score_row(t1, gs1, "vip_sender", "VIP sender score", vip, r, "sender_score")
        r = score_row(t1, gs1, "vip_recipient", "VIP on thread score",
                      {"enabled": vip.get("enabled", True), "score": vip.get("recipient_score", 25)}, r)

        vip_label_lbl = wx.StaticText(t1, label="VIP Label:")
        gs1.Add(vip_label_lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL)
        vip_label_entry = wx.TextCtrl(t1, value=vip.get("label", "VIP"), size=(120, -1))
        gs1.Add(vip_label_entry, pos=(r, 2), span=(1, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)
        r += 1

        r = keyword_editor(t1, gs1, "vip_entries", "VIP sender patterns (name or email fragments):",
                           vip.get("entries", []), r)
        r = section(t1, gs1, "Critical Subject Patterns", r)
        r = score_row(t1, gs1, "crit_subj", "Critical subject score", rules.get("critical_subjects", {}), r)
        r = keyword_editor(t1, gs1, "crit_subj_patterns", "Patterns (matched in subject):",
                           rules.get("critical_subjects", {}).get("patterns", []), r)
        r = section(t1, gs1, "Critical Subject Keywords", r)
        r = score_row(t1, gs1, "crit_kw", "Critical keyword score", rules.get("critical_subject_keywords", {}), r)
        r = keyword_editor(t1, gs1, "crit_kw_patterns", "Keywords (matched in subject):",
                           rules.get("critical_subject_keywords", {}).get("patterns", []), r)
        r = section(t1, gs1, "Critical Sender Domains", r)
        r = score_row(t1, gs1, "crit_dom", "Critical domain score", rules.get("critical_sender_domains", {}), r)
        r = keyword_editor(t1, gs1, "crit_dom_list", "Domains (e.g. .gov):",
                           rules.get("critical_sender_domains", {}).get("domains", []), r)
        r = section(t1, gs1, "Conditional Rules", r)
        cond_rules = rules.get("conditional_rules", [])
        cond_text = json.dumps(cond_rules, indent=2)
        cond_lbl = wx.StaticText(t1, label="JSON format (advanced):")
        gs1.Add(cond_lbl, pos=(r, 0), span=(1, 4), flag=wx.LEFT, border=12)
        r += 1
        cond_txt = wx.TextCtrl(t1, style=wx.TE_MULTILINE | wx.TE_DONTWRAP,
                               size=(500, max(80, min(200, cond_text.count("\n") * 18))))
        cond_txt.SetValue(cond_text)
        gs1.Add(cond_txt, pos=(r, 0), span=(1, 4), flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=12)
        text_widgets["conditional_rules"] = cond_txt
        r += 1

        # ── Tab 2: Keywords ──
        t2, gs2 = make_scroll_tab("  Keywords  ")
        r = 0
        r = section(t2, gs2, "Urgent Keywords", r)
        urg = rules.get("urgent_keywords", {})
        r = score_row(t2, gs2, "urg", "Max score", urg, r, "max_score")
        urg_per_lbl = wx.StaticText(t2, label="Per hit:")
        gs2.Add(urg_per_lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL)
        urg_per_entry = wx.TextCtrl(t2, value=str(urg.get("per_hit", 10)), size=(50, -1))
        gs2.Add(urg_per_entry, pos=(r, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)
        r += 1
        r = keyword_editor(t2, gs2, "urg_kw", "Keywords:", urg.get("keywords", []), r)

        r = section(t2, gs2, "Important Keywords", r)
        imp_r = rules.get("important_keywords", {})
        r = score_row(t2, gs2, "imp", "Max score", imp_r, r, "max_score")
        imp_per_lbl = wx.StaticText(t2, label="Per hit:")
        gs2.Add(imp_per_lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL)
        imp_per_entry = wx.TextCtrl(t2, value=str(imp_r.get("per_hit", 5)), size=(50, -1))
        gs2.Add(imp_per_entry, pos=(r, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)
        r += 1
        r = keyword_editor(t2, gs2, "imp_kw", "Keywords:", imp_r.get("keywords", []), r)

        r = section(t2, gs2, "Low Priority Keywords", r)
        lp_r = rules.get("low_priority_keywords", {})
        r = score_row(t2, gs2, "lp", "Max score", lp_r, r, "max_score")
        lp_per_lbl = wx.StaticText(t2, label="Per hit:")
        gs2.Add(lp_per_lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL)
        lp_per_entry = wx.TextCtrl(t2, value=str(lp_r.get("per_hit", 8)), size=(50, -1))
        gs2.Add(lp_per_entry, pos=(r, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)
        r += 1
        r = keyword_editor(t2, gs2, "lp_kw", "Keywords:", lp_r.get("keywords", []), r)

        r = section(t2, gs2, "Calendar Keywords", r)
        r = score_row(t2, gs2, "cal", "Calendar score", rules.get("calendar_keywords", {}), r)
        r = keyword_editor(t2, gs2, "cal_kw", "Keywords:",
                           rules.get("calendar_keywords", {}).get("keywords", []), r)

        # ── Tab 3: Signals & Scoring ──
        t3, gs3 = make_scroll_tab("  Signals & Scoring  ")
        r = 0
        r = section(t3, gs3, "Base Score", r)
        base_lbl = wx.StaticText(t3, label="Starting score for every email:")
        gs3.Add(base_lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL)
        base_score_entry = wx.TextCtrl(t3, value=str(rules.get("base_score", 30)), size=(50, -1))
        gs3.Add(base_score_entry, pos=(r, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)
        r += 1
        base_hint = wx.StaticText(t3, label="  All rules add to or subtract from this base.")
        base_hint.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        gs3.Add(base_hint, pos=(r, 1), span=(1, 3), flag=wx.ALIGN_CENTER_VERTICAL | wx.BOTTOM, border=4)
        r += 1

        r = section(t3, gs3, "Static Signals", r)
        ss = rules.get("static_signals", {})
        for sig_key, sig_label in [
            ("unread", "Unread email"),
            ("high_importance", "Marked high importance"),
            ("low_importance", "Marked low importance"),
            ("flagged", "Flagged"),
            ("filtered_other", "Filtered to Other"),
            ("direct_to", "Direct recipient (To:)"),
            ("cc_only", "CC'd only"),
            ("has_attachments", "Has attachments"),
            ("short_message", "Short message"),
        ]:
            r = score_row(t3, gs3, f"ss_{sig_key}", sig_label, ss.get(sig_key, {}), r)

        r = section(t3, gs3, "Recency Scores", r)
        rec = rules.get("recency_scores", {})
        rec_enabled_cb = wx.CheckBox(t3)
        rec_enabled_cb.SetValue(rec.get("enabled", True))
        gs3.Add(rec_enabled_cb, pos=(r, 0), flag=wx.LEFT | wx.ALIGN_CENTER_VERTICAL, border=12)
        rec_lbl = wx.StaticText(t3, label="Enable recency scoring")
        gs3.Add(rec_lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL)
        r += 1
        rec_vars = {}
        for rk, rk_label in [("under_1h", "< 1 hour"), ("under_4h", "< 4 hours"),
                               ("under_24h", "< 24 hours"), ("over_7d", "> 7 days")]:
            lbl = wx.StaticText(t3, label=f"  {rk_label}:")
            gs3.Add(lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL)
            entry = wx.TextCtrl(t3, value=str(rec.get(rk, 0)), size=(50, -1))
            gs3.Add(entry, pos=(r, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)
            rec_vars[rk] = entry
            r += 1

        r = section(t3, gs3, "Priority Thresholds", r)
        th = rules.get("priority_thresholds", {})
        th_vars = {}
        for th_key, th_label in [("urgent", "Urgent ≥"), ("important", "Important ≥"),
                                   ("normal", "Normal ≥")]:
            lbl = wx.StaticText(t3, label=f"  {th_label}:")
            gs3.Add(lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL)
            entry = wx.TextCtrl(t3, value=str(th.get(th_key, 0)), size=(50, -1))
            gs3.Add(entry, pos=(r, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)
            th_vars[th_key] = entry
            r += 1

        r = section(t3, gs3, "Automated Senders", r)
        r = score_row(t3, gs3, "auto", "Automated sender score", rules.get("automated_senders", {}), r)
        r = keyword_editor(t3, gs3, "auto_patterns", "Patterns (matched in sender email):",
                           rules.get("automated_senders", {}).get("patterns", []), r)
        r = keyword_editor(t3, gs3, "auto_ext", "Extended automated patterns:",
                           rules.get("automated_senders_extended", []), r)

        r = section(t3, gs3, "Name + Question Detection", r)
        nq = rules.get("name_question_detection", {})
        r = score_row(t3, gs3, "nq", "Name + question score", nq, r)
        r = keyword_editor(t3, gs3, "nq_patterns", "Question patterns (regex):",
                           nq.get("question_patterns", []), r)

        r = section(t3, gs3, "General Question Detection", r)
        gq = rules.get("general_question_detection", {})
        r = score_row(t3, gs3, "gq", "Question detected score", gq, r)
        r = keyword_editor(t3, gs3, "gq_patterns", "Patterns (regex):",
                           gq.get("patterns", []), r)

        r = section(t3, gs3, "Category Keywords", r)
        cat = rules.get("category_keywords", {})
        r = keyword_editor(t3, gs3, "cat_action", "Action category:", cat.get("action", []), r)
        r = keyword_editor(t3, gs3, "cat_fyi", "FYI category:", cat.get("fyi", []), r)

        r = section(t3, gs3, "Auto Archive Senders", r)
        aa = rules.get("auto_archive_senders", {})
        aa_enabled_cb = wx.CheckBox(t3)
        aa_enabled_cb.SetValue(aa.get("enabled", True))
        gs3.Add(aa_enabled_cb, pos=(r, 0), flag=wx.LEFT | wx.ALIGN_CENTER_VERTICAL, border=12)
        aa_lbl = wx.StaticText(t3, label="Auto-archive emails from these senders")
        gs3.Add(aa_lbl, pos=(r, 1), span=(1, 2), flag=wx.ALIGN_CENTER_VERTICAL)
        r += 1
        r = keyword_editor(t3, gs3, "aa_entries", "Senders (one per line):", aa.get("entries", []), r)

        # ── Tab 4: Send & Undo ──
        t4, gs4 = make_scroll_tab("  Send & Undo  ")
        r = 0
        r = section(t4, gs4, "Undo Send", r)
        undo_lbl = wx.StaticText(t4, label="Delay before sending (seconds):")
        gs4.Add(undo_lbl, pos=(r, 1), flag=wx.ALIGN_CENTER_VERTICAL | wx.LEFT, border=12)
        undo_seconds_entry = wx.TextCtrl(t4, value=str(self.config.get("undo_send_seconds", 60)), size=(60, -1))
        gs4.Add(undo_seconds_entry, pos=(r, 2), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=12)
        r += 1

        r = section(t4, gs4, "Send Later Options", r)
        sl_hint = wx.StaticText(t4, label="JSON — each entry needs 'label' and 'hours' or 'preset'.")
        sl_hint.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        gs4.Add(sl_hint, pos=(r, 0), span=(1, 4), flag=wx.LEFT | wx.BOTTOM, border=12)
        r += 1
        send_later_text = wx.TextCtrl(t4, style=wx.TE_MULTILINE | wx.TE_DONTWRAP, size=(500, 200))
        send_later_text.SetValue(json.dumps(
            self.config.get("send_later_options", DEFAULT_CONFIG["send_later_options"]), indent=2))
        gs4.Add(send_later_text, pos=(r, 0), span=(1, 4), flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=12)
        text_widgets["send_later_options"] = send_later_text
        r += 1

        r = section(t4, gs4, "Snooze Options (S key)", r)
        sn_hint = wx.StaticText(t4, label="JSON — each entry needs 'label' and 'hours' or 'preset'.")
        sn_hint.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        gs4.Add(sn_hint, pos=(r, 0), span=(1, 4), flag=wx.LEFT | wx.BOTTOM, border=12)
        r += 1
        snooze_text = wx.TextCtrl(t4, style=wx.TE_MULTILINE | wx.TE_DONTWRAP, size=(500, 140))
        snooze_text.SetValue(json.dumps(
            self.config.get("snooze_options", DEFAULT_CONFIG["snooze_options"]), indent=2))
        gs4.Add(snooze_text, pos=(r, 0), span=(1, 4), flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=12)
        text_widgets["snooze_options"] = snooze_text
        r += 1

        r = section(t4, gs4, "Remind Options (M key)", r)
        rm_hint = wx.StaticText(t4, label="JSON — each entry needs 'label' and 'days'.")
        rm_hint.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        gs4.Add(rm_hint, pos=(r, 0), span=(1, 4), flag=wx.LEFT | wx.BOTTOM, border=12)
        r += 1
        remind_text = wx.TextCtrl(t4, style=wx.TE_MULTILINE | wx.TE_DONTWRAP, size=(500, 120))
        remind_text.SetValue(json.dumps(
            self.config.get("remind_options", DEFAULT_CONFIG["remind_options"]), indent=2))
        gs4.Add(remind_text, pos=(r, 0), span=(1, 4), flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=12)
        text_widgets["remind_options"] = remind_text
        r += 1

        # ── Bottom buttons ──
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        def _get_text_lines(key):
            if key in text_widgets:
                return [l.strip() for l in text_widgets[key].GetValue().strip().split("\n") if l.strip()]
            return []

        def _save(evt=None):
            try:
                new_rules = json.loads(json.dumps(DEFAULT_SCORING_RULES))
                try:
                    new_rules["base_score"] = int(base_score_entry.GetValue())
                except ValueError:
                    new_rules["base_score"] = 30

                new_rules["vip_senders"]["enabled"] = vars_store["vip_sender"]["enabled"].GetValue()
                new_rules["vip_senders"]["sender_score"] = int(vars_store["vip_sender"]["score"].GetValue())
                new_rules["vip_senders"]["recipient_score"] = int(vars_store["vip_recipient"]["score"].GetValue())
                new_rules["vip_senders"]["label"] = vip_label_entry.GetValue().strip() or "VIP"
                new_rules["vip_senders"]["entries"] = _get_text_lines("vip_entries")

                new_rules["critical_subjects"]["enabled"] = vars_store["crit_subj"]["enabled"].GetValue()
                new_rules["critical_subjects"]["score"] = int(vars_store["crit_subj"]["score"].GetValue())
                new_rules["critical_subjects"]["patterns"] = _get_text_lines("crit_subj_patterns")

                new_rules["critical_subject_keywords"]["enabled"] = vars_store["crit_kw"]["enabled"].GetValue()
                new_rules["critical_subject_keywords"]["score"] = int(vars_store["crit_kw"]["score"].GetValue())
                new_rules["critical_subject_keywords"]["patterns"] = _get_text_lines("crit_kw_patterns")

                new_rules["critical_sender_domains"]["enabled"] = vars_store["crit_dom"]["enabled"].GetValue()
                new_rules["critical_sender_domains"]["score"] = int(vars_store["crit_dom"]["score"].GetValue())
                new_rules["critical_sender_domains"]["domains"] = _get_text_lines("crit_dom_list")

                try:
                    cond_json = cond_txt.GetValue().strip()
                    new_rules["conditional_rules"] = json.loads(cond_json)
                except json.JSONDecodeError as je:
                    showerror("Invalid JSON", f"Conditional rules JSON is invalid:\n{je}", win)
                    return

                new_rules["urgent_keywords"]["enabled"] = vars_store["urg"]["enabled"].GetValue()
                new_rules["urgent_keywords"]["max_score"] = int(vars_store["urg"]["score"].GetValue())
                new_rules["urgent_keywords"]["per_hit"] = int(urg_per_entry.GetValue())
                new_rules["urgent_keywords"]["keywords"] = _get_text_lines("urg_kw")

                new_rules["important_keywords"]["enabled"] = vars_store["imp"]["enabled"].GetValue()
                new_rules["important_keywords"]["max_score"] = int(vars_store["imp"]["score"].GetValue())
                new_rules["important_keywords"]["per_hit"] = int(imp_per_entry.GetValue())
                new_rules["important_keywords"]["keywords"] = _get_text_lines("imp_kw")

                new_rules["low_priority_keywords"]["enabled"] = vars_store["lp"]["enabled"].GetValue()
                new_rules["low_priority_keywords"]["max_score"] = int(vars_store["lp"]["score"].GetValue())
                new_rules["low_priority_keywords"]["per_hit"] = int(lp_per_entry.GetValue())
                new_rules["low_priority_keywords"]["keywords"] = _get_text_lines("lp_kw")

                new_rules["calendar_keywords"]["enabled"] = vars_store["cal"]["enabled"].GetValue()
                new_rules["calendar_keywords"]["score"] = int(vars_store["cal"]["score"].GetValue())
                new_rules["calendar_keywords"]["keywords"] = _get_text_lines("cal_kw")

                for sig_key in ["unread", "high_importance", "low_importance", "flagged",
                                "filtered_other", "direct_to", "cc_only", "has_attachments",
                                "short_message"]:
                    vs = vars_store.get(f"ss_{sig_key}")
                    if vs:
                        new_rules["static_signals"][sig_key]["enabled"] = vs["enabled"].GetValue()
                        new_rules["static_signals"][sig_key]["score"] = int(vs["score"].GetValue())

                new_rules["recency_scores"]["enabled"] = rec_enabled_cb.GetValue()
                for rk, entry in rec_vars.items():
                    new_rules["recency_scores"][rk] = int(entry.GetValue())

                for th_key, entry in th_vars.items():
                    new_rules["priority_thresholds"][th_key] = int(entry.GetValue())

                new_rules["automated_senders"]["enabled"] = vars_store["auto"]["enabled"].GetValue()
                new_rules["automated_senders"]["score"] = int(vars_store["auto"]["score"].GetValue())
                new_rules["automated_senders"]["patterns"] = _get_text_lines("auto_patterns")
                new_rules["automated_senders_extended"] = _get_text_lines("auto_ext")

                new_rules["name_question_detection"]["enabled"] = vars_store["nq"]["enabled"].GetValue()
                new_rules["name_question_detection"]["score"] = int(vars_store["nq"]["score"].GetValue())
                new_rules["name_question_detection"]["question_patterns"] = _get_text_lines("nq_patterns")

                new_rules["general_question_detection"]["enabled"] = vars_store["gq"]["enabled"].GetValue()
                new_rules["general_question_detection"]["score"] = int(vars_store["gq"]["score"].GetValue())
                new_rules["general_question_detection"]["patterns"] = _get_text_lines("gq_patterns")

                new_rules["category_keywords"]["action"] = _get_text_lines("cat_action")
                new_rules["category_keywords"]["fyi"] = _get_text_lines("cat_fyi")

                new_rules["auto_archive_senders"]["enabled"] = aa_enabled_cb.GetValue()
                new_rules["auto_archive_senders"]["entries"] = _get_text_lines("aa_entries")

                save_scoring_rules(new_rules)

                try:
                    undo_s = int(undo_seconds_entry.GetValue())
                    self.config["undo_send_seconds"] = max(0, undo_s)
                except ValueError:
                    pass

                for cfg_key, wid in [("send_later_options", send_later_text),
                                      ("snooze_options", snooze_text),
                                      ("remind_options", remind_text)]:
                    try:
                        opts = json.loads(wid.GetValue().strip())
                        if isinstance(opts, list):
                            self.config[cfg_key] = opts
                    except json.JSONDecodeError as je:
                        showerror("Invalid JSON", f"{cfg_key} JSON is invalid:\n{je}", win)
                        return

                save_config(self.config)
                if hasattr(self, 'intelligence') and self.intelligence:
                    self.intelligence.reload_rules()
                win.EndModal(wx.ID_OK)
                self._rescore_and_rerender()
            except ValueError as ve:
                showerror("Invalid Value", f"Please enter valid numbers for all score fields.\n{ve}", win)

        def _reset(evt=None):
            if askyesno("Reset Rules", "Reset all scoring rules to defaults?\nThis cannot be undone.", win):
                save_scoring_rules(DEFAULT_SCORING_RULES)
                if hasattr(self, 'intelligence') and self.intelligence:
                    self.intelligence.reload_rules()
                win.EndModal(wx.ID_CANCEL)
                self._rescore_and_rerender()
                self._open_scoring_settings()

        save_btn = wx.Button(win, label="💾 Save & Close")
        save_btn.Bind(wx.EVT_BUTTON, _save)
        btn_sizer.Add(save_btn, 0, wx.RIGHT, 8)

        reset_btn = wx.Button(win, label="Reset to Defaults")
        reset_btn.Bind(wx.EVT_BUTTON, _reset)
        btn_sizer.Add(reset_btn, 0, wx.RIGHT, 8)

        cancel_btn = wx.Button(win, label="Cancel")
        cancel_btn.Bind(wx.EVT_BUTTON, lambda e: win.EndModal(wx.ID_CANCEL))
        btn_sizer.Add(cancel_btn, 0)

        outer.Add(btn_sizer, 0, wx.ALL, 12)
        win.SetSizer(outer)
        win.ShowModal()
        win.Destroy()
