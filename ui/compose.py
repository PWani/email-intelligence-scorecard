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


class ComposeMixin:
    """Compose and Reply"""

    def _build_signature_html(self):
        """Build HTML email signature matching exact Outlook mobile signature format."""
        p = getattr(self, '_user_profile', {})
        if not p:
            return ""

        name = p.get("displayName", "")
        title = p.get("jobTitle", "")
        email_addr = p.get("mail", "")

        # Phone: prefer mobilePhone, fall back to businessPhones[0]
        phone = p.get("mobilePhone", "")
        if not phone:
            phones = p.get("businessPhones", [])
            phone = phones[0] if phones else ""

        if not name:
            return ""

        # Check if user toggled signature off
        if not getattr(self, '_reply_include_sig', True):
            return ""

        # Match exact Outlook signature HTML structure
        style = ('style="text-align:left;text-indent:0px;text-transform:none;margin:0in;'
                 'font-family:Aptos,Aptos_MSFontService,-apple-system,Roboto,Arial,Helvetica,sans-serif;'
                 'font-size:12pt"')

        lines = []
        lines.append(f'<p class="MsoNormal" {style}>_________________________</p>')
        lines.append(f'<p class="MsoNormal" {style}><b>{name}</b></p>')
        if title:
            lines.append(f'<p class="MsoNormal" {style}>{title}</p>')
        company = self.config.get("signature_company", "")
        address1 = self.config.get("signature_address1", "")
        address2 = self.config.get("signature_address2", "")
        website = self.config.get("signature_website", "")
        if company:
            lines.append(f'<p class="MsoNormal" {style}>{company}</p>')
        if address1:
            lines.append(f'<p class="MsoNormal" {style}>{address1}</p>')
        if address2:
            lines.append(f'<p class="MsoNormal" {style}>{address2}</p>')
        if phone:
            lines.append(
                f'<p class="MsoNormal" {style}>Tel: '
                f'<a href="tel:{phone.replace(" ", "").replace("(", "").replace(")", "").replace("-", "")}" '
                f'style="color:rgb(15,108,189);margin-top:0px;margin-bottom:0px">{phone}</a></p>')
        # Email | Website
        email_link = f'<a href="mailto:{email_addr}" style="color:rgb(0,120,212);margin-top:0px;margin-bottom:0px">{email_addr}</a>'
        if website:
            lines.append(
                f'<p class="MsoNormal" {style}>'
                f'{email_link}'
                f' | <a href="http://{website}" style="color:rgb(0,120,212);margin-top:0px;margin-bottom:0px">{website}</a>'
                f'</p>')
        else:
            lines.append(f'<p class="MsoNormal" {style}>{email_link}</p>')

        sig_html = (
            '<div style="font-family:Aptos,Aptos_MSFontService,-apple-system,Roboto,Arial,Helvetica,sans-serif;'
            'font-size:12pt"><br></div>'
            '<div id="ms-outlook-mobile-signature" style="font-family:Aptos,Aptos_MSFontService,-apple-system,'
            'Roboto,Arial,Helvetica,sans-serif;font-size:12pt">'
            + "\n".join(lines) +
            '<div style="font-family:Aptos,Aptos_MSFontService,-apple-system,Roboto,Arial,Helvetica,sans-serif;'
            'font-size:12pt;color:rgb(33,33,33)"><br></div>'
            '</div>'
        )
        return sig_html

    def _should_include_signature(self, email_data):
        """Determine whether to include signature.
        Include signature if:
        - This is the first time the user is replying in this conversation thread
        - This is a forward (always include)
        Returns True/False.
        """
        # Forwards always get a signature
        if getattr(self, '_is_forward', False):
            return True

        if not email_data or not self.graph:
            return True  # Default to include if we can't determine

        conv_id = email_data.get("conversationId", "")
        if not conv_id:
            return True  # No conversation ID — treat as new, include sig

        # Check if we've already sent in this conversation
        try:
            sent_count = self.graph.get_sent_count_for_conversation(conv_id)
            return sent_count == 0  # Include sig only if we haven't sent before
        except Exception:
            return True  # On error, include sig to be safe

    # ── New Compose ─────────────────────────────────────────────

    def _compose_new(self):
        """Open a standalone compose window for a new email."""
        log.info("[compose] new email")

        # Determine which client to use
        client = self.graph or self.google_client
        if not client:
            showerror("No Account", "Sign in to an email account first.")
            return

        win = wx.Frame(self.root, title="✉ New Message",
                       size=(700, 580),
                       style=wx.DEFAULT_FRAME_STYLE)
        win.SetMinSize((500, 400))
        win.SetBackgroundColour(_hex(C['bg_card']))
        sizer = wx.BoxSizer(wx.VERTICAL)

        # ── To field ──────────────────────────────────────────
        to_row = wx.BoxSizer(wx.HORIZONTAL)
        to_lbl = wx.StaticText(win, label="To:")
        to_lbl.SetMinSize((55, -1))
        to_row.Add(to_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)
        to_entry = wx.TextCtrl(win)
        to_row.Add(to_entry, 1, wx.EXPAND)
        sizer.Add(to_row, 0, wx.EXPAND | wx.ALL, 8)

        # ── CC field ──────────────────────────────────────────
        cc_row = wx.BoxSizer(wx.HORIZONTAL)
        cc_lbl = wx.StaticText(win, label="CC:")
        cc_lbl.SetMinSize((55, -1))
        cc_row.Add(cc_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)
        cc_entry = wx.TextCtrl(win)
        cc_row.Add(cc_entry, 1, wx.EXPAND)
        sizer.Add(cc_row, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)

        # ── Subject field ─────────────────────────────────────
        subj_row = wx.BoxSizer(wx.HORIZONTAL)
        subj_lbl = wx.StaticText(win, label="Subject:")
        subj_lbl.SetMinSize((55, -1))
        subj_row.Add(subj_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)
        subj_entry = wx.TextCtrl(win)
        subj_row.Add(subj_entry, 1, wx.EXPAND)
        sizer.Add(subj_row, 0, wx.EXPAND | wx.ALL, 8)

        # ── Body ──────────────────────────────────────────────
        body_entry = wx.TextCtrl(win, style=wx.TE_MULTILINE | wx.TE_RICH2)
        body_entry.SetFont(_font(FONT, 11))
        sizer.Add(body_entry, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)

        # ── Signature ─────────────────────────────────────────
        sig_check = wx.CheckBox(win, label="Include signature")
        sig_check.SetValue(True)
        sizer.Add(sig_check, 0, wx.LEFT | wx.TOP, 8)

        sig_text = self._build_signature_plain()
        sig_label = wx.StaticText(win, label=sig_text)
        sig_label.SetForegroundColour(_hex(C['muted']))
        sig_label.SetFont(_font(FONT, 9))
        sizer.Add(sig_label, 0, wx.LEFT | wx.BOTTOM, 8)

        def _toggle_sig(evt):
            sig_label.Show(sig_check.GetValue())
            win.Layout()
        sig_check.Bind(wx.EVT_CHECKBOX, _toggle_sig)

        # ── Status bar for spell check ────────────────────────
        status_label = wx.StaticText(win, label="")
        status_label.SetForegroundColour(_hex(C['muted']))
        status_label.SetFont(_font(FONT, 9))
        sizer.Add(status_label, 0, wx.LEFT | wx.RIGHT, 8)

        # ── Buttons ───────────────────────────────────────────
        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        send_btn = wx.Button(win, label="📤 Send")
        send_btn.SetBackgroundColour(_hex(C['btn_primary']))
        send_btn.SetForegroundColour(wx.WHITE)
        btn_row.Add(send_btn, 0, wx.RIGHT, 8)

        fix_btn = wx.Button(win, label="🔧 Fix All")
        btn_row.Add(fix_btn, 0, wx.RIGHT, 8)

        cancel_btn = wx.Button(win, label="Cancel")
        btn_row.Add(cancel_btn, 0)
        sizer.Add(btn_row, 0, wx.ALL, 8)

        # ── Spell check wiring ────────────────────────────────
        _spell_timer = [None]
        _spell_errors = [[]]

        def _clear_spell(text_ctrl):
            _spell_errors[0] = []
            try:
                length = text_ctrl.GetLastPosition()
                if length > 0:
                    attr = wx.TextAttr()
                    attr.SetBackgroundColour(text_ctrl.GetBackgroundColour())
                    text_ctrl.SetStyle(0, length, attr)
            except Exception:
                pass

        def _apply_marks(checked_text, errors, text_ctrl):
            current = text_ctrl.GetValue().strip()
            if current != checked_text:
                return
            _clear_spell(text_ctrl)
            _spell_errors[0] = errors
            if not errors:
                status_label.SetLabel("No spelling or grammar issues ✓")
                return
            for err in errors:
                start = err.get("offset", 0)
                end = start + err.get("length", 0)
                if end <= start:
                    continue
                is_spell = any(r in err.get("rule", "")
                               for r in ["SPEL", "SPELL", "MORFOLOGIK", "HUNSPELL"])
                colour = wx.Colour(255, 255, 150) if is_spell else wx.Colour(255, 200, 120)
                try:
                    attr = wx.TextAttr()
                    attr.SetBackgroundColour(colour)
                    text_ctrl.SetStyle(start, end, attr)
                except Exception:
                    pass
            cs = sum(1 for e in errors if any(r in e.get("rule", "")
                     for r in ["SPEL", "SPELL", "MORFOLOGIK", "HUNSPELL"]))
            cg = len(errors) - cs
            parts = []
            if cs: parts.append(f"{cs} spelling")
            if cg: parts.append(f"{cg} grammar")
            status_label.SetLabel(
                f"Found {' and '.join(parts)} issue{'s' if len(errors) > 1 else ''}")

        def _run_spell():
            text = body_entry.GetValue().strip()
            if not text:
                _clear_spell(body_entry)
                status_label.SetLabel("")
                return
            status_label.SetLabel("Checking spelling...")
            def run():
                errors = self.spell_checker.check(text)
                wx.CallAfter(lambda: _apply_marks(text, errors, body_entry))
            threading.Thread(target=run, daemon=True).start()

        def _on_body_key(evt):
            evt.Skip()
            if _spell_timer[0]:
                _spell_timer[0].Stop()
            _spell_timer[0] = wx.CallLater(800, _run_spell)

        body_entry.Bind(wx.EVT_TEXT, _on_body_key)

        def _on_fix_all(evt):
            text = body_entry.GetValue().strip()
            if not text:
                return
            fix_btn.Disable()
            status_label.SetLabel("Fixing...")
            def run():
                fixed = self.spell_checker.auto_fix(text)
                def apply():
                    current = body_entry.GetValue().strip()
                    if current == text and fixed != text:
                        body_entry.SetValue(fixed)
                        _clear_spell(body_entry)
                        status_label.SetLabel("All errors fixed ✓")
                        if _spell_timer[0]:
                            _spell_timer[0].Stop()
                        _spell_timer[0] = wx.CallLater(500, _run_spell)
                    elif fixed == text:
                        status_label.SetLabel("No fixes needed")
                    fix_btn.Enable()
                wx.CallAfter(apply)
            threading.Thread(target=run, daemon=True).start()

        fix_btn.Bind(wx.EVT_BUTTON, _on_fix_all)

        win.SetSizer(sizer)

        cancel_btn.Bind(wx.EVT_BUTTON, lambda e: win.Destroy())

        # ── Focus helpers ─────────────────────────────────────
        # Prevent main frame's EVT_CHAR_HOOK from intercepting keys in this window
        def _compose_char_hook(e):
            e.Skip()
        win.Bind(wx.EVT_CHAR_HOOK, _compose_char_hook)

        # Each TextCtrl needs click→focus+caret and focus→caret handlers
        def _make_click_focus(ctrl):
            def _handler(e):
                def _do():
                    ctrl.SetFocus()
                    ctrl.SetInsertionPointEnd()
                wx.CallAfter(_do)
                e.Skip()
            return _handler

        def _make_set_focus(ctrl):
            def _handler(e):
                wx.CallAfter(ctrl.SetInsertionPointEnd)
                e.Skip()
            return _handler

        for ctrl in (to_entry, cc_entry, subj_entry, body_entry):
            ctrl.Bind(wx.EVT_LEFT_DOWN, _make_click_focus(ctrl))
            ctrl.Bind(wx.EVT_SET_FOCUS, _make_set_focus(ctrl))

        def _on_send_new(evt):
            to_raw = to_entry.GetValue().strip()
            to_addrs = [a.strip() for a in to_raw.replace(",", ";").split(";")
                        if a.strip() and "@" in a.strip()]
            if not to_addrs:
                showerror("Invalid Address",
                          "Enter valid email address(es), separated by ;",
                          parent=win)
                return
            subject = subj_entry.GetValue().strip()
            body = body_entry.GetValue().strip()
            if not subject and not body:
                showerror("Empty", "Enter a subject or message body.",
                          parent=win)
                return

            cc_raw = cc_entry.GetValue().strip()
            cc_addrs = [a.strip() for a in cc_raw.replace(",", ";").split(";")
                        if a.strip() and "@" in a.strip()] if cc_raw else None

            # Build HTML
            body_html = body.replace("\n", "<br>") if body else ""
            sig_html = ""
            if sig_check.GetValue():
                # Temporarily set flag so _build_signature_html works
                old_flag = getattr(self, '_reply_include_sig', True)
                self._reply_include_sig = True
                sig_html = self._build_signature_html()
                self._reply_include_sig = old_flag
            full_html = (
                '<div style="font-family:Aptos,Aptos_MSFontService,-apple-system,'
                'Roboto,Arial,Helvetica,sans-serif;font-size:12pt;color:rgb(33,33,33)">'
                f'<div>{body_html}</div>{sig_html}</div>'
            )

            send_btn.SetLabel("Sending...")
            send_btn.Disable()
            display = f"New to {'; '.join(to_addrs)}"
            delay_s = self.config.get("undo_send_seconds", 60)
            send_at = datetime.now(timezone.utc) + timedelta(seconds=delay_s)

            def _do_send():
                try:
                    draft = client.create_draft(
                        subject, full_html, to_addrs, cc_addresses=cc_addrs)
                    draft_id = draft.get("id")
                    if not draft_id:
                        wx.CallAfter(lambda: showerror(
                            "Send Failed", "Could not create draft.",
                            parent=win))
                        wx.CallAfter(lambda: (
                            send_btn.SetLabel("📤 Send"),
                            send_btn.Enable()))
                        return

                    # Gmail sends immediately
                    if draft.get("_sent"):
                        log.info("[send] New email sent (Google): %s", display)
                        wx.CallAfter(lambda: self._set_status(
                            f"✓ {display} sent"))
                        wx.CallAfter(win.Destroy)
                        return

                    # MS: queue with undo-send delay
                    cats = [f"send_at:{send_at.isoformat()}"]
                    client.set_email_categories(draft_id, cats)
                    client.move_to_send_queue(draft_id)
                    log.info("[send] Queued new: %s | draft=%s", display,
                             draft_id[:40])
                    wx.CallAfter(win.Destroy)
                    wx.CallAfter(lambda: self._show_send_undo_bar(
                        draft_id, display, delay_s))
                except Exception as e:
                    log.error("[send] New email failed: %s", str(e)[:200])
                    wx.CallAfter(lambda: showerror(
                        "Send Failed", str(e)[:300], parent=win))
                    wx.CallAfter(lambda: (
                        send_btn.SetLabel("📤 Send"),
                        send_btn.Enable()))

            threading.Thread(target=_do_send, daemon=True).start()

        send_btn.Bind(wx.EVT_BUTTON, _on_send_new)

        win.Show()
        def _initial_focus():
            to_entry.SetFocus()
            to_entry.SetInsertionPointEnd()
        wx.CallAfter(_initial_focus)

        # Attach autocomplete after dialog is fully shown
        def _attach_ac():
            try:
                EmailAutocomplete(to_entry,
                                  lambda: getattr(self, '_address_book', []))
                EmailAutocomplete(cc_entry,
                                  lambda: getattr(self, '_address_book', []))
            except Exception:
                pass
        wx.CallAfter(_attach_ac)

    # ── Reply / Forward ───────────────────────────────────────

    def _update_compose_btn_styles(self, active=None):
        """Highlight the active compose button (reply/reply_all/forward), reset others."""
        for btn, key in [(self._reply_btn, "reply"),
                         (self._reply_all_btn, "reply_all"),
                         (self._forward_btn, "forward")]:
            if key == active:
                btn.SetBackgroundColour(_hex(C["btn_primary"])); btn.SetForegroundColour(wx.WHITE)
            else:
                btn.SetBackgroundColour(wx.NullColour); btn.SetForegroundColour(wx.NullColour)

    def _reply(self):
        log.info("[compose] reply: eid=%s", (self.selected_email_id or "")[:40])
        if not self.selected_email_id: return
        self._is_reply_all = False
        self._is_forward = False
        self.reply_label.SetLabel("✏️ Reply")
        self._update_compose_btn_styles("reply")
        self._show_reply()

    def _reply_all(self):
        log.info("[compose] reply-all: eid=%s", (self.selected_email_id or "")[:40])
        if not self.selected_email_id: return
        self._is_reply_all = True
        self._is_forward = False
        self.reply_label.SetLabel("✏️ Reply All")
        self._update_compose_btn_styles("reply_all")
        self._show_reply()

    def _forward(self):
        log.info("[compose] forward: eid=%s", (self.selected_email_id or "")[:40])
        if not self.selected_email_id: return
        self._is_reply_all = False
        self._is_forward = True
        self.reply_label.SetLabel("↪ Forward")
        self._update_compose_btn_styles("forward")
        self._show_reply()

    def _show_reply(self):
        """Show reply/forward composer between action bar and body."""
        self.reply_frame.Hide()
        self._body_container.Hide()
        # Hide edit fields (start collapsed)
        self._edit_subject_frame.Hide()
        self._edit_to_frame.Hide()
        self._edit_subject_visible = False
        self._edit_to_visible = False
        # Show/hide the forward To: field and CC field
        if getattr(self, '_is_forward', False):
            self._fwd_to_frame.Show()
            self._fwd_to_entry.SetValue("") if hasattr(self, "_fwd_to_entry") else None
            self._cc_frame.Hide()
            self._cc_entry.SetValue("") if hasattr(self, "_cc_entry") else None
            # Hide edit recipients link — forward has its own To field
            self._edit_recipients_link.Hide()
        else:
            self._fwd_to_frame.Hide()
            self._cc_frame.Show()
            self._cc_entry.SetValue("") if hasattr(self, "_cc_entry") else None
            # Show edit recipients link for replies
            self._edit_recipients_link.Show()
        self.reply_frame.Show()
        self._body_container.Show()
        self._detail_panel.Layout()

        # Populate edit fields with current email data
        em = next((e for e in self.emails if e.get("id") == self.selected_email_id), None)
        if em:
            self._edit_subject_entry.SetValue(em.get("subject", ""))
            # Build recipients string: for Reply use sender, for Reply All use sender + To + CC minus self
            if getattr(self, '_is_forward', False):
                self._edit_to_entry.SetValue("") if hasattr(self, "_edit_to_entry") else None  # forward has its own To field
            elif self._is_reply_all:
                sender_addr = em.get("from", {}).get("emailAddress", {}).get("address", "")
                all_to = [r.get("emailAddress", {}).get("address", "") for r in em.get("toRecipients", [])]
                all_cc = [r.get("emailAddress", {}).get("address", "") for r in em.get("ccRecipients", [])]
                # Combine: sender first, then other To recipients, exclude self
                me = self._ms_email.lower() or (self.intelligence.user_email if self.intelligence else "")
                reply_to = [sender_addr] + [a for a in all_to if a.lower() != me.lower() and a.lower() != sender_addr.lower()]
                self._edit_to_entry.SetValue("; ".join(a for a in reply_to if a))
                # Pre-fill CC with original CC minus self
                reply_cc = [a for a in all_cc if a.lower() != me.lower()]
                if reply_cc:
                    self._cc_entry.SetValue("; ".join(reply_cc))
            else:
                # Simple reply — To is the sender
                sender_addr = em.get("from", {}).get("emailAddress", {}).get("address", "")
                self._edit_to_entry.SetValue(sender_addr)

        # Determine signature and show preview
        self._sig_preview_frame.Hide()
        self._reply_include_sig = False

        def _check_sig():
            include = False
            if em:
                include = self._should_include_signature(em)
            wx.CallAfter(lambda: self._show_sig_preview(include))

        # Run signature check in background to avoid blocking UI
        threading.Thread(target=_check_sig, daemon=True).start()

        # Layout must complete before SetFocus so the widget has real dimensions
        self.reply_frame.Layout()
        self._detail_panel.Layout()
        if getattr(self, '_is_forward', False):
            wx.CallAfter(self._fwd_to_entry.SetFocus)
        else:
            wx.CallAfter(self.reply_text.SetFocus)

    def _toggle_edit_subject(self):
        """Toggle the editable subject field."""
        if getattr(self, '_edit_subject_visible', False):
            self._edit_subject_frame.Hide()
            self._edit_subject_visible = False
        else:
            self._edit_subject_frame.Show()
            self._edit_subject_entry.SetFocus()
            self._edit_subject_visible = True
        self.reply_frame.Layout()
        self._detail_panel.Layout()

    def _toggle_edit_recipients(self):
        """Toggle the editable recipients field."""
        if getattr(self, '_edit_to_visible', False):
            self._edit_to_frame.Hide()
            self._edit_to_visible = False
        else:
            self._edit_to_frame.Show()
            self._edit_to_entry.SetFocus()
            self._edit_to_visible = True
        self.reply_frame.Layout()
        self._detail_panel.Layout()

    def _show_sig_preview(self, include):
        """Show or hide the signature preview in the reply composer."""
        self._reply_include_sig = include
        self._sig_include_var.SetValue(include)
        if include:
            sig_text = self._build_signature_plain()
            if sig_text:
                self._sig_preview_label.SetLabel(sig_text)
                self._sig_preview_frame.Show()
            else:
                self._sig_preview_frame.Hide()
        else:
            # Still show the frame with checkbox unchecked so user can re-enable
            sig_text = self._build_signature_plain()
            if sig_text:
                self._sig_preview_label.SetLabel("(signature hidden)")
                self._sig_preview_frame.Show()
            else:
                self._sig_preview_frame.Hide()
        # Trigger layout so detail panel reflows and Send button stays visible
        self.reply_frame.Layout()
        self._detail_panel.Layout()

    def _on_sig_toggle(self):
        """Handle signature checkbox toggle — show/hide signature preview pane."""
        include = self._sig_include_var.GetValue()
        self._reply_include_sig = include
        sig_text = self._build_signature_plain()
        if sig_text:
            # Always keep the frame visible so user can re-enable
            self._sig_preview_label.SetLabel(sig_text if include else "(signature hidden)")
            self._sig_preview_frame.Show()
        else:
            self._sig_preview_frame.Hide()
        self.reply_frame.Layout()
        self._detail_panel.Layout()

    def _build_signature_plain(self):
        """Build plain-text version of signature for preview in composer."""
        p = getattr(self, '_user_profile', {})
        if not p:
            return ""
        name = p.get("displayName", "")
        title = p.get("jobTitle", "")
        email_addr = p.get("mail", "")
        phone = p.get("mobilePhone", "")
        if not phone:
            phones = p.get("businessPhones", [])
            phone = phones[0] if phones else ""
        if not name:
            return ""
        lines = ["─" * 30, name]
        if title:
            lines.append(title)
        company = self.config.get("signature_company", "")
        address1 = self.config.get("signature_address1", "")
        address2 = self.config.get("signature_address2", "")
        website = self.config.get("signature_website", "")
        if company:
            lines.append(company)
        if address1:
            lines.append(address1)
        if address2:
            lines.append(address2)
        if phone:
            lines.append(f"Tel: {phone}")
        if website:
            lines.append(f"{email_addr} | {website}")
        else:
            lines.append(email_addr)
        return "\n".join(lines)

    def _on_send(self):
        log.info("[compose] send button pressed")
        """Queue email for sending via Send Queue folder.
        Uses Graph API createReply/createReplyAll/createForward to create proper drafts
        that include quoted body, threading, and attachments.
        The polling loop dispatches when due. Undo = delete the draft."""
        if getattr(self, '_sending', False):
            return
        self._sending = True
        self._send_btn.SetLabel("Sending..."); self._send_btn.Disable()

        # Gather data while composer is open
        is_fwd = getattr(self, '_is_forward', False)
        delay_s = self.config.get("undo_send_seconds", 60)
        send_at = datetime.now(timezone.utc) + timedelta(seconds=delay_s)

        if is_fwd:
            to_raw = self._fwd_to_entry.GetValue().strip()
            to_addrs = [a.strip() for a in to_raw.replace(",", ";").split(";")
                        if a.strip() and "@" in a.strip()]
            if not to_addrs:
                self._reset_send_btn()
                showinfo("Invalid Address", "Enter valid email address(es)")
                return
            self._send_forward()
            fwd_data = dict(self._pending_fwd_data) if self._pending_fwd_data else None
            if not fwd_data:
                self._reset_send_btn(); return
            display = f"Forward to {'; '.join(to_addrs)}"
        else:
            body = self.reply_text.GetValue().strip()
            if not body:
                self._reset_send_btn()
                showinfo("Empty", "Type a reply first.")
                return
            self._send_reply()
            reply_data = dict(self._pending_reply_data) if self._pending_reply_data else None
            if not reply_data:
                self._reset_send_btn(); return
            display = "Reply All" if reply_data["is_all"] else "Reply"

        self._cancel_reply()

        # Create proper draft and move to Send Queue in background
        def _queue_send():
            try:
                if is_fwd:
                    # createForward includes original body + attachments
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
                    log.error("[send] Draft creation returned no ID for %s", display)
                    wx.CallAfter(lambda: self._set_status("❌ Failed to create draft"))
                    wx.CallAfter(self._reset_send_btn)
                    return

                # Gmail sends immediately — no server-side queue/undo available
                if draft.get("_sent"):
                    log.info("[send] Google email sent immediately: %s", display)
                    wx.CallAfter(lambda: self._set_status(f"✓ {display} sent"))
                    wx.CallAfter(self._reset_send_btn)
                    return

                # MS path: tag with send time and move to Send Queue for undo-send delay
                cats = [f"send_at:{send_at.isoformat()}"]
                self._api_for(draft_id).set_email_categories(draft_id, cats)
                self._api_for(draft_id).move_to_send_queue(draft_id)
                log.info("[send] Queued: %s | draft_id=%s | send_at=%s | delay=%ss",
                         display, draft_id[:40], send_at.isoformat(), delay_s)
                # Show undo bar on UI thread
                wx.CallAfter(lambda: self._show_send_undo_bar(draft_id, display, delay_s))
            except Exception as e:
                err = str(e)
                log.error("[send] Queue failed for %s: %s", display, err[:200])
                if is_network_error(err):
                    # Queue the full send data for replay when online
                    if is_fwd:
                        self._offline_queue.enqueue("forward",
                            eid=fwd_data["eid"], to_addrs=fwd_data["to_addrs"],
                            comment_html=fwd_data["comment_html"])
                    else:
                        d = reply_data
                        self._offline_queue.enqueue("reply",
                            eid=d["eid"], is_all=d["is_all"], html=d["html"],
                            extra_cc=d.get("extra_cc"), edited_subject=d.get("edited_subject"),
                            edited_to=d.get("edited_to"))
                    wx.CallAfter(lambda: (
                        self._set_status(f"📴 Offline — {display} queued for when online"),
                        self._reset_send_btn()))
                else:
                    err_short = err[:120]
                    wx.CallAfter(lambda: (
                        self._reset_send_btn(),
                        showerror("Send Failed", f"Your reply could not be sent:\n\n{err_short}\n\n"
                            f"Please try again.")))

        threading.Thread(target=_queue_send, daemon=True).start()

    def _show_send_undo_bar(self, draft_id, description, delay_s):
        """Send queued silently in background — no banner shown."""
        self._undo_active_draft = draft_id
        self._undo_countdown_remaining = delay_s
        self._reset_send_btn()

    def _undo_queued_send(self, draft_id):
        """Cancel a queued send by deleting the draft from Send Queue."""
        log.info("[send-undo] Cancelling draft_id=%s", draft_id[:40] if draft_id else "None")
        self._dismiss_undo_bar()
        self._undo_active_draft = None
        def run():
            try:
                self._api_for(draft_id).delete_email(draft_id)
                log.info("[send-undo] Cancelled OK")
                wx.CallAfter(lambda: self._set_status("↩ Send cancelled"))
            except Exception as e:
                log.warning("[send-undo] Cancel failed: %s", e)
                wx.CallAfter(lambda: self._set_status("⚠ Could not cancel — may have already sent"))
        threading.Thread(target=run, daemon=True).start()

    def _send_reply(self):
        """Prepare reply data for undo-send flow."""
        if not self.selected_email_id: return
        body = self.reply_text.GetValue().strip()
        if not body:
            showinfo("Empty", "Type a reply first."); return
        eid, is_all = self.selected_email_id, self._is_reply_all

        # Parse extra CC recipients
        cc_text = self._cc_entry.GetValue().strip()
        extra_cc = [a.strip() for a in cc_text.replace(",", ";").split(";")
                    if a.strip() and "@" in a.strip()] if cc_text else None

        # Get original email for signature decision
        em = next((e for e in self.emails if e.get("id")==eid), None)

        # Build reply HTML — respect user's signature checkbox choice
        reply_html = body.replace("\n", "<br>")

        sig_html = ""
        if getattr(self, '_reply_include_sig', False):
            sig_html = self._build_signature_html()

        # For createReply/createReplyAll, only send the new content — Graph auto-quotes original
        reply_body_html = f"""<div style="font-family:Aptos,Aptos_MSFontService,-apple-system,Roboto,Arial,Helvetica,sans-serif;font-size:12pt;color:rgb(33,33,33)">
            <div>{reply_html}</div>
            {sig_html}
        </div>"""

        # Store for undo-send
        # Include edited subject/recipients if user modified them
        edited_subject = None
        edited_to = None
        if getattr(self, '_edit_subject_visible', False):
            new_subj = self._edit_subject_entry.GetValue().strip()
            orig_subj = em.get("subject", "") if em else ""
            if new_subj and new_subj != orig_subj:
                edited_subject = new_subj
        if getattr(self, '_edit_to_visible', False):
            new_to = self._edit_to_entry.GetValue().strip()
            if new_to:
                edited_to = [a.strip() for a in new_to.replace(",", ";").split(";")
                             if a.strip() and "@" in a.strip()]

        self._pending_reply_data = {
            "eid": eid, "is_all": is_all, "html": reply_body_html, "extra_cc": extra_cc,
            "edited_subject": edited_subject, "edited_to": edited_to,
        }

    def _send_forward(self):
        """Prepare forward data for undo-send flow."""
        if not self.selected_email_id:
            self._reset_send_btn(); return
        to_raw = self._fwd_to_entry.GetValue().strip()
        to_addrs = [a.strip() for a in to_raw.replace(",", ";").split(";")
                    if a.strip() and "@" in a.strip()]
        if not to_addrs:
            self._reset_send_btn()
            showinfo("Invalid Address", "Enter valid email address(es) to forward to.\nSeparate multiple with ;")
            self._fwd_to_entry.SetFocus()
            return
        comment = self.reply_text.GetValue().strip()
        eid = self.selected_email_id
        comment_html = comment.replace("\n", "<br>") if comment else ""

        # Include signature on forwards (always first time in forwarded thread)
        sig_html = self._build_signature_html()
        if sig_html:
            comment_html = f"{comment_html}{sig_html}" if comment_html else sig_html

        # Store for undo-send
        self._pending_fwd_data = {
            "eid": eid, "to_addrs": to_addrs, "comment_html": comment_html
        }

    def _reset_send_btn(self):
        """Re-enable the send button after error or cancel."""
        self._sending = False
        self._send_btn.SetLabel("📤 Send"); self._send_btn.Enable()

    def _cancel_reply(self):
        # Cancel any pending spell check
        if self._spell_timer:
            self._spell_timer.Stop()
            self._spell_timer = None
        self._spell_errors = []
        self._is_forward = False
        self._reset_send_btn()
        self._update_compose_btn_styles(None)  # Reset all to default
        # Only do pack operations if reply frame is currently visible
        if self.reply_frame.IsShown():
            self.reply_frame.Hide()
            self._fwd_to_frame.Hide()
            self._cc_frame.Hide()
            self._sig_preview_frame.Hide()
            self._edit_subject_frame.Hide()
            self._edit_to_frame.Hide()
            self._edit_subject_visible = False
            self._edit_to_visible = False
            self._reply_include_sig = False
            self.reply_text.Clear()
            self._fwd_to_entry.SetValue("") if hasattr(self, "_fwd_to_entry") else None
            self._cc_entry.SetValue("") if hasattr(self, "_cc_entry") else None
            self._body_container.Hide()
            self._body_container.Show()
            self._detail_panel.Layout()

    # ── Spell/Grammar Check ───────────────────────────────────

    def _on_reply_key(self, event=None):
        """Debounced spell check — triggers 800ms after last keystroke."""
        if self._spell_timer:
            self._spell_timer.Stop() if self._spell_timer else None
        self._spell_timer = wx.CallLater(800, self._run_spell_check)

    def _run_spell_check(self):
        """Run spell check in background thread."""
        text = self.reply_text.GetValue().strip()
        if not text:
            self._clear_spell_marks()
            return
        wx.CallAfter(lambda: self._set_status("Checking spelling..."))
        def run():
            errors = self.spell_checker.check(text)
            wx.CallAfter(lambda: self._apply_spell_marks(text, errors))
        threading.Thread(target=run, daemon=True).start()

    def _clear_spell_marks(self):
        """Remove all spell/grammar highlight styles."""
        self._spell_errors = []
        # Clear all styling on the rich text ctrl back to default
        try:
            length = self.reply_text.GetLastPosition()
            if length > 0:
                attr = wx.TextAttr()
                attr.SetBackgroundColour(self.reply_text.GetBackgroundColour())
                self.reply_text.SetStyle(0, length, attr)
        except Exception:
            pass

    def _apply_spell_marks(self, checked_text, errors):
        """Highlight misspelled/grammar error words with coloured backgrounds."""
        current = self.reply_text.GetValue().strip()
        if current != checked_text:
            return  # text changed since check started, discard stale results

        self._clear_spell_marks()
        self._spell_errors = errors

        if not errors:
            self._set_status("No spelling or grammar issues ✓")
            return

        # Highlight each error: yellow for spelling, orange for grammar
        for err in errors:
            start = err.get("offset", 0)
            end = start + err.get("length", 0)
            if end <= start:
                continue
            is_spell = any(r in err.get("rule", "")
                           for r in ["SPEL", "SPELL", "MORFOLOGIK", "HUNSPELL"])
            colour = wx.Colour(255, 255, 150) if is_spell else wx.Colour(255, 200, 120)
            try:
                attr = wx.TextAttr()
                attr.SetBackgroundColour(colour)
                self.reply_text.SetStyle(start, end, attr)
            except Exception:
                pass

        count_spell = sum(1 for e in errors if any(r in e.get("rule", "")
                          for r in ["SPEL", "SPELL", "MORFOLOGIK", "HUNSPELL"]))
        count_grammar = len(errors) - count_spell
        parts = []
        if count_spell: parts.append(f"{count_spell} spelling")
        if count_grammar: parts.append(f"{count_grammar} grammar")
        self._set_status(f"Found {' and '.join(parts)} issue{'s' if len(errors) > 1 else ''} — right-click to fix")

    def _click_to_char_offset(self, event):
        """Convert a mouse-click position to a character offset in the TextCtrl."""
        # wx.TextCtrl.HitTest(pt) returns (TextCtrlHitTestResult, col, row)
        result, col, row = self.reply_text.HitTest(event.GetPosition())
        return self.reply_text.XYToPosition(col, row)

    def _on_spell_right_click(self, event):
        """Show context menu with spelling/grammar suggestions."""
        click_offset = self._click_to_char_offset(event)

        clicked_error = None
        for err in self._spell_errors:
            start = err["offset"]
            end = err["offset"] + err["length"]
            if start <= click_offset <= end:
                clicked_error = err
                break

        menu = wx.Menu()

        if clicked_error:
            # Disabled header showing the error message
            hdr_item = menu.Append(wx.ID_ANY, f"⚠ {clicked_error['message'][:60]}")
            hdr_item.Enable(False)
            menu.AppendSeparator()

            if clicked_error["replacements"]:
                for suggestion in clicked_error["replacements"][:8]:
                    item = menu.Append(wx.ID_ANY, f"  ✓ {suggestion}")
                    menu.Bind(wx.EVT_MENU, lambda e, s=suggestion, err=clicked_error:
                              self._apply_suggestion(err, s), item)
            else:
                no_item = menu.Append(wx.ID_ANY, "  No suggestions")
                no_item.Enable(False)

            menu.AppendSeparator()

        _wx_menu_item(menu, "  🔧 Fix All Errors", self._fix_all_errors)

        self.reply_text.PopupMenu(menu)
        menu.Destroy()

    def _apply_suggestion(self, error, suggestion):
        """Replace an error with a suggestion using integer character offsets."""
        try:
            start = error["offset"]
            end = error["offset"] + error["length"]
            current_text = self.reply_text.GetValue()
            self.reply_text.SetValue(current_text[:start] + suggestion + current_text[end:])
            # Restore cursor to just after the replacement
            self.reply_text.SetInsertionPoint(start + len(suggestion))
            # Re-run spell check after fix
            if self._spell_timer:
                self._spell_timer.Stop()
            self._spell_timer = wx.CallLater(300, self._run_spell_check)
        except Exception:
            pass

    def _fix_all_errors(self):
        """Apply first suggestion for all errors at once."""
        text = self.reply_text.GetValue().strip()
        if not text:
            return
        def run():
            fixed = self.spell_checker.auto_fix(text)
            wx.CallAfter(lambda: self._apply_fixed_text(text, fixed))
        threading.Thread(target=run, daemon=True).start()

    def _apply_fixed_text(self, original, fixed):
        """Replace reply text with auto-corrected version."""
        current = self.reply_text.GetValue().strip()
        if current != original:
            return  # text changed while fixing
        if fixed == original:
            self._set_status("No fixes needed")
            return
        self.reply_text.Clear()
        self.reply_text.SetValue(fixed)
        self._clear_spell_marks()
        self._set_status("All errors fixed ✓")
        # Re-check to confirm
        if self._spell_timer:
            self._spell_timer.Stop() if self._spell_timer else None
        self._spell_timer = wx.CallLater(500, self._run_spell_check)

