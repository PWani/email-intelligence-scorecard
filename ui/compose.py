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

