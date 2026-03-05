import os, re, threading, webbrowser
from html import unescape

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
    has_remote_images,
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


def _wx_menu_item(menu: wx.Menu, label: str, handler) -> wx.MenuItem:
    """Append a labelled item to a wx.Menu and bind it to handler."""
    item = menu.Append(wx.ID_ANY, label)
    menu.Bind(wx.EVT_MENU, lambda e: handler(), item)
    return item


class DetailViewMixin:
    """Detail View and Rendering"""

    def _select(self, email_id, auto_scroll=True):
        self.selected_email_id = email_id
        em = next((e for e in self.emails if e.get('id') == email_id), None)
        subj = (em.get('subject') or '')[:50] if em else '?'
        prov = em.get('_provider', 'ms') if em else '?'
        log.info('[select] subj=%r provider=%s eid=%s', subj, prov, (email_id or '')[:40])
        self._cancel_reply()
        self._highlight_card(email_id)
        if auto_scroll:
            self._scroll_to_email(email_id)
        if not em:
            return

        intel = em.get('_intel', {})
        self.d_subject.SetLabel(em.get('subject') or '(no subject)')
        sn = em.get('from', {}).get('emailAddress', {}).get('name', 'Unknown')
        sa = em.get('from', {}).get('emailAddress', {}).get('address', '')
        provider_tag = ' [Google]' if em.get('_provider') == 'google' else ''
        self.d_from.SetLabel(f'From: {sn} <{sa}>{provider_tag}')
        self.d_date.SetLabel(self._fmt_date_full(em.get('receivedDateTime', '')))

        # Compact To/CC line
        to_names = [r.get('emailAddress', {}).get('name') or r.get('emailAddress', {}).get('address', '')
                    for r in em.get('toRecipients', [])]
        cc_names = [r.get('emailAddress', {}).get('name') or r.get('emailAddress', {}).get('address', '')
                    for r in em.get('ccRecipients', [])]
        to_str = 'To: ' + ', '.join(to_names[:4])
        if len(to_names) > 4:
            to_str += f' +{len(to_names) - 4} more'
        if cc_names:
            cc_part = ', '.join(cc_names[:3])
            if len(cc_names) > 3:
                cc_part += f' +{len(cc_names) - 3}'
            to_str += f'  |  CC: {cc_part}'
        self.d_to.SetLabel(to_str)
        self.d_to.Show()

        pri = intel.get('priority', 'normal')
        flagged = em.get('flag', {}).get('flagStatus') == 'flagged'
        pri_text = pri.upper() + ('  🚩' if flagged else '')
        self.d_priority.SetLabel(pri_text)
        pri_colour = (C['red'] if pri == 'urgent' else C['orange'] if pri == 'important'
                      else C['blue'] if pri == 'normal' else C['muted'])
        self.d_priority.SetForegroundColour(_hex(pri_colour))
        self.d_score.SetLabel(f"Score: {intel.get('score', 0)}/100")
        self.d_signals.SetLabel('  •  '.join(intel.get('signals', [])[:5]))
        # Force sizer to recompute widths after label text changes
        self.d_priority.InvalidateBestSize()
        self.d_score.InvalidateBestSize()
        self.d_signals.InvalidateBestSize()
        self._detail_panel.Layout()

        # Attachment button visibility
        if em.get('hasAttachments'):
            self._attach_btn.Show()
        else:
            self._attach_btn.Hide()

        self._update_action_buttons_for_folder(em)
        self._attach_frame.Hide()
        self._attach_visible = False
        self._load_images_frame.Hide()
        self._event_time_frame.Hide()

        # Meeting request detection
        if '_is_meeting_request' in em:
            self._apply_meeting_button(em['_is_meeting_request'])
            if em.get('_event_times'):
                self._show_event_time(em['_event_times'])
            elif em['_is_meeting_request'] == 'request':
                def _fetch_times():
                    try:
                        event_times = self._api_for(email_id).get_event_times(email_id)
                        em['_event_times'] = event_times
                        if self.selected_email_id == email_id:
                            wx.CallAfter(self._show_event_time, event_times)
                    except Exception:
                        pass
                threading.Thread(target=_fetch_times, daemon=True).start()
        else:
            self._apply_meeting_button(False)
            def _check_meeting():
                try:
                    _c = self._api_for(email_id)
                    result = _c.is_meeting_request(email_id)
                    em['_is_meeting_request'] = result
                    if result:
                        event_times = _c.get_event_times(email_id)
                        em['_event_times'] = event_times
                        if self.selected_email_id == email_id:
                            wx.CallAfter(self._apply_meeting_button, result)
                            wx.CallAfter(self._show_event_time, event_times)
                except Exception as exc:
                    em['_is_meeting_request'] = False
                    log.warning('[check_meeting] %s → ERROR: %s', email_id[:40], exc)
            threading.Thread(target=_check_meeting, daemon=True).start()

        body = em.get('body', {})
        raw_content = body.get('content', '') or em.get('bodyPreview', '')
        content_type = body.get('contentType', 'text')

        block_imgs = not self._is_safe_image_sender(em)
        provider = em.get('_provider', 'ms')
        self._render_email_body(raw_content, content_type,
                                block_images=block_imgs, provider=provider)

        if not em.get('isRead', True):
            threading.Thread(target=lambda: self._mark_read_bg(email_id), daemon=True).start()

    # ── Rendering ─────────────────────────────────────────────

    def _resolve_cid_images(self, html, email_id, provider):
        """Replace cid: references with base64 data URIs."""
        import re as _re
        if 'cid:' not in html.lower():
            return html
        try:
            client = self._api_for(email_id) if email_id else None
            if not client:
                return html
            attachments = client.get_attachments(email_id)
            cid_map = {}
            for att in attachments:
                cid = att.get('contentId', '').strip('<>')
                cb = att.get('contentBytes', '')
                ct = att.get('contentType', 'image/png')
                if cid and cb:
                    cid_map[cid.lower()] = 'data:' + ct + ';base64,' + cb
            if not cid_map:
                return html
            def _replace_cid(m):
                cid_val = m.group(1).strip('<>').lower()
                uri = cid_map.get(cid_val)
                return ('src="' + uri + '"') if uri else m.group(0)
            pat = _re.compile(r'src=["\']cid:([^"\']+)["\']', _re.IGNORECASE)
            html = pat.sub(_replace_cid, html)
        except Exception as e:
            log.debug('[cid] resolve failed: %s', e)
        return html

    def _render_email_body(self, raw_content, content_type,
                           block_images=True, provider='ms'):
        """Render email body via wx.html2.WebView."""
        log.info('[render] type=%s block=%s len=%d provider=%s',
                 content_type, block_images, len(raw_content or ''), provider)

        # Store for "Load Images" re-render
        self._last_raw_content = raw_content
        self._last_content_type = content_type

        if content_type == 'html':
            html = raw_content
            if '<html' not in html.lower():
                html = f"<html><head><meta charset='utf-8'></head><body>{html}</body></html>"
            # Resolve CID inline images (email_id may be None for previews)
            eid = getattr(self, 'selected_email_id', None)
            if eid and 'cid:' in html.lower():
                html = self._resolve_cid_images(html, eid, provider)
            # Google: strip junk that crashes even good renderers
            if provider == 'google':
                try:
                    html = self._sanitize_google_html(html,
                                                      block_remote_images=block_images)
                except Exception as e:
                    log.warning('[render] google sanitize failed: %s', e)
        else:
            # Plain text → HTML with clickable links
            url_pat = re.compile(r'(https?://\S+)')
            parts = url_pat.split(raw_content)
            html_parts = []
            for i, part in enumerate(parts):
                if i % 2 == 1:
                    display = part.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    href = part.replace('&', '&amp;').replace('"', '&quot;')
                    html_parts.append(
                        f'<a href="{href}" style="color:#2563EB;word-break:break-all;">{display}</a>')
                else:
                    html_parts.append(
                        part.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))
            escaped = ''.join(html_parts)
            html = (f"<html><head><meta charset='utf-8'>"
                    f"<style>body{{font-family:'Aptos','Segoe UI',sans-serif;"
                    f"font-size:14px;white-space:pre-wrap;line-height:1.6;padding:4px 6px;}}"
                    f"</style></head><body>{escaped}</body></html>")

        # Show "Load Images" bar if remote images are present and blocked
        has_remote = has_remote_images(html)
        if has_remote and block_images:
            self._load_images_frame.Show()
            self._detail_panel.Layout()
        else:
            self._load_images_frame.Hide()

        # Hand off to EmailWebView — Python-side blocking already applied
        try:
            log.info('[render] load_html %d chars', len(html))
            self.body.load_html(html, block_images=block_images)
            log.info('[render] load_html OK')
        except Exception as e:
            log.warning('[render] load_html failed: %s', e)
            self._render_plain_fallback(raw_content, content_type)

    def _render_plain_fallback(self, raw_content, content_type):
        """Render as plain text inside the webview."""
        txt = html_to_readable_text(raw_content) if content_type == 'html' else raw_content
        escaped = (txt or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        simple = ('<html><body><pre style="white-space:pre-wrap;font-family:sans-serif;'
                  'font-size:14px;padding:4px 6px;">' + escaped + '</pre></body></html>')
        try:
            self.body.load_html(simple, block_images=False)
        except Exception:
            pass

    # ── Copy / clipboard ──────────────────────────────────────

    def _copy_body_text(self, event=None):
        """Copy selected webview text, or full body text if nothing selected."""
        try:
            sel = self.body.get_selected_text()
            if not sel:
                raw = getattr(self, '_last_raw_content', '')
                if not raw:
                    return
                sel = re.sub(r'<[^>]+>', '', raw)
                sel = unescape(sel)
                sel = re.sub(r'\n{3,}', '\n\n', sel).strip()
            if sel and wx.TheClipboard.Open():
                wx.TheClipboard.SetData(wx.TextDataObject(sel.strip()))
                wx.TheClipboard.Close()
                self._set_status('📋 Copied to clipboard')
        except Exception:
            pass

    def _body_right_click(self, event=None):
        """Context menu on body right-click."""
        menu = wx.Menu()
        _wx_menu_item(menu, '📋 Copy Text', self._copy_body_text)
        link = getattr(self, '_last_hovered_link', None)
        if link:
            _wx_menu_item(menu, '🔗 Copy Link',
                          lambda: self._clipboard_set(link) or self._set_status('📋 Link copied'))
            _wx_menu_item(menu, '🌐 Open in Browser', lambda: webbrowser.open(link))
        self.root.PopupMenu(menu)
        menu.Destroy()

    def _clipboard_set(self, text: str):
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(text))
            wx.TheClipboard.Close()

    # ── HTML sanitisation (Google emails) ─────────────────────

    def _sanitize_google_html(self, html, block_remote_images=True):
        """Strip dangerous tags from Google email HTML using regex only.
        Never uses BS4 — BS4 encodes & in src= URLs breaking images with query params.
        """
        # Remove dangerous/noisy tags and their content
        for tag in ['script', 'iframe', 'object', 'embed', 'video', 'audio',
                    'canvas', 'noscript', 'form']:
            html = re.sub(rf'<{tag}[\s>].*?</{tag}>', '', html,
                          flags=re.IGNORECASE | re.DOTALL)
            html = re.sub(rf'<{tag}[^>]*/>', '', html, flags=re.IGNORECASE)

        # Remove <meta> and <link> tags (self-closing)
        html = re.sub(r'<meta\b[^>]*>', '', html, flags=re.IGNORECASE)
        html = re.sub(r'<link\b[^>]*>', '', html, flags=re.IGNORECASE)

        # Strip CSS background-image references
        html = re.sub(r'background-image\s*:\s*url\s*\([^)]*\)', 'background-image:none',
                      html, flags=re.IGNORECASE)

        # Remove img tags entirely if blocking
        if block_remote_images:
            html = re.sub(r'<img\b[^>]*>', '', html, flags=re.IGNORECASE | re.DOTALL)

        # Wrap in minimal chrome if no <html> tag
        if '<html' not in html.lower():
            html = (
                '<html><head><style>'
                'body{font-family:Segoe UI,Arial,sans-serif;font-size:14px;'
                'color:#1a1a1a;line-height:1.5;padding:4px 6px;word-wrap:break-word}'
                'a{color:#0066cc}img{max-width:100%}'
                '</style></head><body>' + html + '</body></html>'
            )

        return html

    def _sanitize_html_for_display(self, html, block_remote_images=True):
        """Minimal sanitisation for MS emails — mostly handled by webview JS."""
        html = re.sub(r'<script[^>]*>.*?</script>', '', html,
                      flags=re.IGNORECASE | re.DOTALL)
        html = re.sub(r'<iframe[^>]*>.*?</iframe>', '', html,
                      flags=re.IGNORECASE | re.DOTALL)
        html = re.sub(r'<meta[^>]*http-equiv[^>]*refresh[^>]*>', '', html,
                      flags=re.IGNORECASE)
        return html

    # ── Mark-read ─────────────────────────────────────────────

    def _mark_read_bg(self, eid):
        log.info('[mark-read-bg] eid=%s', eid[:40] if eid else 'None')
        try:
            self._api_for(eid).mark_as_read(eid)
            log.info('[mark-read-bg] OK: eid=%s', eid[:40] if eid else 'None')
        except Exception as e:
            err_str = str(e)
            if '404' in err_str:
                # Email no longer exists on server (deleted/moved) — silently skip
                log.debug('[mark-read-bg] 404 skip eid=%s', eid[:40] if eid else 'None')
            else:
                log.error('[mark-read-bg] FAILED: eid=%s err=%s',
                          eid[:40] if eid else 'None', e)
                if is_network_error(err_str):
                    self._offline_queue.enqueue('mark_read', eid=eid)
        for e in self.emails:
            if e.get('id') == eid:
                e['isRead'] = True
                break
        wx.CallAfter(self._update_card_read, eid)
        wx.CallAfter(self._update_stats)

    def _update_card_read(self, eid):
        """Update card styling to reflect read status."""
        if eid not in self._card_refs:
            return
        card = self._card_refs[eid][0]
        try:
            # Walk card children to find subject label (first StaticText in row 1)
            content = card.GetChildren()[0]
            r1 = content.GetChildren()[0]
            for widget in r1.GetChildren():
                if isinstance(widget, wx.StaticText):
                    widget.SetFont(_font(FONT, 10))
                    widget.SetForegroundColour(_hex(C['text2']))
                    break
        except Exception:
            pass

    # ── Actions separator ─────────────────────────────────────
