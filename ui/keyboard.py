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


class KeyboardMixin:
    """Keyboard Navigation"""

    def _highlight_card(self, email_id):
        """Visually highlight the selected card with a left accent bar."""
        ACCENT_COL = '#2563EB'   # blue accent strip
        SEL_BG     = '#EFF6FF'   # very light blue tint for selected card bg

        prev = getattr(self, '_prev_highlighted', None)
        if prev and prev != email_id and prev in self._card_refs:
            entry = self._card_refs[prev]
            panel, orig_bg = entry[0], entry[1]
            accent = entry[2] if len(entry) > 2 else None
            content = entry[3] if len(entry) > 3 else None
            try:
                panel.SetBackgroundColour(orig_bg)
                if accent:
                    accent.SetBackgroundColour(orig_bg)
                if content:
                    content.SetBackgroundColour(orig_bg)
                    for child in content.GetChildren():
                        child.SetBackgroundColour(orig_bg)
                panel.Refresh()
            except Exception:
                pass

        if email_id in self._card_refs:
            entry = self._card_refs[email_id]
            panel, orig_bg = entry[0], entry[1]
            accent = entry[2] if len(entry) > 2 else None
            content = entry[3] if len(entry) > 3 else None
            try:
                panel.SetBackgroundColour(SEL_BG)
                if accent:
                    accent.SetBackgroundColour(ACCENT_COL)
                if content:
                    content.SetBackgroundColour(SEL_BG)
                    for child in content.GetChildren():
                        child.SetBackgroundColour(SEL_BG)
                panel.Refresh()
            except Exception:
                pass
        self._prev_highlighted = email_id

    def _recolor_children(self, widget, bg_colour):
        """Not needed in wx — panels refresh their own background."""
        pass

    def _set_focus(self, pane):
        self._focus_pane = pane
        # Move wx focus to root so key bindings work
        try:
            self.root.SetFocus()
        except Exception:
            pass

    def _on_window_focus(self, event=None):
        focused = wx.Window.FindFocus()
        if focused and isinstance(focused, wx.TextCtrl):
            return
        try:
            self.root.SetFocus()
        except Exception:
            pass

    def _is_typing(self, event=None):
        focused = wx.Window.FindFocus()
        if focused is None:
            return False
        # TextCtrl = typing
        if isinstance(focused, wx.TextCtrl):
            return True
        # SearchCtrl inherits from wx.Control — check class name
        cls = type(focused).__name__
        if 'SearchCtrl' in cls or 'TextCtrl' in cls:
            return True
        # Autocomplete listbox: ListBox focused means user is in address entry context
        if isinstance(focused, wx.ListBox):
            return True
        # Any widget inside reply_frame = composing
        try:
            parent = focused.GetParent()
            while parent:
                if parent is getattr(self, 'reply_frame', None):
                    return True
                parent = parent.GetParent()
        except Exception:
            pass
        return False

    def _dismiss_popup(self):
        """Dismiss any active popup menu (snooze/remind)."""
        if self._active_popup_menu:
            try:
                if hasattr(self._active_popup_menu, "Destroy"): self._active_popup_menu.Destroy()
            except Exception:
                pass
            self._active_popup_menu = None

    def _on_key_archive(self, event=None):
        if not self._is_typing(event):
            log.info("[key] archive pressed")
            self._dismiss_popup()
            # Guard: ignore keypress if async meeting check already in flight
            if getattr(self, '_archive_checking', False):
                return
            if not self._current_action_is_accept and self.selected_email_id:
                em = next((e for e in self.emails if e.get("id") == self.selected_email_id), None)
                if em and "_is_meeting_request" not in em:
                    # Meeting status unknown — run quick check before acting
                    self._archive_checking = True
                    self._set_status("Checking email type...")
                    def check_then_act():
                        try:
                            _c = self._api_for(self.selected_email_id)
                            is_meeting = _c.is_meeting_request(self.selected_email_id)
                            em["_is_meeting_request"] = is_meeting
                            if is_meeting:
                                em["_event_times"] = _c.get_event_times(self.selected_email_id)
                            wx.CallAfter(lambda: (
                                self._apply_meeting_button(is_meeting),
                                self._on_archive_accept()))
                        except Exception:
                            wx.CallAfter(self._on_archive_accept)
                        finally:
                            self._archive_checking = False
                    threading.Thread(target=check_then_act, daemon=True).start()
                    return
            self._on_archive_accept()

    def _on_key_reply(self, event=None):
        if not self._is_typing(event):
            log.info("[key] reply pressed")
            self._dismiss_popup()
            self._reply()

    def _on_key_reply_all(self, event=None):
        if not self._is_typing(event):
            self._dismiss_popup()
            self._reply_all()

    def _on_key_forward(self, event=None):
        if not self._is_typing(event):
            log.info("[key] forward pressed")
            self._dismiss_popup()
            self._forward()

    def _on_key_delete(self, event=None):
        if not self._is_typing(event):
            log.info("[key] delete pressed")
            self._dismiss_popup()
            self._delete()

    def _on_arrow_up(self, event=None):
        if self._is_typing(event):
            return
        if self._focus_pane == "list":
            self._select_adjacent(-1)
        else:
            self._scroll_body("up")

    def _on_arrow_down(self, event=None):
        if self._is_typing(event):
            return
        if self._focus_pane == "list":
            self._select_adjacent(1)
        else:
            self._scroll_body("down")

    def _select_adjacent(self, direction):
        """Select the next (+1) or previous (-1) email in visual sizer order.

        Uses list_inner_sizer item order — this is the true on-screen order,
        which accounts for both MS and Google emails inserted at sorted positions.
        _card_refs insertion order is wrong because Google emails are appended
        after MS emails regardless of their visual position.
        """
        # Build ordered id list from sizer — reflects exact visual order
        panel_to_id = getattr(self, '_panel_to_id', {})
        sizer = self.list_inner_sizer
        rendered_ids = []
        for i in range(sizer.GetItemCount()):
            item = sizer.GetItem(i)
            if item and item.GetWindow():
                eid = panel_to_id.get(id(item.GetWindow()))
                if eid:
                    rendered_ids.append(eid)

        if not rendered_ids:
            return
        if not self.selected_email_id or self.selected_email_id not in rendered_ids:
            self._select(rendered_ids[0])
            return

        i = rendered_ids.index(self.selected_email_id)
        new_idx = i + direction
        if 0 <= new_idx < len(rendered_ids):
            new_id = rendered_ids[new_idx]
            self._select(new_id)
            if new_idx >= len(rendered_ids) - 3 and not self._loading_more:
                self._auto_load_more()

    def _scroll_to_email(self, email_id, email_idx=None):
        """Scroll the list canvas so the card for email_id is visible."""
        wx.CallAfter(lambda: self._do_scroll_to_email(email_id, email_idx))

    def _do_scroll_to_email(self, email_id, email_idx=None):
        if email_id not in self._card_refs:
            return
        panel = self._card_refs[email_id][0]
        try:
            _, scroll_unit = self._list_scroll.GetScrollPixelsPerUnit()
            scroll_unit = max(scroll_unit, 1)

            # Convert the panel's screen rect into the scroll window's client coords.
            # This is the only coordinate that is always correct — screen position is
            # absolute and doesn't depend on layout state or scroll offset.
            panel_screen_y = panel.GetScreenPosition().y
            scroll_screen_y = self._list_scroll.GetScreenPosition().y
            # client_y = where the card top appears in the visible viewport right now
            client_y = panel_screen_y - scroll_screen_y
            card_h = panel.GetSize().height
            view_h = self._list_scroll.GetClientSize().height
            scroll_y_px = self._list_scroll.GetScrollPos(wx.VERTICAL) * scroll_unit
            # virtual_y = client_y + current scroll offset
            virtual_y = client_y + scroll_y_px
            margin = 8

            if client_y < margin:
                # Card is above visible area
                new_pos = max(0, (virtual_y - margin) // scroll_unit)
                self._list_scroll.Scroll(-1, new_pos)
            elif client_y + card_h > view_h - margin:
                # Card bottom is below visible area
                new_pos = max(0, (virtual_y + card_h - view_h + margin) // scroll_unit)
                self._list_scroll.Scroll(-1, new_pos)
        except Exception:
            pass

        if email_idx is not None and email_idx >= len(self.emails) - 3 and not self._loading_more:
            self._auto_load_more()

    def _scroll_body(self, direction):
        """Scroll the email body reading pane via webview JS."""
        lines = -3 if direction == 'up' else 3
        try:
            self.body.scroll(lines)
        except Exception:
            pass

    # ── Detail Selection ──────────────────────────────────────

