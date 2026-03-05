# Email autocomplete — wx version
import wx
from ..core.config import FONT


class EmailAutocomplete:
    """
    Attach autocomplete dropdown to a wx.TextCtrl for email addresses.
    Supports semicolon-separated multiple addresses — only completes the last segment.
    """

    def __init__(self, entry: wx.TextCtrl, get_contacts_fn, max_results=8):
        self.entry = entry
        self.get_contacts = get_contacts_fn  # callable -> [{name, email}, ...]
        self.max_results = max_results
        self._popup = None
        self._listbox = None
        self._matches = []
        self._sel_index = 0

        entry.Bind(wx.EVT_TEXT, self._on_text)
        entry.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
        entry.Bind(wx.EVT_KILL_FOCUS, self._on_kill_focus)

    def _on_kill_focus(self, event):
        event.Skip()  # Critical: let focus transition complete so caret renders
        wx.CallLater(150, self._close)

    # ── Internal helpers ──────────────────────────────────────

    def _get_last_segment(self):
        text = self.entry.GetValue()
        cursor = self.entry.GetInsertionPoint()
        before = text[:cursor]
        sep_idx = max(before.rfind(';'), before.rfind(','))
        if sep_idx >= 0:
            return before[sep_idx + 1:].strip(), sep_idx + 1
        return before.strip(), 0

    def _replace_last_segment(self, replacement):
        text = self.entry.GetValue()
        cursor = self.entry.GetInsertionPoint()
        before = text[:cursor]
        after = text[cursor:]
        sep_idx = max(before.rfind(';'), before.rfind(','))
        if sep_idx >= 0:
            new_text = before[:sep_idx + 1] + ' ' + replacement + '; ' + after.lstrip(' ;,')
        else:
            new_text = replacement + '; ' + after.lstrip(' ;,')
        final = new_text.rstrip('; ') + ('; ' if after.strip() else '')
        self.entry.ChangeValue(final)
        self.entry.SetInsertionPointEnd()

    # ── Event handlers ────────────────────────────────────────

    def _on_text(self, event):
        event.Skip()
        query, _ = self._get_last_segment()
        if len(query) < 2:
            self._close()
            return
        self._search(query)

    def _on_key_down(self, event):
        key = event.GetKeyCode()
        if self._popup and self._matches:
            if key == wx.WXK_DOWN:
                self._sel_index = min(self._sel_index + 1, len(self._matches) - 1)
                self._listbox.SetSelection(self._sel_index)
                return
            if key == wx.WXK_UP:
                self._sel_index = max(self._sel_index - 1, 0)
                self._listbox.SetSelection(self._sel_index)
                return
            if key in (wx.WXK_RETURN, wx.WXK_TAB):
                self._select_current()
                return
            if key == wx.WXK_ESCAPE:
                self._close()
                return
        event.Skip()

    def _search(self, query):
        contacts = self.get_contacts()
        if not contacts:
            self._close()
            return
        q = query.lower()
        matches = []
        for c in contacts:
            if q in c.get('name', '').lower() or q in c.get('email', '').lower():
                matches.append(c)
            if len(matches) >= self.max_results:
                break
        if not matches:
            self._close()
            return
        self._matches = matches
        self._sel_index = 0
        self._show_popup()

    def _show_popup(self):
        items = []
        for c in self._matches:
            name = c.get('name', '')
            email = c.get('email', '')
            items.append(f'{name}  <{email}>' if name else email)

        if self._popup is None:
            parent = self.entry.GetTopLevelParent()
            self._popup = wx.PopupWindow(parent)
            self._listbox = wx.ListBox(self._popup, style=wx.LB_SINGLE)
            self._listbox.Bind(wx.EVT_LEFT_UP, self._on_click)
            self._listbox.Bind(wx.EVT_MOTION, self._on_motion)
        else:
            self._listbox.Clear()

        for item in items:
            self._listbox.Append(item)
        self._listbox.SetSelection(0)

        # Position below entry
        pos = self.entry.ClientToScreen(wx.Point(0, self.entry.GetSize().height))
        w = self.entry.GetSize().width
        h = min(len(self._matches) * 22 + 4, 180)
        self._popup.SetSize(w, h)
        self._popup.SetPosition(pos)
        self._listbox.SetSize(w, h)
        self._popup.Show()

    def _on_click(self, event):
        idx = self._listbox.GetSelection()
        if 0 <= idx < len(self._matches):
            self._sel_index = idx
            self._select_current()

    def _on_motion(self, event):
        # Highlight item under mouse
        idx = self._listbox.HitTest(event.GetPosition())
        if idx != wx.NOT_FOUND:
            self._listbox.SetSelection(idx)
            self._sel_index = idx
        event.Skip()

    def _select_current(self):
        if 0 <= self._sel_index < len(self._matches):
            c = self._matches[self._sel_index]
            self._replace_last_segment(c['email'])
        self._close()

    def _close(self, event=None):
        if self._popup:
            self._popup.Hide()
            self._popup.Destroy()
            self._popup = None
            self._listbox = None
        self._matches = []
