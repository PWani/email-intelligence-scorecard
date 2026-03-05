import os, base64, json, re, sys, tempfile, threading, webbrowser
import wx
from ._wx_common import (
    _hex, _font, askstring, showerror, showinfo, askyesno,
    FONT, FONT_BOLD, CONFIG_DIR, C, P_ICON, log,
    GraphClient, OfflineQueue, is_network_error,
    html_to_readable_text, datetime, timedelta, timezone, unescape,
)


class AttachmentsMixin:
    """Attachments"""

    def _toggle_attachments(self):
        """Toggle the attachment panel. Fetch on first open."""
        if self._attach_visible:
            self._attach_frame.Hide()
            self._refresh_layout()
            self._attach_visible = False
            return
        eid = self.selected_email_id
        if not eid:
            return
        if eid in self._attach_cached:
            self._show_attachment_list(self._attach_cached[eid])
            return
        self._set_status("Loading attachments...")
        self._attach_inner_sizer.Clear(delete_windows=True)
        lbl = wx.StaticText(self._attach_frame, label="Loading attachments...")
        lbl.SetForegroundColour(_hex(C["muted"]))
        self._attach_inner_sizer.Add(lbl, 0, wx.ALL, 4)
        self._attach_frame.Layout()
        self._attach_frame.Show()
        self._refresh_layout()
        self._attach_visible = True

        def fetch():
            try:
                attachments = self.graph.get_attachments(eid)
                self._attach_cached[eid] = attachments
                if self.selected_email_id == eid:
                    wx.CallAfter(self._show_attachment_list, attachments)
            except Exception as e:
                err = str(e)
                if self.selected_email_id == eid:
                    wx.CallAfter(self._set_status, f"Attachment error: {err}")
        threading.Thread(target=fetch, daemon=True).start()

    def _show_attachment_list(self, attachments):
        """Render the attachment list in the panel."""
        self._attach_inner_sizer.Clear(delete_windows=True)
        if not attachments:
            lbl = wx.StaticText(self._attach_frame, label="No downloadable attachments")
            lbl.SetForegroundColour(_hex(C["muted"]))
            self._attach_inner_sizer.Add(lbl, 0, wx.ALL, 4)
        else:
            for att in attachments:
                name = att.get("name", "attachment")
                size = att.get("size", 0)
                if att.get("@odata.type", "") == "#microsoft.graph.itemAttachment":
                    continue
                if att.get("isInline", False) and not name.lower().endswith(
                        ('.pdf', '.docx', '.xlsx', '.pptx', '.zip', '.csv')):
                    continue
                size_str = f" ({self._fmt_size(size)})" if size else ""
                btn = wx.Button(self._attach_frame, label=f"📄 {name}{size_str}")
                btn.SetBackgroundColour(_hex("#E0E8FF"))
                btn.SetForegroundColour(_hex("#1E40AF"))
                btn.Bind(wx.EVT_BUTTON, lambda e, a=att: self._open_attachment(a))
                self._attach_inner_sizer.Add(btn, 0, wx.LEFT | wx.BOTTOM, 4)
        self._attach_frame.Layout()
        self._attach_frame.Show()
        self._refresh_layout()
        self._attach_visible = True
        self._set_status(f"{len(attachments)} attachment(s)")

    def _open_attachment(self, att):
        """Download attachment to temp dir and open with system default app."""
        name = att.get("name", "attachment")
        content_bytes = att.get("contentBytes")
        if not content_bytes:
            showinfo("Attachment", f"Cannot download '{name}' — no content available.")
            return
        self._set_status(f"Opening {name}...")
        def run():
            try:
                data = base64.b64decode(content_bytes)
                tmp_dir = os.path.join(tempfile.gettempdir(), "email_dashboard_attachments")
                os.makedirs(tmp_dir, exist_ok=True)
                filepath = os.path.join(tmp_dir, name)
                if os.path.exists(filepath):
                    base, ext = os.path.splitext(name)
                    i = 1
                    while os.path.exists(filepath):
                        filepath = os.path.join(tmp_dir, f"{base}_{i}{ext}")
                        i += 1
                with open(filepath, "wb") as f:
                    f.write(data)
                os.startfile(filepath)
                wx.CallAfter(self._set_status, f"Opened {name}")
            except Exception as e:
                err = str(e)
                wx.CallAfter(showerror, "Attachment Error", err)
        threading.Thread(target=run, daemon=True).start()

    def _fmt_size(self, size_bytes):
        """Format file size for display."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.0f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.1f} MB"

    # ── Email Signature ────────────────────────────────────────
