"""
Shared wx imports and helpers used by every mixin.
Import this instead of duplicating the boilerplate in every file.
"""
import os, base64, json, re, sys, tempfile, threading, webbrowser
import wx
import wx.adv
from datetime import datetime, timedelta, timezone
from html import unescape
import msal, requests

from ..core.config import (
    FONT, FONT_BOLD, _detect_font,
    CONFIG_DIR, CONFIG_FILE, TOKEN_CACHE_FILE, ADDRESS_BOOK_FILE,
    OFFLINE_QUEUE_FILE, SCORING_RULES_FILE,
    _APP_ICON_B64, DEFAULT_CONFIG,
    ensure_config_dir, load_config, save_config,
    load_address_book_cache, save_address_book_cache,
    C, P_ICON, CAT_ICON, TIMEZONES, DEFAULT_TZ, log,
)
from ..core.auth import OutlookAuth
from ..core.graph_client import GraphClient, OfflineQueue, is_network_error
from ..core.email_intelligence import (
    DEFAULT_SCORING_RULES,
    load_scoring_rules, save_scoring_rules, _deep_merge,
    strip_html, strip_outlook_banners,
    _build_word_pattern, keyword_in_text,
    html_to_readable_text, extract_text, extract_latest_reply,
    EmailIntelligence,
)
from ..core.spell_checker import SpellChecker
from .autocomplete import EmailAutocomplete
from .webview_widget import (detect_renderer, EmailWebView, RENDERER,  # detect_renderer kept for compat
                             block_remote_images, unblock_remote_images, has_remote_images)

try:
    from ..core.google_client import (
        GoogleAuth, GmailClient, GoogleCalendarClient,
        is_google_available, get_google_import_error, GOOGLE_CREDS_FILE,
    )
    _HAS_GOOGLE_MODULE = True
except ImportError:
    _HAS_GOOGLE_MODULE = False


# ── Colour/font helpers ────────────────────────────────────────

_NAMED_COLOURS = {
    "white": "#FFFFFF", "black": "#000000", "red": "#FF0000", "green": "#00AA00",
    "blue": "#0000FF", "gray": "#808080", "grey": "#808080", "transparent": "#000000",
}

def _hex(h: str) -> wx.Colour:
    h = _NAMED_COLOURS.get(h.lower(), h)
    h = h.lstrip('#')
    return wx.Colour(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _font(face=None, size=10, bold=False) -> wx.Font:
    face = face or FONT
    w = wx.FONTWEIGHT_BOLD if bold else wx.FONTWEIGHT_NORMAL
    return wx.Font(size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, w, faceName=face)


# ── Simple dialog replacements ─────────────────────────────────

def askstring(title: str, prompt: str, parent=None) -> str | None:
    dlg = wx.TextEntryDialog(parent, prompt, caption=title)
    if dlg.ShowModal() == wx.ID_OK:
        val = dlg.GetValue().strip()
        dlg.Destroy()
        return val or None
    dlg.Destroy()
    return None


def showerror(title: str, msg: str, parent=None):
    wx.MessageBox(msg, title, wx.OK | wx.ICON_ERROR, parent)


def showinfo(title: str, msg: str, parent=None):
    wx.MessageBox(msg, title, wx.OK | wx.ICON_INFORMATION, parent)


def askyesno(title: str, msg: str, parent=None) -> bool:
    return wx.MessageBox(msg, title, wx.YES_NO | wx.ICON_QUESTION, parent) == wx.YES


def _wx_menu_item(menu: wx.Menu, label: str, handler) -> wx.MenuItem:
    """Append a labelled item to a wx.Menu and bind it to handler."""
    item = menu.Append(wx.ID_ANY, label)
    menu.Bind(wx.EVT_MENU, lambda e: handler(), item)
    return item
