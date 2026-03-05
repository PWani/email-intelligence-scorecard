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
from .webview_widget import EmailWebView, RENDERER
try:
    from ..core.google_client import (
        GoogleAuth, GmailClient, GoogleCalendarClient,
        is_google_available, get_google_import_error, GOOGLE_CREDS_FILE,
    )
    _HAS_GOOGLE_MODULE = True
except ImportError:
    _HAS_GOOGLE_MODULE = False


# ── Colour helpers ─────────────────────────────────────────────

def _hex(h):
    """Convert '#RRGGBB' to wx.Colour."""
    h = h.lstrip('#')
    return wx.Colour(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _font(face=None, size=10, bold=False):
    face = face or FONT
    weight = wx.FONTWEIGHT_BOLD if bold else wx.FONTWEIGHT_NORMAL
    return wx.Font(size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, weight,
                   faceName=face)


# ── Thin wx wrappers that mimic common tk patterns ─────────────

def _label(parent, text='', fg=None, bg=None, font=None, **kw):
    lbl = wx.StaticText(parent, label=text)
    if fg:
        lbl.SetForegroundColour(_hex(fg))
    if bg:
        lbl.SetBackgroundColour(_hex(bg))
    if font:
        lbl.SetFont(font)
    return lbl


def _btn(parent, text, handler=None, fg=None, bg=None, font=None, **kw):
    b = wx.Button(parent, label=text)
    if handler:
        b.Bind(wx.EVT_BUTTON, lambda e: handler())
    if fg:
        b.SetForegroundColour(_hex(fg))
    if bg:
        b.SetBackgroundColour(_hex(bg))
    if font:
        b.SetFont(font)
    return b


class BuildMixin:
    """UI Construction — wx version"""

    def __init__(self):
        self.config = load_config()
        self.auth = None
        self.graph = None
        self.intelligence = None
        self.google_auth = None
        self.google_client = None
        self._google_email = ''
        self._google_name = ''
        self._google_emails = []
        self._ms_email = ''
        self._ms_name = ''
        self._account_filter = 'all'
        self.spell_checker = SpellChecker()
        self._spell_timer = None
        self._spell_errors = []
        self._offline_queue = OfflineQueue()
        self._folder_email_cache = {}
        self.emails = []
        self.current_skip = 0
        self.selected_email_id = None
        self.current_folder = 'inbox'
        self.search_query = None
        self.sort_by_priority = False
        self._is_reply_all = False
        self._address_book = []
        self._todays_events = []
        self._alerted_events = set()
        self._card_refs = {}   # email_id -> (panel, orig_bg)
        self._panel_to_id = {}  # id(panel) -> email_id  (for visual-order navigation)
        self._loading_more = False
        self._initial_load_target = 250
        self._undo_bar = None
        self._undo_active_draft = None
        self._pending_reply_data = None
        self._pending_fwd_data = None
        self._split_mode = 'all'
        self._user_profile = {}
        self._active_popup_menu = None
        self._focus_pane = 'list'
        self._list_width = 620
        self._prev_highlighted = None
        self._use_html_view = True
        self._last_raw_content = ''
        self._last_content_type = 'text'
        self._attach_visible = False
        self._attach_cached = {}

        # ── wx App + Frame ─────────────────────────────────────
        self._wx_app = wx.App()
        self.root = wx.Frame(None, title='Email Intelligence Dashboard v0.35',
                             size=(1440, 920))
        self.root.SetMinSize((1280, 700))
        self.root.SetBackgroundColour(_hex(C['bg']))

        log.info('═' * 60)
        log.info('Dashboard v0.35 starting (wx)')

        self._set_app_icon()
        self._build_topbar()
        self._build_content()
        self._build_statusbar()

        self.root.Bind(wx.EVT_CLOSE, self._on_close)

        # Global key bindings
        self.root.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

        # Start auth after UI is shown
        wx.CallLater(100, self._start_auth)

    def _on_close(self, event):
        self.root.Destroy()

    # ── App Icon ───────────────────────────────────────────────

    def _set_app_icon(self):
        try:
            ico_path = os.path.join(CONFIG_DIR, 'email_dashboard.ico')
            if not os.path.exists(ico_path):
                ensure_config_dir()
                with open(ico_path, 'wb') as f:
                    f.write(base64.b64decode(_APP_ICON_B64))
            icon = wx.Icon(ico_path, wx.BITMAP_TYPE_ICO)
            self.root.SetIcon(icon)
        except Exception:
            pass

    # ── Top Bar ────────────────────────────────────────────────

    def _build_topbar(self):
        tb_bg = _hex(C['topbar_bg'])
        tb_fg = _hex(C['topbar_text'])
        tb_fg2 = _hex(C['topbar_text2'])

        panel = wx.Panel(self.root)
        panel.SetBackgroundColour(tb_bg)

        sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Logo
        lbl_logo = wx.StaticText(panel, label='⚡✉  Email Intelligence')
        lbl_logo.SetFont(_font(FONT_BOLD, 13))
        lbl_logo.SetForegroundColour(tb_fg)
        sizer.Add(lbl_logo, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 12)

        sizer.AddStretchSpacer(1)

        # Search
        self.search_ctrl = wx.SearchCtrl(panel, size=(320, -1))
        self.search_ctrl.SetDescriptiveText('Search emails...')
        self.search_ctrl.Bind(wx.EVT_SEARCH, self._on_search)
        self.search_ctrl.Bind(wx.EVT_SEARCH_CANCEL, lambda e: self._clear_search())
        sizer.Add(self.search_ctrl, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.RIGHT, 8)

        sizer.AddStretchSpacer(1)

        # Right controls
        self.user_label = wx.StaticText(panel, label='',
                                         style=wx.ST_ELLIPSIZE_END | wx.ST_NO_AUTORESIZE)
        self.user_label.SetForegroundColour(tb_fg2)
        self.user_label.SetMaxSize((160, -1))
        sizer.Add(self.user_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)

        for text, handler in [
            ('↻ Refresh', self._refresh),
            ('⚙ Rules', self._open_scoring_settings),
            ('ℹ About', self._show_about),
            ('Sign Out', self._sign_out),
        ]:
            b = wx.Button(panel, label=text)
            b.Bind(wx.EVT_BUTTON, lambda e, h=handler: h())
            sizer.Add(b, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 2)

        self.sort_btn = wx.Button(panel, label='Sort: Date')
        self.sort_btn.Bind(wx.EVT_BUTTON, lambda e: self._toggle_sort())
        sizer.Add(self.sort_btn, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 2)

        self._acct_btn = wx.Button(panel, label='👥 All ▾')
        self._acct_btn.Bind(wx.EVT_BUTTON, lambda e: self._show_account_menu())
        sizer.Add(self._acct_btn, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.RIGHT, 2)

        panel.SetSizer(sizer)

        main_sizer = getattr(self, '_main_sizer', None)
        if main_sizer is None:
            self._main_sizer = wx.BoxSizer(wx.VERTICAL)
            self.root.SetSizer(self._main_sizer)

        self._main_sizer.Add(panel, 0, wx.EXPAND)

    # ── Main Content (splitter) ────────────────────────────────

    def _build_content(self):
        self._splitter = wx.SplitterWindow(self.root, style=wx.SP_LIVE_UPDATE | wx.SP_3D)
        self._splitter.SetSashGravity(0.0)
        self._splitter.SetMinimumPaneSize(350)
        self._main_sizer.Add(self._splitter, 1, wx.EXPAND)

        self._build_list_panel()
        self._build_detail_panel()
        self._splitter.SplitVertically(self._list_panel, self._detail_panel, 620)

    # ── Email List Panel ───────────────────────────────────────

    def _build_list_panel(self):
        self._list_panel = wx.Panel(self._splitter)
        self._list_panel.SetBackgroundColour(_hex(C['bg']))
        outer = wx.BoxSizer(wx.VERTICAL)

        # ── Filter bar ────────────────────────────────────────
        filter_panel = wx.Panel(self._list_panel)
        filter_panel.SetBackgroundColour(_hex(C['bg_card']))
        filter_sizer = wx.BoxSizer(wx.VERTICAL)

        # Row 1: Folder + TZ
        row1 = wx.BoxSizer(wx.HORIZONTAL)
        row1.Add(wx.StaticText(filter_panel, label='Folder:'), 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 8)
        self.folder_var = wx.Choice(filter_panel, choices=[
            'inbox', 'sentitems', 'drafts', 'archive',
            'deleteditems', 'snoozed / reminded', 'send queue'])
        self.folder_var.SetSelection(0)
        self.folder_var.Bind(wx.EVT_CHOICE, lambda e: self._on_folder_change())
        row1.Add(self.folder_var, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 4)

        row1.Add(wx.StaticText(filter_panel, label='TZ:'), 0,
                 wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 12)
        self.tz_var = wx.Choice(filter_panel, choices=list(TIMEZONES.keys()))
        self.tz_var.SetStringSelection(DEFAULT_TZ)
        row1.Add(self.tz_var, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 4)
        filter_sizer.Add(row1, 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 4)

        # Row 2: Date filters
        row2 = wx.BoxSizer(wx.HORIZONTAL)
        row2.Add(wx.StaticText(filter_panel, label='After:'), 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 8)
        self.after_entry = wx.TextCtrl(filter_panel, value='YYYY-MM-DD', size=(95, -1))
        self.after_entry.Bind(wx.EVT_SET_FOCUS,
            lambda e: (self.after_entry.SelectAll() if self.after_entry.GetValue() == 'YYYY-MM-DD' else None, e.Skip()))
        row2.Add(self.after_entry, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 4)

        row2.Add(wx.StaticText(filter_panel, label='Before:'), 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 8)
        self.before_entry = wx.TextCtrl(filter_panel, value='YYYY-MM-DD', size=(95, -1))
        self.before_entry.Bind(wx.EVT_SET_FOCUS,
            lambda e: (self.before_entry.SelectAll() if self.before_entry.GetValue() == 'YYYY-MM-DD' else None, e.Skip()))
        row2.Add(self.before_entry, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 4)

        btn_apply = wx.Button(filter_panel, label='Apply', size=(-1, -1))
        btn_apply.Bind(wx.EVT_BUTTON, lambda e: self._full_refresh())
        row2.Add(btn_apply, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 8)
        btn_clear = wx.Button(filter_panel, label='Clear', size=(-1, -1))
        btn_clear.Bind(wx.EVT_BUTTON, lambda e: self._clear_filters())
        row2.Add(btn_clear, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 2)
        filter_sizer.Add(row2, 0, wx.EXPAND | wx.BOTTOM, 6)

        filter_panel.SetSizer(filter_sizer)
        outer.Add(filter_panel, 0, wx.EXPAND)

        # ── Stats label ────────────────────────────────────────
        self.stats_label = wx.StaticText(self._list_panel, label='')
        self.stats_label.SetForegroundColour(_hex(C['muted']))
        outer.Add(self.stats_label, 0, wx.LEFT | wx.TOP, 8)

        # ── Split inbox tabs ───────────────────────────────────
        tab_panel = wx.Panel(self._list_panel)
        tab_panel.SetBackgroundColour(_hex(C['bg']))
        tab_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self._split_btns = {}
        for label, mode in [('All', 'all'), ('⭐ VIP', 'vip'), ('👥 Team', 'team'),
                             ('📰 Other', 'newsletters'), ('🔕 Low', 'low')]:
            btn = wx.Button(tab_panel, label=label, size=(-1, 26))
            btn.Bind(wx.EVT_BUTTON, lambda e, m=mode: self._set_split(m))
            tab_sizer.Add(btn, 0, wx.LEFT, 2)
            self._split_btns[mode] = btn
        tab_panel.SetSizer(tab_sizer)
        outer.Add(tab_panel, 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 4)

        # ── Scrollable email list ──────────────────────────────
        self._list_scroll = wx.ScrolledWindow(self._list_panel, style=wx.VSCROLL | wx.FULL_REPAINT_ON_RESIZE)
        self._list_scroll.SetDoubleBuffered(True)
        self._list_scroll.SetScrollRate(0, 5)   # 5px units — finer for smoother scroll
        self._list_scroll.SetBackgroundColour(_hex(C['bg']))
        self.list_inner_sizer = wx.BoxSizer(wx.VERTICAL)
        self._list_scroll.SetSizer(self.list_inner_sizer)
        self._wheel_accum = 0  # fractional accumulator for precision trackpads

        # MouseWheel
        self._list_scroll.Bind(wx.EVT_MOUSEWHEEL, self._on_list_mousewheel)

        outer.Add(self._list_scroll, 1, wx.EXPAND)

        # ── Load More ──────────────────────────────────────────
        self._load_more_btn = wx.Button(self._list_panel, label='⬇  Load More Emails')
        self._load_more_btn.Bind(wx.EVT_BUTTON, lambda e: self._load_more())
        outer.Add(self._load_more_btn, 0, wx.EXPAND | wx.ALL, 8)
        self._all_loaded = False

        self._list_panel.SetSizer(outer)
        self._update_split_tabs()

    def _on_list_mousewheel(self, event):
        # Accumulate fractional wheel rotation for precision trackpads / high-res mice.
        # GetWheelRotation() returns the raw delta (e.g. 120 per notch on a standard wheel,
        # but can be any value for smooth-scroll devices). GetWheelDelta() is the notch size
        # (typically 120). We translate to pixels directly instead of using ScrollLines()
        # to avoid the "snap to nearest scroll unit" stutter.
        PIXELS_PER_NOTCH = 80  # how far one full notch scrolls; tune to taste
        delta = event.GetWheelDelta()
        if delta <= 0:
            delta = 120
        self._wheel_accum -= event.GetWheelRotation() * PIXELS_PER_NOTCH
        px = int(self._wheel_accum / delta)
        if px == 0:
            return
        self._wheel_accum -= px * delta

        _, scroll_unit = self._list_scroll.GetScrollPixelsPerUnit()
        if scroll_unit <= 0:
            scroll_unit = 5
        cur = self._list_scroll.GetScrollPos(wx.VERTICAL)
        new_pos = max(0, cur + (px // scroll_unit))
        self._list_scroll.Scroll(-1, new_pos)

    # ── Detail Panel ──────────────────────────────────────────

    def _build_detail_panel(self):
        self._detail_panel = wx.Panel(self._splitter)
        self._detail_panel.SetBackgroundColour(_hex(C['bg']))
        outer = wx.BoxSizer(wx.VERTICAL)

        # ── Header ────────────────────────────────────────────
        hdr = wx.Panel(self._detail_panel)
        hdr.SetBackgroundColour(_hex(C['bg_card']))
        hdr_sizer = wx.BoxSizer(wx.VERTICAL)

        self.d_subject = wx.StaticText(hdr, label='Select an email to view')
        self.d_subject.SetFont(_font(FONT_BOLD, 13))
        hdr_sizer.Add(self.d_subject, 0, wx.EXPAND | wx.ALL, 8)

        meta_row = wx.BoxSizer(wx.HORIZONTAL)
        self.d_from = wx.StaticText(hdr, label='')
        self.d_from.SetForegroundColour(_hex(C['text2']))
        meta_row.Add(self.d_from, 1)
        self.d_date = wx.StaticText(hdr, label='')
        self.d_date.SetForegroundColour(_hex(C['muted']))
        meta_row.Add(self.d_date, 0)
        hdr_sizer.Add(meta_row, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        self.d_to = wx.StaticText(hdr, label='')
        self.d_to.SetForegroundColour(_hex(C['muted']))
        hdr_sizer.Add(self.d_to, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        hdr.SetSizer(hdr_sizer)
        outer.Add(hdr, 0, wx.EXPAND)

        # ── Action bar ────────────────────────────────────────
        ab = wx.Panel(self._detail_panel)
        ab.SetBackgroundColour(_hex(C['bg']))
        ab_sizer = wx.BoxSizer(wx.HORIZONTAL)

        def _ab_btn(label, handler):
            b = wx.Button(ab, label=label)
            b.Bind(wx.EVT_BUTTON, lambda e: handler())
            return b

        self._archive_accept_btn = _ab_btn('&Archive', self._on_archive_accept)
        self._current_action_is_accept = False
        ab_sizer.Add(self._archive_accept_btn, 0, wx.RIGHT, 4)

        self._reply_btn = _ab_btn('&Reply', self._reply)
        ab_sizer.Add(self._reply_btn, 0, wx.RIGHT, 4)

        self._reply_all_btn = _ab_btn('Reply &All', self._reply_all)
        ab_sizer.Add(self._reply_all_btn, 0, wx.RIGHT, 4)

        self._forward_btn = _ab_btn('&Forward', self._forward)
        ab_sizer.Add(self._forward_btn, 0, wx.RIGHT, 4)

        self._snooze_btn = _ab_btn('&Snooze', self._show_snooze_menu)
        ab_sizer.Add(self._snooze_btn, 0, wx.RIGHT, 4)

        self._remind_btn = _ab_btn('Re&mind', self._show_remind_menu)
        ab_sizer.Add(self._remind_btn, 0, wx.RIGHT, 4)

        self._attach_btn = wx.Button(ab, label='📎 Attachments')
        self._attach_btn.SetBackgroundColour(_hex(C['blue']))
        self._attach_btn.SetForegroundColour(wx.WHITE)
        self._attach_btn.Bind(wx.EVT_BUTTON, lambda e: self._toggle_attachments())
        ab_sizer.Add(self._attach_btn, 0, wx.RIGHT, 4)
        self._attach_btn.Hide()

        ab_sizer.AddStretchSpacer()

        self._auto_archive_btn = _ab_btn('Auto-&Archive', self._auto_archive_sender)
        ab_sizer.Add(self._auto_archive_btn, 0, wx.RIGHT, 4)
        self._delete_btn = _ab_btn('&Delete', self._delete)
        ab_sizer.Add(self._delete_btn, 0)

        # Send queue / snoozed buttons (hidden by default)
        self._sq_cancel_btn = _ab_btn('✕ Cancel Send', self._cancel_queued_email)
        self._sq_sendnow_btn = _ab_btn('📤 Send Now', self._send_queued_now)
        self._snz_return_btn = _ab_btn('📥 Return to Inbox', self._return_to_inbox)
        self._snz_reschedule_btn = _ab_btn('⏰ Reschedule', self._show_reschedule_menu)
        self._snz_delete_btn = _ab_btn('🗑 Delete', self._delete_snoozed)
        for b in [self._sq_cancel_btn, self._sq_sendnow_btn,
                  self._snz_return_btn, self._snz_reschedule_btn, self._snz_delete_btn]:
            ab_sizer.Add(b, 0, wx.RIGHT, 4)
            b.Hide()

        self._normal_action_btns = [
            self._archive_accept_btn, self._reply_btn, self._reply_all_btn,
            self._forward_btn, self._snooze_btn, self._remind_btn,
            self._delete_btn, self._auto_archive_btn,
        ]
        self._sq_btns = [self._sq_cancel_btn, self._sq_sendnow_btn]
        self._snz_btns = [self._snz_return_btn, self._snz_reschedule_btn, self._snz_delete_btn]

        ab.SetSizer(ab_sizer)
        outer.Add(ab, 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 4)

        # ── Attachment panel ───────────────────────────────────
        self._attach_frame = wx.Panel(self._detail_panel)
        self._attach_frame.SetBackgroundColour(_hex('#F0F4FF'))
        self._attach_inner_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self._attach_frame.SetSizer(self._attach_inner_sizer)
        outer.Add(self._attach_frame, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)
        self._attach_frame.Hide()

        # ── Intel bar ─────────────────────────────────────────
        intel_panel = wx.Panel(self._detail_panel)
        intel_panel.SetBackgroundColour(_hex(C['bg_card']))
        intel_sizer = wx.BoxSizer(wx.VERTICAL)

        top_row = wx.BoxSizer(wx.HORIZONTAL)
        self.d_priority = wx.StaticText(intel_panel, label='IMPORTANT')  # widest label reserves space
        self.d_priority.SetFont(_font(FONT_BOLD, 10))
        self.d_priority.SetMinSize((85, -1))
        top_row.Add(self.d_priority, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 12)
        self.d_score = wx.StaticText(intel_panel, label='')
        self.d_score.SetForegroundColour(_hex(C['blue']))
        self.d_score.SetFont(_font(FONT, 10))
        top_row.Add(self.d_score, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)
        train_btn = wx.Button(intel_panel, label='⚙ Learn', style=wx.BU_EXACTFIT)
        train_btn.SetFont(_font(FONT, 9))
        train_btn.Bind(wx.EVT_BUTTON, lambda e: self._show_train_rules())
        top_row.Add(train_btn, 0, wx.ALIGN_CENTER_VERTICAL)
        top_row.AddStretchSpacer(1)
        intel_sizer.Add(top_row, 0, wx.EXPAND | wx.ALL, 8)

        self.d_signals = wx.StaticText(intel_panel, label='')
        self.d_signals.SetForegroundColour(_hex(C['muted']))
        intel_sizer.Add(self.d_signals, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        intel_panel.SetSizer(intel_sizer)
        outer.Add(intel_panel, 0, wx.EXPAND)

        # ── Event time bar ─────────────────────────────────────
        self._event_time_frame = wx.Panel(self._detail_panel)
        self._event_time_frame.SetBackgroundColour(_hex('#EFF6FF'))
        et_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.d_event_time = wx.StaticText(self._event_time_frame, label='')
        self.d_event_time.SetForegroundColour(_hex('#1E40AF'))
        self.d_event_time.SetFont(_font(FONT_BOLD, 10))
        et_sizer.Add(self.d_event_time, 0, wx.ALL, 6)
        self._event_time_frame.SetSizer(et_sizer)
        outer.Add(self._event_time_frame, 0, wx.EXPAND)
        self._event_time_frame.Hide()

        # ── Load Images bar ────────────────────────────────────
        self._load_images_frame = wx.Panel(self._detail_panel)
        self._load_images_frame.SetBackgroundColour(_hex('#FEF3C7'))
        li_sizer = wx.BoxSizer(wx.HORIZONTAL)
        li_warn = wx.StaticText(self._load_images_frame,
                                label='⚠ Remote images blocked for security')
        li_warn.SetForegroundColour(_hex('#92400E'))
        li_sizer.Add(li_warn, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 8)
        li_load = wx.Button(self._load_images_frame, label='Load Images')
        li_load.Bind(wx.EVT_BUTTON, lambda e: self._load_images())
        li_sizer.Add(li_load, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 8)
        li_always = wx.Button(self._load_images_frame, label='Always Load')
        li_always.SetBackgroundColour(_hex(C['green']))
        li_always.SetForegroundColour(wx.WHITE)
        li_always.Bind(wx.EVT_BUTTON, lambda e: self._safe_sender_images())
        li_sizer.Add(li_always, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 4)
        self._load_images_frame.SetSizer(li_sizer)
        outer.Add(self._load_images_frame, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)
        self._load_images_frame.Hide()

        # ── Reply composer ─────────────────────────────────────
        self.reply_frame = wx.Panel(self._detail_panel)
        self.reply_frame.SetBackgroundColour(_hex(C['bg_card']))
        self._build_reply_composer()
        outer.Add(self.reply_frame, 0, wx.EXPAND | wx.ALL, 4)
        self.reply_frame.Hide()

        # ── Send undo bar (shown after queuing a send) ─────────
        self._undo_bar = wx.Panel(self._detail_panel)
        self._undo_bar.SetBackgroundColour(_hex(C.get('orange', '#E8A000')))
        ub_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self._undo_countdown_label = wx.StaticText(self._undo_bar, label="")
        self._undo_countdown_label.SetForegroundColour(wx.WHITE)
        self._undo_countdown_label.SetFont(_font(FONT, 9))
        ub_sizer.Add(self._undo_countdown_label, 1, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 8)
        undo_btn = wx.StaticText(self._undo_bar, label="✕ UNDO")
        undo_btn.SetForegroundColour(wx.WHITE)
        undo_btn.SetFont(_font(FONT_BOLD, 9))
        undo_btn.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        undo_btn.Bind(wx.EVT_LEFT_UP,
                      lambda e: self._undo_queued_send(getattr(self, '_undo_active_draft', None)))
        ub_sizer.Add(undo_btn, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 12)
        self._undo_bar.SetSizer(ub_sizer)
        outer.Add(self._undo_bar, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 4)
        self._undo_bar.Hide()

        # ── Email body (WebView) ───────────────────────────────
        self._body_container = wx.Panel(self._detail_panel)
        self._body_container.SetBackgroundColour(_hex(C['bg']))
        self._body_container.SetDoubleBuffered(True)
        body_sizer = wx.BoxSizer(wx.VERTICAL)
        self.body = EmailWebView(self._body_container)
        body_sizer.Add(self.body, 1, wx.EXPAND)
        self._body_container.SetSizer(body_sizer)
        outer.Add(self._body_container, 1, wx.EXPAND)

        self._detail_panel.SetSizer(outer)

    # ── Reply Composer ─────────────────────────────────────────

    def _build_reply_composer(self):
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Header row
        hdr = wx.BoxSizer(wx.HORIZONTAL)
        self.reply_label = wx.StaticText(self.reply_frame, label='✏️ Reply')
        self.reply_label.SetFont(_font(FONT_BOLD, 11))
        hdr.Add(self.reply_label, 1, wx.ALIGN_CENTER_VERTICAL)
        hdr.AddStretchSpacer()

        link_subj = wx.StaticText(self.reply_frame, label='✎ Subject')
        link_subj.SetForegroundColour(_hex(C['muted']))
        link_subj.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        link_subj.Bind(wx.EVT_LEFT_UP, lambda e: self._toggle_edit_subject())
        hdr.Add(link_subj, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)

        self._edit_recipients_link = wx.StaticText(self.reply_frame, label='✎ Recipients')
        self._edit_recipients_link.SetForegroundColour(_hex(C['muted']))
        self._edit_recipients_link.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self._edit_recipients_link.Bind(wx.EVT_LEFT_UP, lambda e: self._toggle_edit_recipients())
        hdr.Add(self._edit_recipients_link, 0, wx.ALIGN_CENTER_VERTICAL)
        sizer.Add(hdr, 0, wx.EXPAND | wx.ALL, 8)

        # Editable subject (hidden by default)
        self._edit_subject_frame = wx.Panel(self.reply_frame)
        esf_s = wx.BoxSizer(wx.HORIZONTAL)
        esf_s.Add(wx.StaticText(self._edit_subject_frame, label='Subject:'), 0,
                  wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)
        self._edit_subject_entry = wx.TextCtrl(self._edit_subject_frame)
        esf_s.Add(self._edit_subject_entry, 1)
        self._edit_subject_frame.SetSizer(esf_s)
        sizer.Add(self._edit_subject_frame, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)
        self._edit_subject_frame.Hide()

        # Editable recipients (hidden by default)
        self._edit_to_frame = wx.Panel(self.reply_frame)
        etf_s = wx.BoxSizer(wx.HORIZONTAL)
        etf_s.Add(wx.StaticText(self._edit_to_frame, label='To:'), 0,
                  wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)
        self._edit_to_entry = wx.TextCtrl(self._edit_to_frame)
        etf_s.Add(self._edit_to_entry, 1)
        self._edit_to_frame.SetSizer(etf_s)
        sizer.Add(self._edit_to_frame, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)
        self._edit_to_frame.Hide()

        # Forward To field
        self._fwd_to_frame = wx.Panel(self.reply_frame)
        ftf_s = wx.BoxSizer(wx.HORIZONTAL)
        _fwd_to_lbl = wx.StaticText(self._fwd_to_frame, label='To:')
        ftf_s.Add(_fwd_to_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)
        self._fwd_to_entry = wx.TextCtrl(self._fwd_to_frame)
        self._fwd_to_ac = EmailAutocomplete(self._fwd_to_entry,
                                             lambda: getattr(self, '_address_book', []))
        ftf_s.Add(self._fwd_to_entry, 1)
        self._fwd_to_frame.SetSizer(ftf_s)
        sizer.Add(self._fwd_to_frame, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)
        self._fwd_to_frame.Hide()
        # Clicking the label or panel background redirects focus to the TextCtrl
        def _fwd_to_click(e):
            def _f():
                self._fwd_to_entry.SetFocus()
                self._fwd_to_entry.SetInsertionPointEnd()
            wx.CallAfter(_f)
            e.Skip()
        for _w in (self._fwd_to_frame, _fwd_to_lbl):
            _w.Bind(wx.EVT_LEFT_DOWN, _fwd_to_click)

        # CC field
        self._cc_frame = wx.Panel(self.reply_frame)
        ccf_s = wx.BoxSizer(wx.HORIZONTAL)
        _cc_lbl = wx.StaticText(self._cc_frame, label='CC:')
        ccf_s.Add(_cc_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)
        self._cc_entry = wx.TextCtrl(self._cc_frame)
        self._cc_ac = EmailAutocomplete(self._cc_entry,
                                         lambda: getattr(self, '_address_book', []))
        ccf_s.Add(self._cc_entry, 1)
        self._cc_frame.SetSizer(ccf_s)
        sizer.Add(self._cc_frame, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 4)
        self._cc_frame.Hide()
        def _cc_click(e):
            def _f():
                self._cc_entry.SetFocus()
                self._cc_entry.SetInsertionPointEnd()
            wx.CallAfter(_f)
            e.Skip()
        for _w in (self._cc_frame, _cc_lbl):
            _w.Bind(wx.EVT_LEFT_DOWN, _cc_click)

        # Reply text area
        self.reply_text = wx.TextCtrl(self.reply_frame,
                                      style=wx.TE_MULTILINE,
                                      size=(-1, 120))
        self.reply_text.SetBackgroundColour(wx.WHITE)
        self.reply_text.SetFont(_font(FONT, 10))
        self.reply_text.Bind(wx.EVT_KEY_UP, self._on_reply_key)
        self.reply_text.Bind(wx.EVT_CONTEXT_MENU, self._on_spell_right_click)
        # Track compose focus so char hook doesn't steal keys/focus during composition.
        # Use only SET_FOCUS to enter compose mode. For KILL_FOCUS, defer via CallLater
        # so FindFocus() sees the new widget (Windows moves focus after KILL_FOCUS fires).
        def _set_compose_focus(e):
            self._focus_pane = 'compose'
            e.Skip()
        def _clear_compose_focus(e):
            e.Skip()
            def _check():
                try:
                    if not self.reply_frame.IsShown():
                        self._focus_pane = 'list'
                        return
                    fw = wx.Window.FindFocus()
                    if fw is None:
                        return  # transient — don't change pane
                    p = fw
                    depth = 0
                    while p and depth < 20:
                        if p is self.reply_frame:
                            return  # still inside compose
                        p = p.GetParent()
                        depth += 1
                    self._focus_pane = 'list'
                except Exception:
                    pass  # widget may have been destroyed — ignore
            wx.CallLater(50, _check)
        for _cw in (self.reply_text, self._fwd_to_entry, self._cc_entry,
                    self._edit_to_entry, self._edit_subject_entry):
            _cw.Bind(wx.EVT_SET_FOCUS, _set_compose_focus)
            _cw.Bind(wx.EVT_KILL_FOCUS, _clear_compose_focus)
            # Force focus via CallAfter on click so WebView releases before we grab
            def _make_click_handler(w):
                def _h(e):
                    def _focus_and_caret():
                        w.SetFocus()
                        w.SetInsertionPointEnd()
                    wx.CallAfter(_focus_and_caret)
                    e.Skip()
                return _h
            _cw.Bind(wx.EVT_LEFT_DOWN, _make_click_handler(_cw))
        sizer.Add(self.reply_text, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)

        # Signature preview
        self._sig_preview_frame = wx.Panel(self.reply_frame)
        self._sig_preview_frame.SetBackgroundColour(_hex('#F8FAFC'))
        sig_s = wx.BoxSizer(wx.VERTICAL)
        self._sig_include_var = wx.CheckBox(self._sig_preview_frame,
                                            label='Include signature')
        self._sig_include_var.SetValue(True)
        self._sig_include_var.Bind(wx.EVT_CHECKBOX, lambda e: self._on_sig_toggle())
        sig_s.Add(self._sig_include_var, 0, wx.ALL, 4)
        self._sig_preview_label = wx.StaticText(self._sig_preview_frame, label='')
        self._sig_preview_label.SetForegroundColour(_hex(C['text2']))
        sig_s.Add(self._sig_preview_label, 0, wx.EXPAND | wx.LEFT | wx.BOTTOM, 8)
        self._sig_preview_frame.SetSizer(sig_s)
        sizer.Add(self._sig_preview_frame, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)
        self._sig_preview_frame.Hide()

        # Send row
        send_row = wx.BoxSizer(wx.HORIZONTAL)
        self._send_btn = wx.Button(self.reply_frame, label='📤 Send')
        self._send_btn.SetBackgroundColour(_hex(C['btn_primary']))
        self._send_btn.SetForegroundColour(wx.WHITE)
        self._send_btn.Bind(wx.EVT_BUTTON, lambda e: self._on_send())
        send_row.Add(self._send_btn, 0, wx.RIGHT, 4)

        btn_later = wx.Button(self.reply_frame, label='⏰', size=(36, -1))
        btn_later.Bind(wx.EVT_BUTTON, lambda e: self._show_send_later_menu())
        send_row.Add(btn_later, 0, wx.RIGHT, 4)

        btn_fix = wx.Button(self.reply_frame, label='🔧 Fix All')
        btn_fix.Bind(wx.EVT_BUTTON, lambda e: self._fix_all_errors())
        send_row.Add(btn_fix, 0, wx.RIGHT, 4)

        btn_cancel = wx.Button(self.reply_frame, label='Cancel')
        btn_cancel.Bind(wx.EVT_BUTTON, lambda e: self._cancel_reply())
        send_row.Add(btn_cancel, 0)

        sizer.Add(send_row, 0, wx.ALL, 8)
        self.reply_frame.SetSizer(sizer)

    # ── Status Bar ────────────────────────────────────────────

    def _build_statusbar(self):
        self._status_bar = self.root.CreateStatusBar(2)
        self._status_bar.SetStatusWidths([-1, 130])
        self._status_bar.SetStatusText('Starting...', 0)
        # wx has no built-in progress in status bar; we simulate with text
        self._progress_running = False

    def _set_status(self, t):
        wx.CallAfter(self._status_bar.SetStatusText, t, 0)

    # progress shims (used in email_loading and elsewhere)
    class _FakeProgress:
        def start(self, *a): pass
        def stop(self): pass

    @property
    def progress(self):
        return self._FakeProgress()

    # ── Key bindings ───────────────────────────────────────────

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        ctrl = event.ControlDown()

        # Ctrl+K = command palette
        if ctrl and key in (ord('K'), ord('k')):
            self._on_command_palette()
            return

        # Don't fire single-key shortcuts while typing or composing
        if self._focus_pane == 'compose' or self._is_typing():
            event.Skip()
            return

        if key == wx.WXK_UP:
            if self._focus_pane == 'list':
                self._select_adjacent(-1)
            else:
                self._scroll_body('up')
        elif key == wx.WXK_DOWN:
            if self._focus_pane == 'list':
                self._select_adjacent(1)
            else:
                self._scroll_body('down')
        elif key == wx.WXK_DELETE:
            self._on_key_delete()
        elif key in (ord('a'), ord('A')):
            self._on_key_archive()
        elif key == ord('r'):
            self._on_key_reply()
        elif key == ord('R'):
            self._on_key_reply_all()
        elif key in (ord('f'), ord('F')):
            self._on_key_forward()
        elif key in (ord('s'), ord('S')):
            self._on_key_snooze()
        elif key in (ord('m'), ord('M')):
            self._on_key_remind()
        else:
            event.Skip()

    # ── after() compatibility shim ─────────────────────────────
    # Many mixin methods call self.root.after(ms, fn).
    # wx uses wx.CallLater or wx.CallAfter.

    def _patch_root_after(self):
        """Monkey-patch self.root so mixins can call self.root.after().
        Always routes through wx.CallAfter so it is safe from any thread.
        wx.CallLater can only be started from the main thread, so we defer
        timer creation via wx.CallAfter when ms > 0.
        """
        frame = self.root

        def _after(ms, fn, *args):
            if ms == 0:
                wx.CallAfter(fn, *args)
            else:
                def _start_timer():
                    wx.CallLater(ms, fn, *args)
                wx.CallAfter(_start_timer)

        def _after_idle(fn, *args):
            wx.CallAfter(fn, *args)

        frame.after = _after
        frame.after_idle = _after_idle

    # ── Layout refresh helper ──────────────────────────────────

    def _refresh_layout(self):
        """Force a layout pass on the detail panel."""
        self._detail_panel.Layout()

    def _update_scroll_region(self):
        self._list_scroll.FitInside()

    # ── run() entry point ─────────────────────────────────────

    def run(self):
        self._patch_root_after()
        self.root.Show()
        self._wx_app.MainLoop()
