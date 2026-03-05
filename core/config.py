# Config and constants
import os, json, logging
from logging.handlers import RotatingFileHandler as _RotatingFileHandler

FONT = "Aptos"
FONT_BOLD = "Aptos Semibold"

def _detect_font():
    """Detect best available font at runtime using wx."""
    global FONT, FONT_BOLD
    try:
        import wx
        app = wx.App.Get()
        if app is None:
            app = wx.App()
        e = wx.FontEnumerator()
        e.EnumerateFacenames()
        available = set(f.lower() for f in e.GetFacenames())
        if "aptos" in available:
            FONT, FONT_BOLD = "Aptos", "Aptos Semibold"
        elif "segoe ui variable" in available:
            FONT, FONT_BOLD = "Segoe UI Variable", "Segoe UI Variable"
        else:
            FONT, FONT_BOLD = "Segoe UI", "Segoe UI Semibold"
    except Exception:
        FONT, FONT_BOLD = "Segoe UI", "Segoe UI Semibold"

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".outlook_dashboard")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
TOKEN_CACHE_FILE = os.path.join(CONFIG_DIR, "token_cache.bin")
ADDRESS_BOOK_FILE = os.path.join(CONFIG_DIR, "address_book_cache.json")
OFFLINE_QUEUE_FILE = os.path.join(CONFIG_DIR, "offline_queue.json")
SCORING_RULES_FILE = os.path.join(CONFIG_DIR, "scoring_rules.json")

# ── Application logger ────────────────────────────────────────
import logging
from logging.handlers import RotatingFileHandler as _RotatingFileHandler

os.makedirs(CONFIG_DIR, exist_ok=True)
_log_file = os.path.join(CONFIG_DIR, "dashboard.log")

# Clear log on every fresh start
try:
    with open(_log_file, "w", encoding="utf-8") as _f:
        _f.truncate(0)
except Exception:
    pass

_log_handler = _RotatingFileHandler(_log_file, maxBytes=2*1024*1024, backupCount=3, encoding="utf-8")
_log_handler.setFormatter(logging.Formatter(
    "%(asctime)s  %(levelname)-5s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))

# Flush after every log entry so crashes don't lose recent messages
class _FlushHandler(_RotatingFileHandler):
    def emit(self, record):
        super().emit(record)
        self.flush()

_log_handler_flush = _FlushHandler(_log_file, maxBytes=2*1024*1024, backupCount=3, encoding="utf-8")
_log_handler_flush.setFormatter(_log_handler.formatter)
log = logging.getLogger("dashboard")
log.setLevel(logging.INFO)
log.addHandler(_log_handler_flush)

# Embedded app icon (multi-size ICO, base64)
_APP_ICON_B64 = "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"

DEFAULT_CONFIG = {
    "client_id": "",
    "user_name": "",
    "scopes": ["User.Read", "Mail.Read", "Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite", "People.Read", "Contacts.Read"],
    "redirect_uri": "http://localhost:8400",
    "authority": "https://login.microsoftonline.com/common",
    "emails_per_page": 30,
    "undo_send_seconds": 60,
    "google_enabled": False,
    "google_credentials_file": "",
    "send_later_options": [
        {"label": "In 1 hour", "hours": 1},
        {"label": "In 2 hours", "hours": 2},
        {"label": "Tomorrow 8 AM", "preset": "tomorrow_morning"},
        {"label": "Monday 8 AM", "preset": "next_monday"},
    ],
    "snooze_options": [
        {"label": "6 hours", "hours": 6},
        {"label": "Tomorrow morning (9 AM)", "preset": "tomorrow_morning"},
        {"label": "Next week (7 days)", "preset": "next_week"},
    ],
    "remind_options": [
        {"label": "Remind if no reply (3 days)", "days": 3},
        {"label": "Remind if no reply (7 days)", "days": 7},
    ],
}


def ensure_config_dir():
    os.makedirs(CONFIG_DIR, exist_ok=True)


def load_config() -> dict:
    ensure_config_dir()
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            stored = json.load(f)
            result = dict(DEFAULT_CONFIG)
            # Overlay all stored keys onto defaults
            for key in stored:
                if stored[key] is not None and stored[key] != "":
                    result[key] = stored[key]
            return result
    return dict(DEFAULT_CONFIG)


def save_config(config: dict):
    ensure_config_dir()
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)


def load_address_book_cache():
    """Load cached address book from disk."""
    try:
        if os.path.exists(ADDRESS_BOOK_FILE):
            with open(ADDRESS_BOOK_FILE, "r") as f:
                return json.load(f)
    except Exception:
        pass
    return []


def save_address_book_cache(contacts):
    """Save address book to disk cache."""
    try:
        ensure_config_dir()
        with open(ADDRESS_BOOK_FILE, "w") as f:
            json.dump(contacts, f, indent=2)
    except Exception:
        pass

# ═══════════════════════════════════════════════════════════════
# THEME CONSTANTS
# ═══════════════════════════════════════════════════════════════

C = {
    "bg":           "#F5F6F8",
    "bg_card":      "#FFFFFF",
    "bg_hover":     "#EDF0F4",
    "bg_input":     "#FFFFFF",
    "border":       "#DDE1E8",
    "text":         "#1E293B",
    "text2":        "#475569",
    "muted":        "#8392A5",
    "blue":         "#2563EB",
    "green":        "#059669",
    "red":          "#DC2626",
    "orange":       "#D97706",
    "purple":       "#7C3AED",
    "urgent_bg":    "#FEF2F2",  "urgent_bd":   "#EF4444",
    "important_bg": "#FFFBEB",  "important_bd":"#F59E0B",
    "normal_bg":    "#FFFFFF",  "normal_bd":   "#DDE1E8",
    "low_bg":       "#F8FAFC",  "low_bd":      "#E2E8F0",
    "btn_primary":  "#2563EB",  "btn_primary_h":"#1D4ED8",
    "btn_danger":   "#DC2626",
    "btn_sec":      "#E2E8F0",  "btn_sec_h":   "#CBD5E1",
    "scrollbar_bg": "#F1F5F9",  "scrollbar_fg":"#CBD5E1",
    "topbar_bg":    "#1E3A5F",  "topbar_text": "#FFFFFF",
    "topbar_text2": "#94A3B8",
    "selected_bg":  "#DBEAFE",  "selected_bd":  "#3B82F6",
    "accent":       "#2563EB",
    "accent_light": "#EFF6FF",
    "accent_muted": "#93C5FD",
}
P_ICON = {"urgent": "🔴", "important": "🟠", "normal": "🔵", "low": "⚪"}
CAT_ICON = {"meeting_invite": "📅", "meeting": "📅", "action": "⚡", "newsletter": "📰", "fyi": "ℹ️", "general": "✉️"}

# Timezone definitions (offset hours from UTC)
TIMEZONES = {
    "Pacific (SF)":  -8,
    "Mountain":      -7,
    "Central":       -6,
    "Eastern (NY)":  -5,
    "UTC":            0,
    "London":         0,
    "CET":            1,
    "IST":            5.5,
    "JST":            9,
}
DEFAULT_TZ = "Pacific (SF)"

