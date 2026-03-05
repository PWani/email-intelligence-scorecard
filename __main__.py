import os, sys, traceback
os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"] = "1"
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except Exception: pass

from .core.config import _detect_font, log
from .app import EmailDashboard

def _exc(t, v, tb):
    log.critical("UNCAUGHT:\n%s", "".join(traceback.format_exception(t, v, tb)))
sys.excepthook = _exc

# Filter libjpeg noise: WebView probes PNG/GIF inline images as JPEG and
# libjpeg writes "Not a JPEG file: starts with 0x89 0x50" directly to fd 2.
class _StderrFilter:
    _SUPPRESS = ("Not a JPEG file:", "JPEG datastream contains no image")
    def __init__(self, wrapped): self._w = wrapped
    def write(self, s):
        if not any(p in s for p in self._SUPPRESS):
            self._w.write(s)
    def flush(self): self._w.flush()
    def __getattr__(self, a): return getattr(self._w, a)

sys.stderr = _StderrFilter(sys.stderr)

# _detect_font is called after wx.App is created (inside __init__)
EmailDashboard().run()
