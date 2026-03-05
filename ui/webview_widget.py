# Email body renderer — wx.html2.WebView backend
# Replaces tkinterweb with wx.html2 (WebKit on Windows/macOS/Linux).
# HTML is loaded directly via SetPage() — no local HTTP server needed.
# Remote image blocking is done via JS after load, not by stripping src attrs,
# so images can be un-blocked later without re-fetching the email.

import logging
import os
import re
import tempfile
import webbrowser
from html import unescape

import wx
import wx.html2

try:
    from bs4 import BeautifulSoup, Tag
    _HAS_BS4 = True
except ImportError:
    _HAS_BS4 = False

log = logging.getLogger('dashboard')

RENDERER = 'wx.html2'


# ── Plain-text helpers (used by scoring + plain fallback) ─────

def strip_html(html: str) -> str:
    """Remove all tags/style/script — used for scoring, not display."""
    html = re.sub(r'<style[^>]*>.*?</style>', ' ', html, flags=re.IGNORECASE | re.DOTALL)
    html = re.sub(r'<script[^>]*>.*?</script>', ' ', html, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<[^>]+>', ' ', html)
    text = unescape(text)
    return re.sub(r'\s+', ' ', text).strip()


def html_to_readable_text(html: str) -> str:
    """Convert HTML email to readable plain text with structure preserved."""
    if not html:
        return ''
    text = html
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</br>', '\n', text, flags=re.IGNORECASE)
    for tag in ['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                'tr', 'li', 'blockquote', 'section', 'article']:
        text = re.sub(rf'<{tag}[^>]*>', '\n\n', text, flags=re.IGNORECASE)
        text = re.sub(rf'</{tag}>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<hr[^>]*>', '\n' + '─' * 50 + '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<li[^>]*>', '\n  • ', text, flags=re.IGNORECASE)
    text = re.sub(r'<td[^>]*>', '  ', text, flags=re.IGNORECASE)
    text = re.sub(r'</td>', '\t', text, flags=re.IGNORECASE)
    text = re.sub(r'<th[^>]*>', '  ', text, flags=re.IGNORECASE)
    text = re.sub(r'</th>', '\t', text, flags=re.IGNORECASE)
    text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<[^>]+>', '', text)
    text = unescape(text)
    lines = text.split('\n')
    cleaned = [re.sub(r'[ \t]+', ' ', line).strip() for line in lines]
    text = '\n'.join(cleaned)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


# ── Compat shim ───────────────────────────────────────────────

def detect_renderer() -> str:
    return RENDERER


# ── Python-side image blocking (works on all WebView backends) ─
#
# Rather than relying on JS injection (Edge/WebView2 only), we mutate
# the HTML in Python before passing it to SetPage().  This means:
#   • Works on IE/legacy backend too
#   • No race condition between SetPage and JS injection
#   • load_images() restores from the original HTML cleanly
#
# Strategy:
#   block  → move src → data-blocked-src, replace with 1px transparent gif
#            strip CSS background-image
#   unblock→ reverse: restore src from data-blocked-src

_TRANSPARENT_GIF = ('data:image/gif;base64,'
                    'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7')

_REMOTE_URL_RE = re.compile(r'^https?://', re.IGNORECASE)
_BG_IMAGE_RE   = re.compile(r'url\s*\(\s*["\']?(https?://[^"\')\s]+)["\']?\s*\)',
                             re.IGNORECASE)


def block_remote_images(html: str) -> str:
    """Return HTML with all remote images neutralised using regex only.

    BS4 re-serialization mangles some email HTML (drops links in LinkedIn etc).
    Regex is safer — we only touch src= on img tags, nothing else.
    Each blocked <img> keeps its src in data-blocked-src for later restoration.
    """
    if not html:
        return html

    def _block_src(m):
        full = m.group(0)
        src = m.group(1)
        if not _REMOTE_URL_RE.match(src):
            return full
        return (full.replace(src, _TRANSPARENT_GIF, 1)
                + f' data-blocked-src="{src}"')

    # Replace src="https://..." on img tags only
    html = re.sub(
        r'(?i)(<img\b[^>]*?\bsrc=")([^"]*)"',
        lambda m: _block_src(m) if _REMOTE_URL_RE.match(m.group(2))
                  else m.group(0),
        html)
    html = re.sub(
        r"(?i)(<img\b[^>]*?\bsrc=')([^']*)'",
        lambda m: (m.group(0).replace(m.group(2), _TRANSPARENT_GIF, 1)
                   + f' data-blocked-src="{m.group(2)}"')
                  if _REMOTE_URL_RE.match(m.group(2)) else m.group(0),
        html)

    # Strip CSS background-image references
    html = _BG_IMAGE_RE.sub('url()', html)
    return html


def unblock_remote_images(html: str) -> str:
    """Restore remote images previously blocked by block_remote_images().

    Swaps data-blocked-src back to src.  Works on BS4 or regex fallback.
    """
    if not html:
        return html

    if _HAS_BS4:
        soup = BeautifulSoup(html, 'html.parser')
        for img in soup.find_all('img', attrs={'data-blocked-src': True}):
            img['src'] = img['data-blocked-src']
            del img['data-blocked-src']
            # Clean up the placeholder styling we added
            style = img.get('style', '')
            style = style.replace(';outline:1px dashed #ccc;opacity:0.3;', '')
            if style:
                img['style'] = style
            else:
                del img['style']
        return str(soup)

    else:
        # Regex fallback
        html = re.sub(
            r'<img([^>]*)\sdata-blocked-src=["\']([^"\']*)["\']([^>]*)src=["\'][^"\']*["\']',
            lambda m: f'<img{m.group(1)} src="{m.group(2)}"{m.group(3)}',
            html, flags=re.IGNORECASE)
        html = re.sub(r'\sdata-blocked-src=["\'][^"\']*["\']', '', html,
                      flags=re.IGNORECASE)
        return html


def has_remote_images(html: str) -> bool:
    """Return True if the HTML contains remote image references."""
    if _HAS_BS4:
        soup = BeautifulSoup(html, 'html.parser')
        for img in soup.find_all('img'):
            if _REMOTE_URL_RE.match(img.get('src', '')):
                return True
        for tag in soup.find_all(style=True):
            if _BG_IMAGE_RE.search(tag.get('style', '')):
                return True
        return False
    else:
        return bool(re.search(r'<img[^>]*src=["\']https?://', html, re.IGNORECASE))


# ── JS snippets injected after page load ─────────────────────
# All JS is written for maximum compatibility (IE11 / legacy WebView on Windows).
# No arrow functions, no const/let, no template literals, no .closest(),
# no .dataset, no NodeList.forEach (use Array.prototype.slice).

# Hides all remote <img> tags and CSS background-image rules,
# replacing them with a styled placeholder span.
_JS_BLOCK_IMAGES = r"""
(function() {
    try {
        var imgs = document.getElementsByTagName('img');
        var toReplace = [];
        var i;
        for (i = 0; i < imgs.length; i++) {
            var src = imgs[i].getAttribute('src') || '';
            if (/^https?:\/\//i.test(src)) { toReplace.push(imgs[i]); }
        }
        for (i = 0; i < toReplace.length; i++) {
            var img = toReplace[i];
            var src2 = img.getAttribute('src') || '';
            var alt = img.getAttribute('alt') || 'Image';
            var span = document.createElement('span');
            span.style.cssText = 'display:inline-block;background:#E5E7EB;color:#6B7280;padding:2px 6px;font-size:11px;border:1px solid #D1D5DB;';
            span.innerHTML = alt ? ('[img: ' + alt + ']') : '[img]';
            span.setAttribute('data-blocked-src', src2);
            if (img.parentNode) { img.parentNode.replaceChild(span, img); }
        }
        var styled = document.getElementsByTagName('*');
        for (i = 0; i < styled.length; i++) {
            var st = styled[i].style;
            if (st && st.backgroundImage && /url\(/i.test(st.backgroundImage)) {
                st.backgroundImage = 'none';
            }
        }
    } catch(e) {}
})();
"""

# Restores blocked images by swapping placeholder spans back to <img>.
_JS_UNBLOCK_IMAGES = r"""
(function() {
    try {
        var spans = document.getElementsByTagName('span');
        var toRestore = [];
        var i;
        for (i = 0; i < spans.length; i++) {
            if (spans[i].getAttribute('data-blocked-src')) { toRestore.push(spans[i]); }
        }
        for (i = 0; i < toRestore.length; i++) {
            var span = toRestore[i];
            var img = document.createElement('img');
            img.src = span.getAttribute('data-blocked-src');
            img.style.maxWidth = '100%';
            if (span.parentNode) { span.parentNode.replaceChild(img, span); }
        }
    } catch(e) {}
})();
"""

# Opens all links in the system browser via document.title signal to Python.
_JS_INTERCEPT_LINKS = r"""
(function() {
    if (window.__linksIntercepted) { return; }
    window.__linksIntercepted = true;
    function findAnchor(el) {
        while (el && el.tagName) {
            if (el.tagName === 'A' && el.getAttribute('href')) { return el; }
            el = el.parentNode;
        }
        return null;
    }
    document.onclick = function(e) {
        var evt = e || window.event;
        var target = evt.target || evt.srcElement;
        var a = findAnchor(target);
        if (!a) { return; }
        var href = a.getAttribute('href') || '';
        if (/^https?:\/\//i.test(href) || /^mailto:/i.test(href)) {
            if (evt.preventDefault) { evt.preventDefault(); } else { evt.returnValue = false; }
            document.title = '__OPEN__' + href;
            setTimeout(function() { document.title = ''; }, 300);
        }
    };
})();
"""


# ── Main widget ───────────────────────────────────────────────

class EmailWebView(wx.Panel):
    """
    wx.html2.WebView-backed email body renderer.

    Public API (drop-in replacement for the old tkinterweb EmailWebView):
        load_html(html, block_images=True)
        load_html_clean(html, block_images=True)  -- alias, no recreation needed
        load_images()                              -- un-block remote images in place
        get_selected_text() -> str
        get_full_text()     -> str
        open_in_browser()
        scroll(lines: int)

        is_html_capable    -> True
        has_size_limit     -> False
        is_showing_preview -> False
        renderer_name      -> 'wx.html2'
    """

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._last_html: str = ''
        self._block_images: bool = True
        self._loading_page: bool = False
        self._temp_html_path = os.path.join(
            tempfile.gettempdir(), 'email_dashboard_preview.html')

        # On Windows, prefer Edge/Chromium WebView2 over legacy IE.
        # wx.html2.WebViewBackendEdge is only present in wxPython 4.1+.
        backend = wx.html2.WebViewBackendDefault
        self._is_edge = False
        try:
            if hasattr(wx.html2, 'WebViewBackendEdge'):
                if wx.html2.WebView.IsBackendAvailable(wx.html2.WebViewBackendEdge):
                    backend = wx.html2.WebViewBackendEdge
                    self._is_edge = True
                    log.info('[webview] Using Edge/WebView2 backend')
        except Exception:
            pass

        if not self._is_edge:
            log.info('[webview] Using default (IE/WebKit) backend — JS injection disabled')

        self._wv = wx.html2.WebView.New(self, backend=backend)
        # Suppress script error dialogs — errors are caught in JS try/catch anyway
        try:
            self._wv.EnableContextMenu(False)
        except Exception:
            pass
        self._wv.Bind(wx.html2.EVT_WEBVIEW_NAVIGATING, self._on_navigating)
        self._wv.Bind(wx.html2.EVT_WEBVIEW_LOADED, self._on_loaded)
        self._wv.Bind(wx.html2.EVT_WEBVIEW_TITLE_CHANGED, self._on_title_changed)
        self._wv.Bind(wx.html2.EVT_WEBVIEW_NEWWINDOW, self._on_new_window)
        try:
            self._wv.Bind(wx.html2.EVT_WEBVIEW_ERROR, self._on_wv_error)
        except Exception:
            pass

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self._wv, 1, wx.EXPAND)
        self.SetSizer(sizer)
        self.SetDoubleBuffered(True)

        # Force WebView repaint on resize to eliminate GDI scroll artifacts
        # (text ghosting / strikethrough rendering seen when scrolling emails)
        self._wv.Bind(wx.EVT_SIZE, self._on_wv_size)

    def _on_wv_size(self, event):
        event.Skip()
        # Post a refresh after resize so WebView repaints cleanly
        wx.CallAfter(self._wv.Refresh)

    # ── Event handlers ────────────────────────────────────────

    def _on_wv_error(self, event: wx.html2.WebViewEvent):
        """Silently swallow WebView errors (e.g. JS eval failures on IE backend)."""
        log.debug('[webview] error: %s', event.GetString())

    def _on_navigating(self, event: wx.html2.WebViewEvent):
        url = event.GetURL()
        if not url or url in ('about:blank', '') or url.startswith('data:'):
            event.Allow()
            return
        if url.startswith('file://'):
            event.Allow()
            return
        # Allow navigations triggered by SetPage (sub-resource loads for images/CSS/fonts).
        # Only intercept user-initiated link clicks — detected by the _loading flag being clear.
        if getattr(self, '_loading_page', False):
            event.Allow()
            return
        if url.startswith(('http://', 'https://', 'mailto:')):
            webbrowser.open(url)
        event.Veto()

    def _on_new_window(self, event: wx.html2.WebViewEvent):
        """Handle target=_blank links (IE backend fires new-window instead of navigating)."""
        url = event.GetURL()
        if url and url.startswith(('http://', 'https://', 'mailto:')):
            webbrowser.open(url)
        event.Veto()

    def _on_loaded(self, event: wx.html2.WebViewEvent):
        """Clear loading flag and inject JS helpers after every page load."""
        self._loading_page = False  # must clear on ALL backends so link clicks are intercepted
        if not self._is_edge:
            return  # IE backend doesn't support RunScript — skip JS injection
        try:
            self._wv.RunScript(_JS_INTERCEPT_LINKS)
        except Exception as e:
            log.debug('[webview] link-intercept JS failed: %s', e)
        if self._block_images:
            try:
                self._wv.RunScript(_JS_BLOCK_IMAGES)
            except Exception as e:
                log.debug('[webview] block-images JS failed: %s', e)

    def _on_title_changed(self, event: wx.html2.WebViewEvent):
        """Receive link signals from JS via document.title."""
        title = self._wv.GetCurrentTitle()
        if title.startswith('__OPEN__'):
            url = title[len('__OPEN__'):]
            if url:
                webbrowser.open(url)

    # ── Public API ────────────────────────────────────────────

    def _strip_body_margins(self, html):
        """Strip margin/padding from body tag and CSS to minimise whitespace.
        IE backend ignores !important so we mutate the HTML directly."""
        import re as _re
        def _fix_body_attr(m):
            tag = m.group(0)
            tag = _re.sub(r'(?i)(margin|padding)\s*:[^;]+;?\s*', '', tag)
            return tag
        html = _re.sub(r'(?i)<body[^>]*>', _fix_body_attr, html)
        def _fix_css_body(m):
            block = m.group(0)
            block = _re.sub(r'(?i)(margin|padding)\s*:[^;]+;?\s*', '', block)
            return block
        html = _re.sub(r'(?is)body\s*\{[^}]*\}', _fix_css_body, html)
        return html

    def load_html(self, html: str, block_images: bool = True):
        """Load an HTML string. Remote images are blocked by default via Python-side HTML mutation."""
        self._last_html = html
        self._block_images = block_images

        # Strip body margin/padding — IE backend ignores CSS !important overrides
        display_base = self._strip_body_margins(html)

        # Strip target=_blank so IE fires EVT_WEBVIEW_NAVIGATING for all links
        import re as _re2
        display_base = _re2.sub(r'(?i)\s*target\s*=\s*["\']_blank["\']', '', display_base)

        # Block images in Python before handing to WebView — works on all backends.
        display_html = block_remote_images(display_base) if block_images else display_base

        try:
            log.info('[webview] SetPage %d chars block_images=%s', len(html), block_images)
            self._loading_page = True
            self._wv.SetPage(display_html, 'about:blank')
            # Safety net: IE backend sometimes skips EVT_WEBVIEW_LOADED — clear after 2s
            if not self._is_edge:
                wx.CallLater(2000, lambda: setattr(self, '_loading_page', False))
        except Exception as e:
            self._loading_page = False
            log.warning('[webview] SetPage failed: %s', e)
            self._load_plain_fallback(html)

    def load_html_clean(self, html: str, block_images: bool = True):
        """Alias — wx.html2 never needs widget recreation."""
        self.load_html(html, block_images=block_images)

    def load_images(self):
        """Un-block remote images in the currently loaded email without reloading."""
        self._block_images = False
        if self._last_html:
            # Restore from the original (unmodified) HTML — no re-fetch needed.
            restored = unblock_remote_images(
                block_remote_images(self._last_html))  # get blocked version first
            # Simpler: just reload from original without blocking
            try:
                self._loading_page = True
                self._wv.SetPage(self._last_html, 'about:blank')
                log.info('[webview] images unblocked (Python restore)')
            except Exception as e:
                self._loading_page = False
                log.warning('[webview] unblock reload failed: %s', e)

    def get_selected_text(self) -> str:
        try:
            return self._wv.GetSelectedText()
        except Exception:
            return ''

    def get_full_text(self) -> str:
        if not self._last_html:
            return ''
        t = re.sub(r'<[^>]+>', '', self._last_html)
        t = unescape(t)
        return re.sub(r'\n{3,}', '\n\n', t).strip()

    def open_in_browser(self):
        try:
            with open(self._temp_html_path, 'w', encoding='utf-8') as f:
                f.write(self._last_html)
            webbrowser.open('file:///' + self._temp_html_path.replace('\\', '/'))
        except Exception as e:
            log.warning('[webview] open_in_browser failed: %s', e)

    def scroll(self, lines: int):
        if not self._is_edge:
            return
        try:
            px = lines * 120
            self._wv.RunScript(
                f"window.scrollBy({{top:{px},left:0,behavior:'smooth'}});")
        except Exception:
            pass

    # ── Private ───────────────────────────────────────────────

    def _load_plain_fallback(self, html: str):
        txt = html_to_readable_text(html)
        escaped = txt.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        plain = ('<html><body><pre style="white-space:pre-wrap;font-family:sans-serif;'
                 'font-size:14px;padding:16px;">' + escaped + '</pre></body></html>')
        try:
            self._wv.SetPage(plain, 'about:blank')
        except Exception:
            pass

    # ── Properties ────────────────────────────────────────────

    @property
    def renderer_name(self) -> str:
        return RENDERER

    @property
    def is_html_capable(self) -> bool:
        return True

    @property
    def has_size_limit(self) -> bool:
        return False

    @property
    def is_showing_preview(self) -> bool:
        return False
