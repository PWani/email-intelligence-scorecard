# Scoring engine
import json, os, re
from datetime import datetime, timedelta, timezone
from html import unescape
from .config import CONFIG_DIR, SCORING_RULES_FILE, ensure_config_dir, log

DEFAULT_SCORING_RULES = {
    "_version": 2,
    "_description": "Email prioritization scoring rules. Edit via Settings > Scoring Rules in the app.",

    "base_score": 40,

    "priority_thresholds": {
        "urgent": 75,
        "important": 55,
        "normal": 35
    },

    "vip_senders": {
        "enabled": True,
        "entries": ["gov", "irs", "sec"],
        "sender_score": 35,
        "recipient_score": 25,
        "label": "VIP"
    },

    "critical_subjects": {
        "enabled": True,
        "patterns": ["response required", "urgent"],
        "score": 40
    },

    "critical_subject_keywords": {
        "enabled": True,
        "patterns": ["tax", "sec", "irs"],
        "score": 40
    },

    "critical_sender_domains": {
        "enabled": True,
        "domains": [".gov"],
        "score": 40
    },

    "conditional_rules": [
        {
            "enabled": True,
            "name": "Auditor addressed to me",
            "sender_contains": "auditor",
            "must_be_to_recipient": True,
            "score": 40
        }
    ],

    "automated_senders": {
        "enabled": True,
        "patterns": ["noreply", "no-reply", "notifications", "mailer-daemon"],
        "score": -20
    },

    "automated_senders_extended": [
        "noreply", "no-reply", "notifications", "mailer-daemon",
        "marketing", "newsletter", "updates", "info@", "news@", "hello@",
        "support@", "team@", "digest", "campaign"
    ],

    "name_question_detection": {
        "enabled": True,
        "score": 15,
        "question_patterns": [
            "\\?\\s*$", "can you\\b", "could you\\b", "would you\\b",
            "please advise", "thoughts\\?", "your take", "what do you think",
            "let me know", "let us know", "how would", "where would",
            "where should", "who should"
        ]
    },

    "urgent_keywords": {
        "enabled": True,
        "keywords": [
            "urgent", "asap", "immediately", "critical", "emergency", "deadline today",
            "due today", "action required", "time sensitive", "time-sensitive",
            "high priority", "escalation", "p0", "p1", "blocker", "blocking"
        ],
        "max_score": 25,
        "per_hit": 10
    },

    "important_keywords": {
        "enabled": True,
        "keywords": [
            "important", "please review", "action needed", "action item",
            "follow up", "follow-up", "decision needed", "approval", "approve",
            "signature required", "sign off", "deadline", "due date", "deliverable",
            "milestone", "board meeting", "investor", "lp", "limited partner",
            "fund", "portfolio", "diligence", "term sheet", "loi", "closing",
            "wire", "capital call", "bank balance", "account balance", "cash"
        ],
        "max_score": 15,
        "per_hit": 5
    },

    "general_question_detection": {
        "enabled": True,
        "patterns": [
            "\\?\\s*$", "can you\\b", "could you\\b", "would you\\b",
            "please advise", "thoughts\\?", "your take", "what do you think",
            "let me know", "let us know"
        ],
        "score": 8
    },

    "low_priority_keywords": {
        "enabled": True,
        "keywords": [
            "unsubscribe", "no-reply", "noreply", "do not reply", "newsletter",
            "digest", "weekly update", "monthly update", "marketing", "promotion",
            "deal of the day", "opt out", "opt-out", "notification", "automated message"
        ],
        "max_score": 25,
        "per_hit": 8
    },

    "calendar_keywords": {
        "enabled": True,
        "keywords": [
            "meeting", "calendar", "invite", "rsvp", "schedule",
            "zoom", "teams meeting", "webex", "conference call"
        ],
        "score": 5
    },

    "static_signals": {
        "unread": {"enabled": True, "score": 5},
        "high_importance": {"enabled": True, "score": 15},
        "low_importance": {"enabled": True, "score": -10},
        "flagged": {"enabled": True, "score": 12},
        "filtered_other": {"enabled": True, "score": -15},
        "direct_to": {"enabled": True, "score": 8},
        "cc_only": {"enabled": True, "score": -5},
        "has_attachments": {"enabled": True, "score": 3},
        "short_message": {"enabled": True, "score": 4, "min_len": 10, "max_len": 200}
    },

    "recency_scores": {
        "enabled": True,
        "under_1h": 10,
        "under_4h": 6,
        "under_24h": 3,
        "over_7d": -5
    },

    "category_keywords": {
        "action": ["action", "approve", "review", "sign", "deadline"],
        "fyi": ["fyi", "for your information"]
    },

    "safe_image_senders": {
        "enabled": True,
        "entries": []
    },

    "auto_archive_senders": {
        "enabled": True,
        "entries": [
            "quarantine@messaging.microsoft.com"
        ]
    }
}


def load_scoring_rules() -> dict:
    """Load scoring rules from config file, falling back to defaults."""
    ensure_config_dir()
    if os.path.exists(SCORING_RULES_FILE):
        try:
            with open(SCORING_RULES_FILE, "r") as f:
                stored = json.load(f)
            # Merge: use stored values but fill in any missing keys from defaults
            merged = json.loads(json.dumps(DEFAULT_SCORING_RULES))
            _deep_merge(merged, stored)
            return merged
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_SCORING_RULES))


def save_scoring_rules(rules: dict):
    """Save scoring rules to config file."""
    ensure_config_dir()
    with open(SCORING_RULES_FILE, "w") as f:
        json.dump(rules, f, indent=2)


def _deep_merge(base: dict, override: dict):
    """Merge override into base, recursively for dicts."""
    for key, val in override.items():
        if key in base and isinstance(base[key], dict) and isinstance(val, dict):
            _deep_merge(base[key], val)
        else:
            base[key] = val


def strip_html(html):
    """Basic strip for scoring — removes all tags and style/script content."""
    # Remove style and script blocks entirely (content + tags)
    html = re.sub(r"<style[^>]*>.*?</style>", " ", html, flags=re.IGNORECASE | re.DOTALL)
    html = re.sub(r"<script[^>]*>.*?</script>", " ", html, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"<[^>]+>", " ", html)
    text = unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def strip_outlook_banners(html):
    """Remove Outlook/Exchange safety banners and external sender warnings before scoring.
    These banners contain words like 'important', 'caution', 'external' that pollute keyword matching.

    IMPORTANT: All table-matching patterns use [^<]*(?:<(?!/table)[^<]*)* instead of .*?
    to prevent crossing </table> boundaries. Without this, a <table> in an email signature
    can match forward through reply separators to reach banner text deep in quoted threads,
    destroying the reply separator and causing extract_latest_reply() to return the entire
    thread — leading to massive false-positive scoring."""
    if not html:
        return html
    # Helper: match content within a single table (won't cross </table> boundaries)
    _intable = r'[^<]*(?:<(?!/table>)[^<]*)*'
    # Outlook external sender banner: "You don't often get email from ... Learn why this is important"
    # This appears as a <table> with massive inline revert styles
    html = re.sub(
        r'<table[^>]*>' + _intable + r'You don.t often get email from' + _intable + r'Learn why this is important' + _intable + r'</table>',
        '', html, flags=re.IGNORECASE | re.DOTALL)
    # Catch-all for any element containing "Learn why this is important" (Microsoft safety link)
    html = re.sub(
        r'<table[^>]*>' + _intable + r'Learn why this is important' + _intable + r'</table>',
        '', html, flags=re.IGNORECASE | re.DOTALL)
    # CAUTION: This email originated from outside — match only single elements
    html = re.sub(
        r'<(?:table|div|p)[^>]*>[^<]*CAUTION:\s*This email originated from outside[^<]*</(?:table|div|p)>',
        '', html, flags=re.IGNORECASE)
    # Also catch CAUTION banners with inline styling/bold tags
    html = re.sub(
        r'<(?:div|p)[^>]*>\s*(?:<(?:strong|b|span)[^>]*>)?[^<]*CAUTION:\s*This email originated from outside[^<]*(?:</(?:strong|b|span)>)?\s*</(?:div|p)>',
        '', html, flags=re.IGNORECASE)
    # [External] or [External E-Mail] tags — various wrappers
    html = re.sub(
        r'<(?:div|p|strong|span)[^>]*>\s*(?:<(?:strong|b)>)?\s*\[External(?:\s+E-?Mail)?\]\s*(?:</(?:strong|b)>)?\s*</(?:div|p|strong|span)>',
        '', html, flags=re.IGNORECASE | re.DOTALL)
    # Outlook Safe Links / safety tips: "This sender isn't verified"
    html = re.sub(
        r'<table[^>]*>' + _intable + r'This (?:sender|message) (?:isn.t|is not|failed) (?:verified|authenticated)' + _intable + r'</table>',
        '', html, flags=re.IGNORECASE | re.DOTALL)
    # Generic external source banners - match only SINGLE elements (no nested content)
    html = re.sub(
        r'<(?:div|p|table)[^>]*>[^<]*(?:external\s+(?:email|sender|source)|sent\s+from\s+outside)[^<]*</(?:div|p|table)>',
        '', html, flags=re.IGNORECASE)
    # Also catch banners with inline bold/strong tags
    html = re.sub(
        r'<(?:div|p)[^>]*>\s*(?:<(?:strong|b)>)?[^<]*(?:external\s+(?:email|sender|source)|sent\s+from\s+outside)[^<]*(?:</(?:strong|b)>)?\s*</(?:div|p)>',
        '', html, flags=re.IGNORECASE)
    # Confidentiality / legal notices at the end (common in forwarded chains)
    html = re.sub(
        r'<(?:div|p|h5)[^>]*>\s*(?:NOTICE:|DISCLAIMER:|CONFIDENTIAL(?:ITY)?:).*?(?:(?:do not|cannot)\s+(?:review|retransmit|disclose|disseminate|copy).*?)?</(?:div|p|h5)>',
        '', html, flags=re.IGNORECASE | re.DOTALL)
    return html


def _build_word_pattern(keyword):
    """Build a regex pattern with word boundaries for a keyword.
    Keywords with leading/trailing spaces are already boundary-delimited (legacy format)."""
    kw = keyword.strip()
    if not kw:
        return None
    # Multi-word phrases: use word boundaries on edges
    return re.compile(r'\b' + re.escape(kw) + r'\b', re.IGNORECASE)


def keyword_in_text(keyword, text):
    """Check if keyword appears in text using word boundaries.
    Single words use \\b boundaries; multi-word phrases use start/end boundaries."""
    pat = _build_word_pattern(keyword)
    return bool(pat and pat.search(text))


def html_to_readable_text(html):
    """Convert HTML email to readable plain text with proper line breaks and structure."""
    if not html:
        return ""
    text = html

    # Preserve line breaks
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</br>", "\n", text, flags=re.IGNORECASE)

    # Block elements get double newlines
    for tag in ["p", "div", "h1", "h2", "h3", "h4", "h5", "h6",
                "tr", "li", "blockquote", "section", "article"]:
        text = re.sub(rf"<{tag}[^>]*>", "\n\n", text, flags=re.IGNORECASE)
        text = re.sub(rf"</{tag}>", "\n", text, flags=re.IGNORECASE)

    # Horizontal rules
    text = re.sub(r"<hr[^>]*>", "\n" + "─" * 50 + "\n", text, flags=re.IGNORECASE)

    # List items
    text = re.sub(r"<li[^>]*>", "\n  • ", text, flags=re.IGNORECASE)

    # Table cells — add spacing
    text = re.sub(r"<td[^>]*>", "  ", text, flags=re.IGNORECASE)
    text = re.sub(r"</td>", "\t", text, flags=re.IGNORECASE)
    text = re.sub(r"<th[^>]*>", "  ", text, flags=re.IGNORECASE)
    text = re.sub(r"</th>", "\t", text, flags=re.IGNORECASE)

    # Remove style and script blocks entirely
    text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.IGNORECASE | re.DOTALL)

    # Remove all remaining tags
    text = re.sub(r"<[^>]+>", "", text)

    # Decode HTML entities
    text = unescape(text)

    # Clean up whitespace while preserving newlines
    lines = text.split("\n")
    cleaned = []
    for line in lines:
        cleaned_line = re.sub(r"[ \t]+", " ", line).strip()
        cleaned.append(cleaned_line)
    text = "\n".join(cleaned)

    # Collapse 3+ newlines to 2
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def extract_text(email):
    body = email.get("body", {})
    if body.get("contentType") == "html":
        html = strip_outlook_banners(body.get("content", ""))
        return strip_html(html)
    return body.get("content", email.get("bodyPreview", ""))


def extract_latest_reply(email):
    """Extract only the most recent message from an email thread, stripping quoted replies.
    Works with both HTML and plain text emails."""
    body = email.get("body", {})
    content = body.get("content", "")
    content_type = body.get("contentType", "text")

    if content_type == "html":
        # Strip Outlook safety banners before extracting text
        content = strip_outlook_banners(content)
        # Cut before common Outlook/Gmail reply markers in HTML
        # Outlook: <div id="divRplyFwdMsg"> or <div id="appendonsend">
        # Also: border-top styled divs with From: headers, <hr> blocks
        # IMPORTANT: Find the EARLIEST match across ALL patterns, not the first
        # pattern that matches anywhere. A border-top separator near the top of
        # the email must win over a divRplyFwdMsg deep in the quoted thread.
        _html_sep_patterns = [
            r'<div\s+id\s*=\s*"divRplyFwdMsg".*',
            r'<div\s+id\s*=\s*"appendonsend".*',
            r'<div[^>]*style\s*=\s*"[^"]*border-top[^"]*"[^>]*>\s*<p[^>]*>(?:\s*<[^/][^>]*>)*\s*From:.*',
            r'<div>\s*<div[^>]*style\s*=\s*"[^"]*border-top[^"]*"[^>]*>\s*<p[^>]*>(?:\s*<[^/][^>]*>)*\s*From:.*',
            r'<hr[^>]*>\s*(?:<div[^>]*>)?\s*<p[^>]*>(?:\s*<[^/][^>]*>)*\s*From:.*',
            r'-{3,}\s*Original\s+(?:Appointment|Message)\s*-{3,}',
            r'<div[^>]*style\s*=\s*"text-align:\s*center[^"]*">\s*<hr[^>]*/?\s*>\s*</div>',
            r'<blockquote[^>]*>.*',
        ]
        earliest_pos = len(content)
        earliest_pat = None
        for pattern in _html_sep_patterns:
            match = re.search(pattern, content, flags=re.IGNORECASE | re.DOTALL)
            if match and match.start() < earliest_pos:
                earliest_pos = match.start()
                earliest_pat = pattern[:50]
        if earliest_pat is not None:
            content = content[:earliest_pos]
        text = strip_html(content)
    else:
        text = content
        # Plain text reply markers
        for pattern in [
            r'\n\s*-{3,}\s*Original (?:Message|Appointment)\s*-{3,}',  # --- Original Message/Appointment ---
            r'\n\s*_{3,}',                               # _____
            r'\nFrom:\s+\S+.*\nSent:\s+',               # From: ... Sent: ...
            r'\nOn\s+.{10,60}\s+wrote:',                 # On Mon, Jan 30... wrote:
            r'\n>',                                       # > quoted lines
        ]:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                text = text[:match.start()]
                break

    result = text.strip().lower()
    return result


class EmailIntelligence:
    def __init__(self, user_email="", user_name="", rules=None):
        self.user_email = user_email.lower()
        self.user_names = [n.lower() for n in user_name.split() if len(n) >= 2] if user_name else []
        self.rules = rules or load_scoring_rules()

    def reload_rules(self):
        """Reload scoring rules from disk."""
        self.rules = load_scoring_rules()

    def score_email(self, email):
        R = self.rules
        score = R.get("base_score", 30)
        signals = []
        text = extract_text(email).lower()
        subject = (email.get("subject") or "").lower()
        preview = (email.get("bodyPreview") or "").lower()
        latest_reply = extract_latest_reply(email)
        # Use latest reply for keyword matching to avoid false positives from quoted thread history
        combined = f"{subject} {latest_reply}"

        # ── Cancelled/declined meetings — force low priority early ──
        is_cancelled = any(subject.startswith(p) for p in ["cancelled:", "canceled:", "declined:"])
        if is_cancelled:
            return {
                "score": 15, "priority": "low",
                "summary": f"{email.get('from', {}).get('emailAddress', {}).get('name', 'Unknown')} — {preview.strip()[:120] or subject}",
                "signals": ["Cancelled/declined meeting"], "category": "meeting",
            }

        sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "").lower()
        to_addrs = [r.get("emailAddress", {}).get("address", "").lower()
                    for r in email.get("toRecipients", [])]
        cc_addrs = [r.get("emailAddress", {}).get("address", "").lower()
                    for r in email.get("ccRecipients", [])]

        # ── Static signals ────────────────────────────────────
        ss = R.get("static_signals", {})

        if ss.get("unread", {}).get("enabled", True):
            if not email.get("isRead", True):
                score += ss["unread"].get("score", 5); signals.append("Unread")

        if ss.get("high_importance", {}).get("enabled", True):
            imp = email.get("importance", "normal").lower()
            if imp == "high":
                score += ss["high_importance"].get("score", 15); signals.append("Marked high importance")

        if ss.get("low_importance", {}).get("enabled", True):
            imp = email.get("importance", "normal").lower()
            if imp == "low":
                score += ss["low_importance"].get("score", -10); signals.append("Marked low importance")

        if ss.get("flagged", {}).get("enabled", True):
            if email.get("flag", {}).get("flagStatus") == "flagged":
                score += ss["flagged"].get("score", 12); signals.append("Flagged")

        if ss.get("filtered_other", {}).get("enabled", True):
            if email.get("inferenceClassification") == "other":
                score += ss["filtered_other"].get("score", -15); signals.append("Filtered to Other")

        if ss.get("direct_to", {}).get("enabled", True):
            if self.user_email and self.user_email in to_addrs:
                score += ss["direct_to"].get("score", 8); signals.append("Direct recipient (To:)")
            elif ss.get("cc_only", {}).get("enabled", True):
                if self.user_email and self.user_email in cc_addrs:
                    score += ss["cc_only"].get("score", -5); signals.append("CC'd only")

        # ── Automated senders ─────────────────────────────────
        auto_cfg = R.get("automated_senders", {})
        is_automated_basic = False
        if auto_cfg.get("enabled", True):
            auto_pats = auto_cfg.get("patterns", [])
            if any(s in sender_email for s in auto_pats):
                score += auto_cfg.get("score", -20); signals.append("Automated sender")
                is_automated_basic = True
        # Extended automated sender check (used to suppress false-positive boosts)
        auto_ext = [s.lower() for s in R.get("automated_senders_extended", [])]
        is_automated = is_automated_basic or any(s in sender_email for s in auto_ext)
        # Also treat emails with unsubscribe keywords in body as automated/marketing
        if not is_automated and any(kw in text for kw in ["unsubscribe", "opt out", "opt-out",
                "view in browser", "email preferences", "manage preferences",
                "privacy policy", "terms of service", "update your preferences"]):
            is_automated = True

        # ── VIP senders ───────────────────────────────────────
        vip_cfg = R.get("vip_senders", {})
        is_vip_sender = False
        if vip_cfg.get("enabled", True):
            vip_entries = vip_cfg.get("entries", [])
            if any(vip in sender_name or vip in sender_email for vip in vip_entries):
                score += vip_cfg.get("sender_score", 35)
                signals.append(f"{vip_cfg.get('label', 'VIP')} sender")
                is_vip_sender = True

            # VIP on thread (recipient check)
            if not is_vip_sender and vip_cfg.get("recipient_score", 0):
                all_recipients = to_addrs + cc_addrs
                all_recipient_names = [r.get("emailAddress", {}).get("name", "").lower()
                                       for r in email.get("toRecipients", []) + email.get("ccRecipients", [])]
                if any(vip in addr or vip in name for addr in all_recipients
                       for vip in vip_entries for name in all_recipient_names):
                    score += vip_cfg.get("recipient_score", 25)
                    signals.append(f"{vip_cfg.get('label', 'VIP')} on thread")

        # ── Critical subject patterns ─────────────────────────
        crit_subj = R.get("critical_subjects", {})
        if crit_subj.get("enabled", True):
            for cs in crit_subj.get("patterns", []):
                if keyword_in_text(cs, subject):
                    score += crit_subj.get("score", 40)
                    signals.append(f"Critical subject: {cs}")
                    break

        # ── Critical subject keywords ─────────────────────────
        crit_kw = R.get("critical_subject_keywords", {})
        if crit_kw.get("enabled", True):
            for ck in crit_kw.get("patterns", []):
                if keyword_in_text(ck, subject):
                    score += crit_kw.get("score", 40)
                    signals.append(f"Critical: {ck}")
                    break

        # ── Critical sender domains ───────────────────────────
        crit_dom = R.get("critical_sender_domains", {})
        if crit_dom.get("enabled", True):
            for dom in crit_dom.get("domains", []):
                if sender_email.endswith(dom):
                    score += crit_dom.get("score", 40)
                    signals.append(f"Government/org sender: {sender_email.split('@')[-1]}")
                    break

        # ── Conditional rules ─────────────────────────────────
        for rule in R.get("conditional_rules", []):
            if not rule.get("enabled", True):
                continue
            match = True
            if "sender_contains" in rule:
                if rule["sender_contains"].lower() not in sender_email:
                    match = False
            if rule.get("must_be_to_recipient"):
                if not (self.user_email and self.user_email in to_addrs):
                    match = False
            if match:
                score += rule.get("score", 0)
                signals.append(rule.get("name", "Conditional rule"))

        # ── Name + question detection ─────────────────────────
        nq = R.get("name_question_detection", {})
        if nq.get("enabled", True) and self.user_names:
            addressed_by_name = any(name in latest_reply for name in self.user_names)
            if addressed_by_name and not is_automated:
                q_pats = nq.get("question_patterns", [])
                has_q = "?" in latest_reply or any(re.search(p, latest_reply) for p in q_pats)
                if has_q:
                    score += nq.get("score", 30)
                    signals.append("Addressed by name with question")

        # ── Urgent keywords (skip for automated/marketing senders) ──
        urg = R.get("urgent_keywords", {})
        if urg.get("enabled", True) and not is_automated:
            hits = [kw for kw in urg.get("keywords", []) if keyword_in_text(kw, combined)]
            if hits:
                score += min(urg.get("max_score", 25), len(hits) * urg.get("per_hit", 10))
                signals.append(f"Urgent: {', '.join(h.strip() for h in hits[:3])}")

        # ── Important keywords (skip for automated/marketing senders) ──
        imp = R.get("important_keywords", {})
        if imp.get("enabled", True) and not is_automated:
            hits = [kw for kw in imp.get("keywords", []) if keyword_in_text(kw, combined)]
            if hits:
                score += min(imp.get("max_score", 15), len(hits) * imp.get("per_hit", 5))
                signals.append(f"Important: {', '.join(h.strip() for h in hits[:3])}")

        # ── General questions (skip for automated/marketing senders) ──
        gq = R.get("general_question_detection", {})
        if gq.get("enabled", True) and not is_automated:
            q_pats = gq.get("patterns", [])
            if any(re.search(p, combined) for p in q_pats):
                score += gq.get("score", 8); signals.append("Contains questions/requests")

        # ── Low priority keywords ─────────────────────────────
        lp = R.get("low_priority_keywords", {})
        lp_hits = []
        if lp.get("enabled", True):
            lp_hits = [kw for kw in lp.get("keywords", []) if keyword_in_text(kw, combined)]
            if lp_hits:
                score -= min(lp.get("max_score", 25), len(lp_hits) * lp.get("per_hit", 8))
                signals.append(f"Low-priority: {', '.join(h.strip() for h in lp_hits[:3])}")

        # ── Calendar keywords ─────────────────────────────────
        cal = R.get("calendar_keywords", {})
        is_calendar = False
        if cal.get("enabled", True):
            cal_kws = cal.get("keywords", [])
            if any(keyword_in_text(kw, combined) for kw in cal_kws):
                score += cal.get("score", 5); signals.append("Meeting/calendar related")
                is_calendar = True

        # ── Recency ───────────────────────────────────────────
        rec = R.get("recency_scores", {})
        if rec.get("enabled", True):
            received = email.get("receivedDateTime", "")
            if received:
                try:
                    recv_dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
                    age_h = (datetime.now(timezone.utc) - recv_dt).total_seconds() / 3600
                    if age_h < 1:
                        score += rec.get("under_1h", 10); signals.append("< 1 hour ago")
                    elif age_h < 4:
                        score += rec.get("under_4h", 6); signals.append("< 4 hours ago")
                    elif age_h < 24:
                        score += rec.get("under_24h", 3)
                    elif age_h > 168:
                        score += rec.get("over_7d", -5)
                except Exception:
                    pass

        # ── Attachments ───────────────────────────────────────
        if ss.get("has_attachments", {}).get("enabled", True):
            if email.get("hasAttachments"):
                score += ss["has_attachments"].get("score", 3); signals.append("Has attachments")

        # ── Short message ─────────────────────────────────────
        sm = ss.get("short_message", {})
        if sm.get("enabled", True):
            min_l, max_l = sm.get("min_len", 10), sm.get("max_len", 200)
            if min_l < len(text) < max_l:
                score += sm.get("score", 4); signals.append("Short (likely needs response)")

        # ── Clamp and classify ────────────────────────────────
        score = max(0, min(100, score))

        thresholds = R.get("priority_thresholds", {})
        if score >= thresholds.get("urgent", 75):
            priority = "urgent"
        elif score >= thresholds.get("important", 55):
            priority = "important"
        elif score >= thresholds.get("normal", 35):
            priority = "normal"
        else:
            priority = "low"

        # Category
        cat_kw = R.get("category_keywords", {})
        cal_kws = R.get("calendar_keywords", {}).get("keywords", [])
        lp_kws = R.get("low_priority_keywords", {}).get("keywords", [])
        if is_calendar:
            category = "meeting"
        elif lp_hits and any(keyword_in_text(kw, combined) for kw in lp_kws):
            category = "newsletter"
        elif any(keyword_in_text(kw, combined) for kw in cat_kw.get("action", [])):
            category = "action"
        elif any(keyword_in_text(kw, combined) for kw in cat_kw.get("fyi", [])):
            category = "fyi"
        else:
            category = "general"

        # Summary
        sender_disp = email.get("from", {}).get("emailAddress", {}).get("name", "Unknown")
        clean_preview = preview.strip()[:120]
        summary = f"{sender_disp} — {clean_preview}" if clean_preview else f"{sender_disp}: {subject}"

        return {
            "score": score, "priority": priority, "summary": summary,
            "signals": signals, "category": category,
        }

    def process_emails(self, emails):
        enriched = []
        for email in emails:
            intel = self.score_email(email)
            em = {**email, "_intel": intel}
            # Pre-tag meeting requests using @odata.type already returned by Graph
            # This avoids a separate API call and makes the Accept button appear instantly
            odata_type = email.get("@odata.type", "")
            if "eventMessage" in str(odata_type):
                # It's a meeting-related message — mark as request by default
                # (the background check can refine this to cancellation/response later)
                em["_is_meeting_request"] = "request"
            enriched.append(em)
        enriched.sort(
            key=lambda e: (e["_intel"]["score"], e.get("receivedDateTime", "")),
            reverse=True,
        )
        return enriched

