# Spell checker
import requests
LANGUAGETOOL_URL = 'https://api.languagetool.org/v2/check'

class SpellChecker:
    """Calls LanguageTool API for spelling and grammar checking."""

    def __init__(self, api_url=LANGUAGETOOL_URL, language="en-US"):
        self.api_url = api_url
        self.language = language
        self._enabled = True

    def check(self, text):
        """Returns list of {'offset', 'length', 'message', 'replacements', 'rule'}."""
        if not text.strip() or not self._enabled:
            return []
        try:
            r = requests.post(self.api_url, data={
                "text": text,
                "language": self.language,
                "disabledRules": "WHITESPACE_RULE",
            }, timeout=5)
            r.raise_for_status()
            matches = r.json().get("matches", [])
            results = []
            for m in matches:
                results.append({
                    "offset": m["offset"],
                    "length": m["length"],
                    "message": m.get("message", ""),
                    "replacements": [r["value"] for r in m.get("replacements", [])[:5]],
                    "rule": m.get("rule", {}).get("id", ""),
                    "rule_desc": m.get("rule", {}).get("description", ""),
                })
            return results
        except Exception as e:
            self._last_error = str(e)
            return []

    def auto_fix(self, text):
        """Apply first suggestion for each error, return corrected text."""
        errors = self.check(text)
        if not errors:
            return text
        # Apply fixes from end to start so offsets stay valid
        errors.sort(key=lambda e: e["offset"], reverse=True)
        for err in errors:
            if err["replacements"]:
                start = err["offset"]
                end = start + err["length"]
                text = text[:start] + err["replacements"][0] + text[end:]
        return text
