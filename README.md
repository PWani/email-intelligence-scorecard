# Email Intelligence Scorecard

A Windows desktop email client that unifies **Microsoft 365 (Outlook)** and **Google Gmail** into a single dashboard with intelligent email scoring, priority ranking, and keyboard-driven workflows.

**Created by [Praneeth Wanigasekera](https://github.com/PWani)**

![Python](https://img.shields.io/badge/python-3.10+-blue)
![wxPython](https://img.shields.io/badge/UI-wxPython-green)
![License](https://img.shields.io/badge/license-MIT-brightgreen)

---

## Features

- **Unified Inbox** — Microsoft Graph API and Gmail API side-by-side in one view
- **Intelligent Scoring Engine** — Configurable rule-based email prioritization (urgent / important / normal / low)
- **VIP Senders** — Boost scores for key contacts
- **Keyword Detection** — Urgent, action-required, and calendar keywords detected in subject and body
- **Auto-Archive** — Suppress known noise senders automatically
- **Conditional Rules** — Score emails based on sender + recipient combinations
- **Compose with Signatures** — HTML and plaintext email composition with configurable signature blocks
- **Calendar Integration** — View and manage meetings alongside email
- **Send Later / Snooze / Remind** — Time-shifted email workflows
- **Undo Send** — Configurable delay window before emails are actually dispatched
- **Keyboard-First** — Full keyboard navigation (j/k, Enter, r, a, f, c, and more)
- **Spell Check** — Inline spell checking in compose
- **Offline Queue** — Queued actions dispatched when connectivity returns

---

## Setup

### Prerequisites

- Python 3.10+
- Windows 10/11 (wxPython + WebView2 dependency)

### Install Dependencies

```bash
pip install wxPython msal requests google-auth google-auth-oauthlib google-api-python-client pyspellchecker
```

### Microsoft 365

Each Microsoft 365 organization (tenant) is a walled garden. Users need an Azure AD app registration that matches their tenant.

**If you're setting this up for your own org:**

1. Go to [Azure AD App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and create a new registration
2. Set the platform to **Mobile and desktop applications** with redirect URI: `http://localhost:8400`
3. Add **Delegated** API permissions: `User.Read`, `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `People.Read`, `Contacts.Read`
4. Copy the **Application (client) ID**
5. Launch the app — it will prompt for the client ID on first run

**If you're distributing to multiple orgs**, each organization's admin will need to register their own app in their own Azure AD tenant (or you can register a multi-tenant app and have each org grant admin consent). Set `authority` to your tenant:

```
https://login.microsoftonline.com/YOUR_TENANT_ID
```

The default `https://login.microsoftonline.com/common` supports personal and work/school accounts across tenants.

### Google Gmail

Each user sets up their own Google OAuth credentials. **Do not share or commit `google_credentials.json`** — it contains your OAuth client secret.

1. Create a project in [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the **Gmail API**, **People API**, and **Google Calendar API**
3. Go to **Credentials** → Create **OAuth 2.0 Client ID** (application type: **Desktop app**)
4. Download the credentials JSON file
5. Place it at `~/.outlook_dashboard/google_credentials.json`, or enable Google in the app's **Settings** and browse to the file when prompted
6. Sign in with your Google account in the browser popup that opens

Both Microsoft and Google accounts can be connected simultaneously. The inbox merges emails from both with a filter toggle.

### Run

```bash
python -m email_dashboard
```

---

## Configuration

Config is stored at `~/.outlook_dashboard/config.json`. Key settings:

| Setting | Description | Default |
|---|---|---|
| `client_id` | Microsoft Azure AD app client ID | *(prompted on first run)* |
| `authority` | Microsoft login authority URL | `common` |
| `google_enabled` | Enable Gmail integration | `false` |
| `google_credentials_file` | Path to Google OAuth credentials JSON | |
| `emails_per_page` | Emails loaded per page | `30` |
| `undo_send_seconds` | Delay before send is committed | `60` |
| `signature_company` | Company name for email signature | |
| `signature_address1` | Address line 1 for signature | |
| `signature_address2` | Address line 2 for signature | |
| `signature_website` | Website URL for signature | |

Your name, title, email, and phone are pulled automatically from your Microsoft/Google account profile. The signature fields above are optional additions you can set by editing `config.json` directly.

---

## Scoring Rules

Edit via **Settings → Scoring Rules** in the app, or directly in `~/.outlook_dashboard/scoring_rules.json`.

The scoring engine evaluates each email against configurable rules and assigns a composite score mapped to four priority levels:

| Score | Priority | Color |
|---|---|---|
| **75+** | Urgent | 🔴 Red |
| **55–74** | Important | 🟠 Orange |
| **35–54** | Normal | 🔵 Blue |
| **< 35** | Low | ⚪ Gray |

### Rule Categories

- **VIP Senders** — Boost score when specific people or domains appear as sender or recipient
- **Urgent / Important Keywords** — Pattern matching in subject and body (e.g., "deadline", "approval", "ASAP")
- **Conditional Rules** — Compound logic like "sender contains X AND addressed directly to me"
- **Static Signals** — Unread, flagged, high-importance header, direct-to vs CC, attachments
- **Recency** — Recent emails score higher; emails older than 7 days get a penalty
- **Calendar Keywords** — Detect meeting invites and scheduling language
- **Question Detection** — Emails that ask you a question by name get a boost
- **Automated Sender Suppression** — noreply, notifications, newsletters scored down
- **Auto-Archive** — Completely suppress known noise senders from the inbox

All rules, keywords, and thresholds are fully user-editable through the in-app rules editor.

---

## Keyboard Shortcuts

| Key | Action |
|---|---|
| `↑` / `↓` | Navigate emails (list) or scroll body (detail) |
| `r` | Reply |
| `R` | Reply All |
| `a` / `A` | Archive |
| `f` / `F` | Forward |
| `s` / `S` | Snooze |
| `m` / `M` | Remind |
| `Delete` | Delete |
| `n` / `N` | Compose new email |
| `Ctrl+K` | Command palette |

---

## Project Structure

```
email_dashboard/
├── __main__.py              # Entry point
├── app.py                   # Main app class (mixin composition)
├── core/
│   ├── config.py            # Configuration, theme, constants
│   ├── auth.py              # MSAL authentication (Microsoft)
│   ├── graph_client.py      # Microsoft Graph API client
│   ├── google_client.py     # Gmail + Google Calendar API client
│   ├── email_intelligence.py # Scoring engine
│   └── spell_checker.py     # Spell check
├── ui/
│   ├── build.py             # Main window layout
│   ├── auth_ui.py           # Authentication UI flows
│   ├── compose.py           # Email composition
│   ├── detail_view.py       # Email reading pane
│   ├── list_render.py       # Email list rendering
│   ├── actions.py           # Email actions (archive, delete, snooze, etc.)
│   ├── settings.py          # Settings panel
│   ├── train_rules.py       # Scoring rules editor
│   ├── meetings.py          # Calendar / meetings view
│   ├── keyboard.py          # Keyboard shortcuts
│   ├── email_loading.py     # Email fetch / sync
│   ├── attachments.py       # Attachment handling
│   ├── autocomplete.py      # Address autocomplete
│   ├── utils.py             # UI utilities
│   └── webview_widget.py    # HTML email rendering (WebView2)
├── assets/
│   ├── scoring_rules.json   # Default scoring rules
│   ├── email_dashboard.ico  # Application icon
│   └── logo.png             # Logo
```

---

## License

MIT — see [LICENSE](LICENSE).
