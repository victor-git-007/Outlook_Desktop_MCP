# Outlook Desktop MCP

**by Disrup Technology — https://disrup-tech.com**

A one-command setup that connects Microsoft Outlook Desktop (Classic) to your AI
assistants (Claude Desktop and Hermes Agent) as two MCP servers — one for
calendar, one for email — and applies two bug-fixes Disrup Technology developed
for non-US Windows date locales.

> Requirements: Windows + **Outlook Desktop (Classic)** running. No Microsoft
> Graph, no Entra app registration — pure local COM automation.

---

## What you get

| Capability | Tools |
|---|---|
| **Calendar** | List / create / update / delete events, find free slots, attendee status |
| **Email** | Send / read / flag / move mail, rules, out-of-office, attachments |

After setup, just ask your assistant naturally:
- "What's on my calendar tomorrow?"
- "Send an email from my disrup-tech.com account to …"
- "List my flagged (red-flag) follow-up emails"

---

## The two Disrup fixes

1. **Calendar date parse** — on locales such as `en-AE` (short date
   `dd/MM/yyyy`) the calendar server throws *"End date cannot be before start
   date"*. Our patch forces `MM/DD/YYYY` parsing.
2. **Flagged-email filter** — the email server's built-in follow-up-flag filter
   throws a COM error on some Outlook installs. Our patch replaces it with a
   safe in-process check and adds a `flagged` field to every email summary.

Both ship as `.patch` files under `patches/` and are applied automatically by
`setup.ps1`.

---

## Quick start (Windows, run as Administrator)

```powershell
cd Outlook_Desktop_MCP
.\setup.ps1
```

The script will:
1. Install the calendar server (npm).
2. Create an isolated Python environment and install the email server
   (pinned so it stays compatible).
3. Apply the two Disrup patches.
4. Back up and update your Claude Desktop and Hermes configs.

Restart Claude Desktop (or start a new Hermes session) when it finishes.

---

## Manual setup

If you prefer to edit configs yourself, copy the snippets from
`config-examples/`:

- `claude_desktop_config.json` → merge into `%APPDATA%\Claude\claude_desktop_config.json`
- `hermes_config.yaml` → add under `mcp_servers:` in `%LOCALAPPDATA%\Hermes\config.yaml`

Apply the patches manually with any `patch` tool:
```powershell
python -m pip install patch
python -c "import patch; patch.fromfile('patches/outlook-desktop/flagged-only-python-filter.patch').apply(strip=1, root=r'VENV\Lib\site-packages\outlook_desktop_mcp')"
```

---

## Usage notes

- Always pass the `account` parameter (e.g. `info@disrup-tech.com`). The
  default store can resolve to an empty profile and return nothing.
- Red follow-up flags are returned by `list_emails(flagged_only=True, account="...")`.

---

## License

MIT — covers all installer code, patches, and documentation in this repository.
The underlying server packages retain their own upstream licenses; this project
does not modify or redistribute their source.

© Disrup Technology — https://disrup-tech.com
