# outlook-triage

Read Microsoft 365 / Outlook mail **read-only** and classify messages for AI agents.

A mirror of [`gmail-triage`](../gmail-triage): a deterministic rule pass flags a
preliminary category for every message, and the agent finishes classification
against the rubric in `skill/RUBRIC.md`. Uses the Microsoft Graph API with the
`Mail.Read` (read-only) delegated permission via OAuth 2.0 **device code flow** —
no client secret, no redirect URI.

## Install

```bash
pip install -e .
```

## First-time auth

Register an app (one-time, 5 minutes — see `skill/setup/SKILL.md` for the
guided flow), save the **Application (client) ID** in `credentials.json`:

```json
{"client_id": "<your-application-client-id>"}
```

then:

```bash
outlook-triage auth-run
```

The device code flow prints a URL + code; open it, sign in, and the token is
saved to `.outlook-triage/<account>/token.json`.

## Usage

```bash
# list the last 24h from the Inbox
outlook-triage inbox --hours 24

# classified JSON for the agent pass (read-only)
outlook-triage classify --json --since 2026-08-01 --until 2026-08-14

# human digest grouped by category
outlook-triage report --format md
```

Common flags:

- `--hours N` / `--since YYYY-MM-DD` / `--until YYYY-MM-DD` — the window
- `--unread` — unread only
- `--folder <name>` — `inbox` (default), `junkemail`, `sentitems`, `drafts`,
  `deleteditems`, `archive`, `allitems`, `focused`, `other`, or a custom folder
  display name
- `--full` — fetch full message bodies (default: `bodyPreview` snippet only)
- `--query '<OData filter>'` — extra OData `$filter` terms (not Gmail syntax),
  e.g. `from/emailAddress/address eq 'payments@microsoft.com'`
- `--json` — emit raw classified features (used by the agent)
- `--format md` — markdown report
- `--account <name>` — multi-account support

Read-only, always: this tool never sends, modifies, archives, deletes, or
labels mail.

## Project layout

```
src/outlook_triage/
  cli.py          commands (auth-run, inbox, classify, report, config)
  auth.py         MSAL device-code OAuth, Mail.Read scope
  graph_api.py    Microsoft Graph client + OData query builder
  models.py       Message model -> features dict
  classify/       deterministic rule pass + rubric categories
  config.py       YAML config merge (categories, rules, label hints)
```

The skills live in `skill/` (installed to `~/.config/opencode/skills/`):

- `skill/SKILL.md` — the `outlook-triage` agent skill (triage process)
- `skill/RUBRIC.md` — the 9-category classification rubric
- `skill/setup/SKILL.md` — the `outlook-triage-setup` skill (install + auth)
