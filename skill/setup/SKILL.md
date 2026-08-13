---
name: outlook-triage-setup
description: Install and configure the outlook-triage tool — Python CLI plus Microsoft Entra app registration, OAuth device-code flow, first-time authentication, adding accounts, and troubleshooting auth/access errors. Use when the user asks to install or set up outlook-triage, add another Outlook/Microsoft 365 account, or when authentication fails (403 access_denied, invalid_client, missing token, "No such option", etc.). Triages mail only after setup; this skill is about setup, not triage.
---

Installs the `outlook-triage` CLI and wires up Microsoft Entra OAuth so it can read the user's Outlook / Microsoft 365 mail **read-only**. The agent runs every shell step itself; it guides the user click-by-click through the parts that require their Microsoft account and browser, and waits for their handoff at those points.

## Guardrail

**Read-only, always.** This skill sets up a read-only tool (`Mail.Read` delegated permission only). Never use it to send, modify, archive, delete, or label mail. `credentials.json` and the `.outlook-triage/` directory are secrets — never print, echo, or commit their contents.

## Scope

Setup only. Once auth works, stop: classification/triage is the `outlook-triage` skill's job. End by pointing the user there.

## Process

### 1. Locate the project

State dir is `<cwd>/.outlook-triage`, so **every command must run from the project dir**. Default is `C:\Users\a956064\outlook-triage` (confirm with the user if unsure). If the project isn't there, ask where they cloned it.

### 2. Check the state

- `credentials.json` exists in the project dir?
- `.outlook-triage/<account>/token.json` exists for the account they want?
- Command available? (`outlook-triage --help`)

Report what's present before starting; only do the steps that are missing.

### 3. Microsoft Entra admin center — guided, one click at a time

The user must do these in their browser (the agent has no Microsoft access). Walk them **one step at a time**, pausing for a clear "done" before continuing. Register a new app, or reuse an existing one.

1. **Create/select an app registration** at https://entra.microsoft.com/ → **Identity** → **Applications** → **App registrations** → **New registration**.
2. **Name** it e.g. `outlook-triage` and set **Supported account types** to **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)** — this is required so the app works with both Microsoft 365 work/school mailboxes and personal `outlook.com`/`hotmail.com` accounts. Leave **Redirect URI** empty. Click **Register**.
3. **Copy the Application (client) ID** from the overview page. Save it into the project dir as `credentials.json`:
   `{"client_id": "<that-id>"}`
4. **Add the permission**: **Manage → API permissions → Add a permission → Microsoft Graph → Delegated permissions** → search for and check **Mail.Read** (read-only) → **Add permissions**. (`offline_access` is granted by default.)
5. **Allow public client flows** (required for the device-code flow): **Manage → Authentication →** scroll to **Advanced settings → Allow public client flows** → set toggle to **Yes** → **Save**.

**Handoff point:** verify `credentials.json` now exists in the project dir and holds a client_id before proceeding. If missing, re-prompt rather than skipping.

No client secret and no redirect URI are needed — this is a public client using the OAuth 2.0 device-code flow.

### 4. Install

```bash
pip install -e .
```

Verify with `outlook-triage --help`.

### 5. Authenticate

```bash
outlook-triage --account <name> auth-run
```

The device-code flow prints a URL and a code. **Handoff point:** give the user the URL + code; they open it, sign in with the account to add, approve the `Mail.Read` permission, and report back. The tool polls until they finish, then saves the token to `.outlook-triage/<name>/token.json`. No browser automation is needed — the user signs in on any device.

Default account name is `default` (`outlook-triage auth-run`); use `--account <name>` before the command for any other name.

### 6. Add more accounts

Repeat step 5 with a new `--account <name>`. Multi-tenant + personal accounts means any Microsoft account can authorize against the same app — no test-user allowlisting (unlike Google).

### 7. Verify

```bash
outlook-triage --account <name> classify --hours 24
```

Must return a categorized message list with no auth errors. Then hand off to the `outlook-triage` skill for actual triage.

## Troubleshooting

| Symptom | Cause → Fix |
|---|---|
| `AADSTS7000218` / `invalid_client` | The app isn't configured as a public client. Enable **Allow public client flows** (step 3.5) and retry. |
| `AADSTS65001` (consent not granted) | `Mail.Read` permission missing or not consented. Re-check step 3.4 and re-run `auth-run`. |
| `403 access_denied` on fetch | Token lacks `Mail.Read`; re-authorize. Also confirm the account is a mailbox with the permission actually granted (API permissions → **Grant admin consent** if the tenant requires it). |
| `No 'client_id' found` / `No app credentials found` | `credentials.json` missing/malformed. Back to step 3.3. |
| `No such option '--account'` | `--account` is a group option — it goes **before** the command: `outlook-triage --account work auth-run`, not `outlook-triage auth-run --account work`. |
| `outlook-triage: command not found` | Not installed. Run step 4. |
| Auth never completes / times out | Device code expired (codes expire after ~15 min) or the wrong account signed in. Re-run `auth-run`. |
| `Authentication failed` with a specific code | Read the `error_description` — e.g. `AADSTS70011` (bad request) vs `AADSTS700016` (client not found — wrong client_id). |
| Wrong mailbox / looks like another account | `.outlook-triage/<account>/` holds one token per account; each account needs its own `auth-run` (step 6). |
| Refresh failures after a long idle | Delete `.outlook-triage/<account>/token.json` and re-run `auth-run` for that account. |

## Completion criterion

The chosen account(s) return live mail via `classify --hours 24` with no auth errors, the user knows how to add another account, and the user is pointed to the `outlook-triage` skill for triage.
