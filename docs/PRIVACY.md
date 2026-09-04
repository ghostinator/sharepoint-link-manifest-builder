# Privacy

## Summary

This application sends data to **Microsoft identity and Microsoft Graph endpoints only**. It
contacts no other service. There is no telemetry, no analytics, no crash reporting and no update
ping unless you press *Check for updates*.

Everything else stays on your machine.

## What is stored locally

Under the platform application-data directory. The exact paths are shown on the Settings page.

| Location | Contents | Credentials? |
|---|---|---|
| `settings.json` | Theme, defaults, retention, log level | No |
| `tenant.json` | Tenant ID, client ID, consent state, required and granted scopes | No |
| `profiles/` | Saved job configurations | No |
| `history/` | Job history: counts, targets, manifest locations, sanitized errors | No |
| `audit/` | Record of changes made to your Microsoft Entra tenant | No |
| `logs/` | Application logs, passed through redaction | No |
| `cache/` | MSAL token cache | **Yes**, protected by the OS |
| `exports/` | Diagnostic bundles you exported | Only what you approved |

### Sign-in details

Held by the operating system's secure store:

| Platform | Mechanism |
|---|---|
| Windows | DPAPI-protected file, readable only by your Windows account |
| macOS | Keychain |
| Linux | Secret Service / libsecret keyring |

If secure storage is unavailable, the application says so and keeps tokens **in memory only**.
It never writes them to disk unprotected. You will sign in again each launch.

### Job history

Records what a job did, never a credential. The account is recorded as a privacy-conscious
identifier (`user@yourdomain.com` reduced to the domain), not the full user principal name.

Retention is configurable on the Settings page, defaulting to 100 entries. Clear it at any time
from Settings or Diagnostics.

## What is sent, and where

| Destination | What | When |
|---|---|---|
| `login.microsoftonline.com` (or your sovereign-cloud equivalent) | Authentication and consent, in your system browser | When you sign in or consent |
| `graph.microsoft.com` | Requests to read sites, folders and files; create sharing links; write manifests | While you use the application |

Nothing else. No third-party analytics, error-reporting or metrics service is contacted, and
none is present in the dependency set.

## Telemetry

**Disabled, and not implemented.** The Settings page has a telemetry opt-in so the architecture
is explicit about consent, but no telemetry pipeline exists in this build. Even were one added,
the following would never be eligible for transmission:

- File or folder names
- SharePoint or OneDrive URLs
- Sharing links
- Tenant identifiers or names
- User identities or email addresses
- File contents

## Diagnostic bundles

Built from an explicit allow-list, and the Diagnostics tab in Settings shows you the categories before
anything is written.

**Always included:** application version, platform, runtime; whether a tenant is configured;
consent and registration status; secure-storage availability; sanitized recent errors reduced to
HTTP status, Graph error code and correlation identifier; counts from the most recent job;
settings excluding any identifier.

**Never included, under any option:** access tokens, refresh tokens, authorization codes,
authorization headers, client secrets, certificates, passwords, sharing links, file contents.

**Included only if you explicitly tick the box:** file and folder names; email addresses and user
principal names; full tenant-specific URLs and identifiers.

A bundle is written to a local file. Nothing is uploaded. You decide whether to send it anywhere.

## Logging

Every logging provider is wrapped in a redaction filter, so no call site — including a
third-party library logging through the same pipeline — can write a bearer token, JWT,
authorization header or credential query parameter to a log. Request URLs are logged with their
query strings removed.

Log level is configurable and defaults to Information. Logs roll by size and a bounded number of
previous files is kept.

## Data you create in Microsoft 365

Sharing links and manifest files created by this application live in **your** Microsoft 365
tenant and are governed by your organization's policies and retention rules, not by this
application. Removing the application does not remove them.

A manifest is an index of links to real content. Treat it with the same care as the content it
points to.

## Your controls

| Control | Where |
|---|---|
| Forget account | Settings |
| Clear token cache | Settings |
| Clear cached data | Settings or Diagnostics |
| Remove tenant configuration | Settings or Permissions |
| Clear job history | Settings, Diagnostics or Job History |
| Set history retention | Settings |
| Open the data folder | Settings |
| Choose diagnostic bundle contents | Diagnostics |

## Children

Not directed at children and not intended for use by anyone under 16.

## Changes

Material changes to this document will be noted in `CHANGELOG.md`.

## Contact

`PLACEHOLDER-PRIVACY@example.invalid`

> Placeholder. A publisher replaces this, and the privacy policy URL shown on the consent
> screen, before distribution.
