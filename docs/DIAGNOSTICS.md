# Diagnostics

## The Diagnostics tab in Settings

### Environment

Application version, operating system, architecture, .NET runtime, connection state,
registration and consent status, and secure-storage availability.

### Connectivity test

Issues one minimal Graph request (`GET /me?$select=id`) and reports the round-trip time. This
distinguishes "not signed in" from "signed in but Graph is unreachable" — a distinction that
matters when a proxy or firewall is involved.

### Recent errors

Sanitized failures from recent jobs: the normalized error kind and message, never a raw response
body.

### Log folder

Opens the local log directory. Every logging provider is wrapped in a redaction filter, so no
bearer token, JWT, authorization header or credential query parameter reaches a log file.

Logs roll at 5 MB and five previous files are retained, so they cannot grow without bound.

---

## The diagnostic bundle

A ZIP written to a local file. **Nothing is uploaded.** You decide whether to send it anywhere.

### Built from an allow-list

The bundle is assembled by explicitly adding approved categories, not by collecting everything
and then stripping what looks sensitive. Exclusion-based redaction is how tenant data ends up in
support tickets.

The page shows the categories **before** you export.

### Always included

- Application version, platform, architecture and .NET runtime
- Whether a tenant is configured
- Registration source, consent state and consent type
- Required, granted and missing scopes
- Secure-storage availability and mechanism
- Counts from the most recent job
- Application settings, excluding any identifier
- Sanitized recent errors: HTTP status, Graph error code, correlation identifier

Tenant and client identifiers are **masked** by default (`1234****9abc`). They are not secret,
but they identify your organization, and a bundle is meant to be shareable.

### Never included, under any option

- Access tokens, refresh tokens, authorization codes
- Authorization headers
- Client secrets, certificates, passwords
- Sharing links produced by any job
- File contents

### Included only if you tick the box

- File and folder names
- Email addresses and user principal names
- Full tenant-specific URLs and identifiers

Each is off by default. Turn one on only when the person helping you actually needs it.

### Contents

| Entry | What it holds |
|---|---|
| `summary.txt` | Environment, connection, secure storage, settings, latest job |
| `included-categories.txt` | Exactly which categories are and are not present |
| `recent-errors.txt` | Sanitized errors from recent jobs |

`included-categories.txt` exists so the recipient can see what they were given without having to
infer it.

---

## Correlation identifiers

Every Graph request carries a `client-request-id` GUID, and the service's `request-id` is
captured from the response. Both appear in error details and in the bundle.

These are the most useful thing you can supply when reporting a Graph problem: they let a
maintainer, or Microsoft support, find the exact request in tenant-side telemetry.

---

## Reading the logs

```
2026-09-02T14:15:30Z [INF] SharePointLinkManifestBuilder.Graph.Http.GraphApiClient:
  Graph GET https://graph.microsoft.com/v1.0/drives/.../children?[REDACTED] (attempt 1, correlation 3f2a...)
```

- Timestamps are UTC and ISO 8601.
- Levels: `TRC`, `DBG`, `INF`, `WRN`, `ERR`, `CRT`.
- Query strings are replaced with `[REDACTED]`, so no parameter of any kind reaches a log.

Raise the log level on the Settings page (it takes effect at next launch). Even at `Trace`,
redaction still applies — there is no log level that bypasses it.

---

## Clearing local data

| Command | Effect | Signs you out? |
|---|---|---|
| Clear local cache | Deletes exported bundles and scratch data | No |
| Clear job history | Deletes all history entries | No |
| Forget account | Removes the cached account | Yes |
| Clear token cache | Removes all cached tokens | Yes |
| Remove tenant configuration | Removes local settings and tokens. Your tenant is untouched. | Yes |

*Clear local cache* deliberately does not touch the token cache, so it cannot sign you out as a
side effect.

---

## Reporting a problem

Include: application version, platform, what you did, what happened, what you expected, the
error message, and the correlation ID.

Never include: tokens, secrets, sharing links, real SharePoint or OneDrive URLs, tenant or client
identifiers, email addresses, or an unreviewed diagnostic bundle.

See [SUPPORT.md](SUPPORT.md).
