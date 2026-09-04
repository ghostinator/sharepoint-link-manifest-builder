# Threat model

Scope: the desktop application, its local state, its use of Microsoft identity and Microsoft
Graph, and its release/distribution path.

Method: asset-centric, with STRIDE used as a checklist. Each threat lists the mitigation
implemented in this repository and any residual risk.

---

## 1. Assets

| Asset | Why it matters |
|---|---|
| Access and refresh tokens | Bearer credentials for the user's Microsoft 365 data |
| Authorization codes / PKCE verifier | Short-lived, exchangeable for tokens |
| Tenant configuration (tenant ID, client ID) | Not secret, but identifies the customer |
| Sharing links produced by a job | Grant access to real content — sometimes anonymously |
| Manifests written into SharePoint | Become an authoritative index for an AI system |
| Local job history, audit and diagnostics | Contain filenames, paths, URLs, identities |
| The Entra app registration | Controls what the product may do in the tenant |
| The release artifacts | Executed on administrator machines |

## 2. Trust boundaries

1. User <-> application UI
2. Application <-> system browser (authentication and consent)
3. Application <-> OS secure storage
4. Application <-> Microsoft identity platform / Microsoft Graph (TLS)
5. Application <-> local filesystem (exports, logs, bundles)
6. Repository <-> GitHub / release consumers

---

## 3. Threats and mitigations

### T1 — Token theft from local storage
*Spoofing / Information disclosure.*
**Mitigation:** MSAL cache in OS-native secure storage (DPAPI, Keychain, Secret Service). The
store is probed at startup with a real round-trip; if unavailable the app tells the user and
falls back to **memory-only**, never plaintext (ADR-0008). Tokens are never written to
`settings.json`, job history, exports, or diagnostic bundles.
**Residual:** a compromised user account on the local machine can read that user's own secure
store. Out of scope for an application-level control.

### T2 — Token or authorization-code leakage through logs
*Information disclosure.*
**Mitigation:** a `RedactingLogger` decorator passes every message and scope value through
`SensitiveDataRedactor`, which strips bearer tokens, `Authorization` headers, and
`code` / `id_token` / `access_token` / `refresh_token` query parameters. Redaction is unit
tested. Request URLs are logged with sensitive query values removed. No log level, including
`Trace`, bypasses redaction.

### T3 — Consent phishing / a fake Microsoft sign-in screen
*Spoofing.*
**Mitigation:** authentication and consent **always** open the system browser at an official
`login.microsoftonline.com` URL. There is no embedded web view (ADR-0004), the application
never renders a consent-like screen, and it never accepts credentials in any field.
**Detection aid:** the Permissions tab in Settings shows the exact client ID and tenant in use so a user
can compare against what the browser displayed.

### T4 — Cross-tenant confusion
*Spoofing / Elevation.*
**Mitigation:** all identity requests are **tenant-specific** (never `/common`). The tenant
returned by the sign-in and consent flows is compared against the configured tenant, and a
mismatch is a hard failure with an explicit message. Manifests record the tenant they were
generated against.

### T5 — CSRF / replay against the consent redirect
*Tampering.*
**Mitigation:** a cryptographically random `state` is generated per consent attempt, held in
memory, and compared on return; a mismatched or missing `state` is rejected. The loopback
listener binds `127.0.0.1` on an ephemeral port and accepts exactly one request.

### T6 — Over-sharing (the product's most consequential risk)
*Elevation of privilege, by design misuse.*
This tool creates real sharing links in bulk. An `anonymous` link on the wrong folder is a data
breach.
**Mitigations:**
- Dry run performs no content or permission modification, and is the recommended first step.
- Preview shows target count, candidate files, requested permission, audience and recipients
  before anything runs; execution requires an explicit **Start**.
- "Anyone with the link" is visibly flagged as the highest-risk audience and is subject to
  tenant policy, which the app never claims to override.
- Overlapping targets are detected and deduplicated so a parent selection does not silently
  process far more than intended.
- Every result records the *actual* outcome (`Created`, `Reused`, `PolicyBlocked`,
  `AccessDenied`), so a blocked link is never reported as a success.
**Residual:** an authorized user can still deliberately over-share. This is a governance
control, not a technical one; the audit trail and manifests make it visible.

### T7 — Manifest poisoning
*Tampering.*
A manifest is consumed by an AI system, so its contents are effectively instructions-adjacent.
**Mitigation:** manifests are only ever *written* from data the app itself produced. When
*reading* a manifest for update mode, content is parsed defensively: unknown lines are ignored,
entries are matched on `(driveId, itemId)` (ADR-0007), and no field is executed or interpreted.
Foreign files at a manifest path are never overwritten by default — the app writes a
timestamped version instead (ADR-0010).

### T8 — Accidental destructive manifest overwrite
*Tampering / Denial of service.*
**Mitigation:** ETag `If-Match` on every update; `412` surfaces as a typed conflict and is
never blindly retried. Default policy is `UpdateSafely` only when the app itself generated the
file and it parses; otherwise `CreateTimestampedVersion`.

### T9 — CSV formula injection
*Tampering, executing on the victim's spreadsheet.*
A filename such as `=cmd|'/c calc'!A1.docx` is attacker-controlled data that lands in a CSV.
**Mitigation:** `CsvSanitizer` prefixes any cell beginning with `=`, `+`, `-`, `@`, tab or
carriage return with a single quote, and quotes/escapes per RFC 4180. Unit tested.

### T10 — Malicious filenames and untrusted metadata in the UI and in Markdown
*Tampering / XSS-equivalent.*
**Mitigation:** Avalonia renders text, not HTML, so there is no script execution surface.
Markdown manifests escape untrusted content (`MarkdownEscaper`), and JSON is written through
`System.Text.Json` with strict encoding.

### T11 — Path traversal and symlink abuse on local export
*Tampering.*
**Mitigation:** `SafePathBuilder` rejects absolute paths, `..` segments, rooted or
device-prefixed paths, and reserved Windows device names, then verifies the resolved full path
is still inside the intended directory. Local exports never silently overwrite; the user
confirms.

### T12 — Hostile pasted URLs
*Spoofing / SSRF-adjacent.*
**Mitigation:** pasted URLs are parsed, not fetched. Only `https` is accepted; the host must be
a SharePoint/OneDrive-shaped host; resolution happens exclusively through Microsoft Graph
(`/sites/{host}:/{path}` or `/shares/{token}`), never by requesting the pasted URL directly.
The app therefore cannot be induced to make an arbitrary outbound request.

### T13 — Stale links and stale manifests
*Integrity over time.*
A manifest can outlive the permissions it describes.
**Mitigation:** every entry is timestamped and carries its status; manifests carry a generation
timestamp and job ID; update mode can mark or remove entries whose files have disappeared.
**Residual:** the app cannot know that an administrator later revoked a link. Documented.

### T14 — Bootstrap application abuse
*Elevation of privilege.*
The bootstrap identity can create app registrations in a customer tenant.
**Mitigation:** it is a separate identity from normal operation; the default tier is
**create-only** (`AppRegistration.Create`), not tenant-wide app write (ADR-0005, ADR-0009); it
never touches SharePoint content; every tenant modification is previewed and audited; and it is
visible in the tenant as its own enterprise application, so an administrator can review or
block it.

### T15 — Unauthorized registration deletion
*Denial of service.*
**Mitigation:** deletion is never the default, never happens at uninstall, requires an
app-created (not manually supplied) registration, requires typing the display name, warns that
other installations may depend on it, and is audited.

### T16 — Excessive directory permissions requested "just in case"
*Elevation of privilege.*
**Mitigation:** tiered, opt-in scopes (ADR-0009). Broad scopes are labelled broad in the UI with
the least-privilege alternative named. `Directory.*Write*.All` is never requested.

### T17 — Compromised dependency
*Supply chain.*
**Mitigation:** central package management pins every version in one file; `NuGetAudit` is on
at `low` for all dependencies; CI runs dependency review and generates an SBOM; Dependabot is
configured.

### T18 — Malicious or tampered release artifact
*Supply chain.*
**Mitigation:** release workflow publishes SHA-256 checksums and an SBOM. Artifacts are
**clearly labelled unsigned** — this repository has no signing credentials and never claims
signing that did not occur.
**Residual:** unsigned binaries trigger OS warnings and cannot be verified as publisher-issued.
Documented in `docs/PACKAGING.md`.

### T19 — Secret exposure through the public repository
*Information disclosure.*
**Mitigation:** `.gitignore` excludes tokens, caches, certificates, generated manifests and
diagnostics *before the first commit*; a publication safety gate scans the working tree **and
git history** for tokens, client IDs, tenant URLs, sharing links, UPNs, private keys and
home-directory paths; CI runs secret scanning and CodeQL. The gate fails closed and the
documentation states that exposed credentials must be revoked, not merely deleted.

### T20 — Diagnostic bundle leaking tenant data
*Information disclosure.*
**Mitigation:** the bundle excludes tokens, secrets, authorization headers, private sharing
URLs and, unless explicitly approved by the user, filenames, email addresses and full
tenant-specific URLs. The exact categories to be included are shown **before** export.

### T21 — Telemetry leaking customer data
*Information disclosure.*
**Mitigation:** telemetry is **disabled by default and not implemented as a live pipeline**.
The opt-in architecture exists, but no filename, URL, tenant ID, identity, or sharing link is
ever eligible for transmission.

### T22 — Denial of service against the tenant by aggressive enumeration
*Denial of service, self-inflicted.*
**Mitigation:** bounded concurrency (default 4), optional inter-request delay, strict
`Retry-After` compliance, exponential backoff with jitter, and lazy tree loading so the app
never enumerates a whole tenant to draw a screen.

---

## 4. Explicitly out of scope

- A compromised operating system or a local attacker with the user's own privileges.
- Malicious behaviour by an authorized administrator (governance, not a software control).
- Microsoft 365 service-side vulnerabilities.
- Content classification — the app does not inspect file contents.

## 5. Reporting

See `SECURITY.md`. Do not open a public issue containing tokens, sharing links, tenant
identifiers, or customer data.
