# Security policy

## Reporting a vulnerability

**Do not open a public issue.** A public report puts every user at risk before a fix exists.

Report privately using either:

1. **GitHub private vulnerability reporting** — the Security tab, then *Report a vulnerability*.
   Preferred: it keeps the report, the discussion and the advisory together.
2. **Email** — `github@ghostinator.co`

> The email address above is a placeholder. A publisher replaces it before distribution. If it
> is still a placeholder, use GitHub private vulnerability reporting.

### What to include

- What the vulnerability allows an attacker to do
- Steps to reproduce, or a proof of concept
- Affected version and platform
- Any mitigation you have found

### What NOT to include

Never send a working credential. If reproducing requires one, describe the shape of it instead.

- Access tokens, refresh tokens, authorization codes
- Client secrets, certificates, private keys
- Real tenant identifiers, SharePoint URLs or sharing links
- Personal data belonging to anyone

### What to expect

| Stage | Target |
|---|---|
| Acknowledgement | 3 business days |
| Initial assessment | 10 business days |
| Fix or mitigation plan | Depends on severity; communicated after assessment |

We will tell you when a fix ships and credit you unless you prefer otherwise. Please give us a
reasonable opportunity to fix an issue before disclosing it publicly.

## Supported versions

Pre-1.0. Only the latest release receives fixes. This will change at 1.0.

## Scope

### In scope

- Token or credential disclosure, including through logs, exports or diagnostic bundles
- Authentication or consent bypass
- Cross-tenant data access
- Privilege escalation through the app registration or the bootstrap identity
- Injection reaching a user: CSV formula injection, path traversal, manifest poisoning
- The application creating a sharing link the user did not request or confirm
- Any silent modification of tenant configuration

### Out of scope

- Vulnerabilities in Microsoft 365, Microsoft Graph or Microsoft Entra — report those to
  Microsoft (<https://msrc.microsoft.com>)
- An authorized user deliberately over-sharing content they already control; that is a
  governance matter, not a software defect
- A compromised operating system, or a local attacker already holding the user's own privileges
- Unsigned release artifacts triggering an OS warning; this is documented, not a defect
- Missing rate limiting on the local user interface

## Security design

Summarised here; the full analysis with residual risk is in
[docs/THREAT-MODEL.md](docs/THREAT-MODEL.md).

| Control | Implementation |
|---|---|
| No client secret | Public client, Authorization Code Flow with PKCE |
| No credential handling | Authentication and consent in the system browser only; no embedded web view |
| Token storage | OS-native secure store, with a visible memory-only fallback and never plaintext |
| Log safety | Redaction wraps every logging provider, so no call site can bypass it |
| Tenant isolation | Tenant-specific authorities; a token from another tenant is rejected |
| Consent integrity | Cryptographically random state, validated on return; result verified by token acquisition |
| Least privilege | Tiered, opt-in scopes; broad scopes flagged with the narrower alternative named |
| Destructive actions | Previewed, confirmed by typing the name, and audited locally |
| Injection defences | CSV formula neutralization, Markdown escaping, path containment |
| Supply chain | Central version pinning, NuGet audit, Dependabot, CodeQL, SBOM, SHA-pinned actions |
| Publication safety | A gate scanning the working tree and full git history, failing closed |

## If a credential is exposed

Removing a secret from git does **not** un-disclose it. Anyone who cloned or forked the
repository, and any cache or mirror, may still hold it.

1. **Revoke and rotate it first.** Treat it as compromised from the moment it was pushed.
2. Remove it from history with `git filter-repo` — a later deletion commit is not enough.
3. Force-update the remote, and tell anyone with a clone or fork.
4. Re-run `./scripts/scan-secrets.sh` until it passes.
5. Review tenant sign-in and audit logs for use of the exposed credential.
