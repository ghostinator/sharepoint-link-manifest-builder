# ADR-0008: OS-native secure token storage with an explicit memory-only fallback

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
Refresh tokens are long-lived bearer credentials. Writing them to a plain file would let any
local process or backup sweep them up.

## Decision
Use `Microsoft.Identity.Client.Extensions.Msal` to place the MSAL cache in OS-native storage:

| Platform | Mechanism |
|---|---|
| Windows | DPAPI-protected cache file (per-user) |
| macOS | Keychain |
| Linux | Secret Service / libsecret keyring |

At startup the store is **probed with a real read/write round-trip**. If it is unavailable —
a headless Linux session with no keyring being the common case — the application:

1. tells the user plainly, in the UI, that secure storage is unavailable;
2. falls back to a **memory-only** cache, so sign-in is required each launch;
3. **never** silently writes tokens in plaintext.

Tokens are never written to the settings file, job history, diagnostic bundles, or logs.

## Consequences
- Credentials are protected by the OS on all three platforms.
- Linux without a keyring degrades to a worse but safe experience, visibly.
- `Forget Account`, `Clear Token Cache`, `Clear Cached Data` and `Remove Tenant Configuration`
  are exposed as explicit user actions.

## Alternatives considered
- **Encrypted file with an app-derived key** — the key must live in the binary, so it is
  obfuscation, not protection.
- **Plaintext cache file** — unacceptable.
