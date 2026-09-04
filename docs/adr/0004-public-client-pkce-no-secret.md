# ADR-0004: Public client with Authorization Code Flow + PKCE, and no client secret

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
A desktop application distributed to end users cannot keep a secret. Anything shipped in the
binary is readable by anyone who has the binary.

## Decision
The application is a **public client**. Authentication is MSAL interactive Authorization Code
Flow with PKCE against the **system browser** using a loopback redirect
(`http://localhost` with a dynamic port). The registration is created with
`isFallbackPublicClient: true` and a native/loopback redirect URI. **No `passwordCredentials`
are ever created for the desktop application**, by automatic setup or otherwise.

## Consequences
- MFA and Conditional Access work, because authentication happens in the real browser where
  the device/session state and any broker live.
- No secret exists to leak, rotate, or scan for.
- The system browser is required; there is no embedded web view. This is deliberate — an
  embedded view lets the host application observe credentials, which is exactly the
  consent-phishing shape this product must not have.
- Loopback redirect requires binding a local ephemeral port at sign-in time.

## Alternatives considered
- **Confidential client with a secret** — impossible to protect in a distributed desktop app.
- **Device code flow** — usable but a worse experience and weaker Conditional Access story;
  retained only as a documented fallback idea, not implemented.
- **Embedded WebView** — rejected on security grounds, per the brief.
