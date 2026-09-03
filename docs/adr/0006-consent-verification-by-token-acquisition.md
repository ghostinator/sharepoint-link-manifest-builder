# ADR-0006: Verify consent by acquiring a token, not by reading directory objects

- **Status:** Accepted
- **Date:** 2026-09-02

## Context
The brief requires consent to be *verified* rather than assumed. The obvious approach is to
read `/oauth2PermissionGrants` and compare — but that needs a directory read permission
(`Directory.Read.All` or an equivalent grant-read scope) that the product otherwise has no
reason to hold.

## Decision
Primary verification is **an actual token acquisition against the tenant-specific client ID
for the required scopes**, inspecting `AuthenticationResult.Scopes` — the scopes Microsoft
Entra actually issued.

- Success with all required scopes present -> consent verified.
- `MsalUiRequiredException` with a consent/interaction-required condition -> consent pending.
- Missing scopes in an otherwise successful result -> partial consent, listed precisely.

Directory-object inspection is an *optional enhancement*, used only when an elevated read
permission happens to already be granted.

## Consequences
- Verification requires **no additional permission**, which keeps the least-privilege posture.
- It tests the thing that actually matters — whether the app can obtain a usable token — not a
  directory object that might not reflect effective access.
- `AuthenticationResult.Scopes` is read directly from MSAL; **the raw access token is never
  parsed, logged, or inspected**.
- Verification cannot distinguish "admin consented tenant-wide" from "this user consented"
  by token alone; consent *type* is reported as such, and refined via directory read when
  that is available.

## Alternatives considered
- **Read `/oauth2PermissionGrants` always** — requires a permission the product does not need,
  and models intent rather than effective access.
- **Assume the consent redirect implies success** — explicitly forbidden by the brief, and
  wrong: a redirect can carry an error, a wrong tenant, or a partial grant.
