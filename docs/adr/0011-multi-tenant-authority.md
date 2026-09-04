# ADR-0011: Optional multi-tenant authority, with the tenant resolved at sign-in

- **Status:** Accepted
- **Date:** 2026-09-03

## Context
The original design fixed the authority to a single tenant
(`https://login.microsoftonline.com/{tenantId}`) and refused any token whose tenant did not
match. That is the strongest possible defence against cross-tenant confusion, because a token
from another tenant is not merely rejected — it is unobtainable.

It also makes a legitimate use case impossible. A consultant or managed service provider works
across several customer tenants. Under the original design each one needs its own app
registration, its own client ID, and a full re-run of the setup wizard to move between them.
Re-running setup to switch customers is enough friction that the realistic outcome is people
avoiding the tool, or worse, keeping several installs.

A second problem surfaced in practice. A single-tenant authority rejects an account from
another tenant with `AADSTS50020`, which is returned on the loopback redirect. MSAL renders its
"authentication complete" page as soon as *any* redirect arrives, so the browser reports success
and the application reports a bare failure. The genuine cause was invisible.

## Decision
The audience is **configurable per installation**, defaulting to the narrower option:

- `TenantAudience.SingleTenant` (default) keeps the tenant-specific authority and the existing
  hard rejection of any token from another tenant. Unchanged behaviour.
- `TenantAudience.AnyOrganization` uses `https://login.microsoftonline.com/organizations` and
  resolves the tenant from the token at sign-in.

Deliberately `/organizations`, **never `/common`**. `/common` additionally admits personal
Microsoft accounts, which have no Entra tenant and no SharePoint or OneDrive for Business, so
they can only ever fail later and less clearly.

Three constraints keep the multi-tenant path honest:

1. **Consent is always requested in a named directory.** The admin-consent URL never uses
   `/organizations`; it names an explicit tenant GUID, resolved from the signed-in account.
   When no tenant can be determined, consent is refused rather than guessed. An explicit tenant
   in the consent URL is what stops an administrator signed into several directories from
   consenting in the wrong one.
2. **Consent does not travel.** Each tenant grants its own consent to the multi-tenant
   registration. Consent in tenant A confers nothing in tenant B.
3. **The active tenant is explicit and visible.** Switching organizations is a deliberate user
   action in the organization switcher, and the active organization is shown on the home page.
   All Graph calls use the token for the active account, and site and drive identifiers are
   tenant-scoped.

## Consequences
- One installation serves many customer tenants; switching is one click and silent whenever the
  cached refresh token is still valid.
- The registration must be created with `signInAudience: AzureADMultipleOrgs`. This cannot be
  changed after creation through the create-only bootstrap path, because changing it is a
  `PATCH` and `PATCH /applications/{id}` requires `Application.ReadWrite.All` — a far broader
  permission than the create-only tier this product asks for (see ADR-0009). Choosing the
  audience is therefore a setup-time decision.
- The compile-time guarantee against cross-tenant token acceptance is weakened to a runtime one
  for multi-tenant installations. This is the real cost of the decision. It is mitigated by the
  three constraints above, and it does not apply at all to the default single-tenant mode.
- Every tenant an operator connects to appears in the local MSAL cache. "Forget account" and
  "Clear token cache" remove them.

## Supersedes
The blanket statement in ADR-0004 and in the original threat model that the authority is
*always* tenant-specific. That remains true for the default configuration and is now an
explicit, per-installation choice rather than an invariant.
