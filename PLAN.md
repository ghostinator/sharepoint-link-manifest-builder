# PLAN — SharePoint Link Manifest Builder

Status: living document. Updated as phases complete.
Schema of this plan: goal -> constraints -> architecture -> phased delivery -> acceptance.

---

## 1. Goal

A cross-platform desktop application that lets an authorized Microsoft 365 user graphically
select SharePoint sites, document libraries, folders, their own OneDrive, or an accessible
user's OneDrive; enumerate the files within; create Microsoft 365 sharing links using a
requested permission and audience *when tenant policy allows*; and write link manifests back
into SharePoint or OneDrive.

The manifests exist to give Microsoft Copilot (or another authorized AI system) an explicit,
reliable list of file links rather than relying on implicit discovery.

## 2. Non-negotiable constraints

| # | Constraint | How it is enforced |
|---|---|---|
| C1 | No tenant-specific value is hard-coded | All tenant/client/site/drive values come from runtime config or user selection. CI runs a publication safety scan. |
| C2 | No client secret in the desktop app | Public client + Authorization Code Flow with PKCE via MSAL. No `passwordCredentials` are ever created. |
| C3 | Admin consent is always a human action in Microsoft's own UI | The app only *builds* the official consent URL and opens the system browser. It never renders a consent-like screen. |
| C4 | Consent is verified, never assumed | Verification acquires a real token and compares granted scopes. |
| C5 | Sharing links never override tenant policy | The app requests; Microsoft 365 decides. Result states distinguish Created / Reused / PolicyBlocked / AccessDenied. |
| C6 | Tokens are never logged, exported, or written to settings | Secure OS storage via MSAL extensions; memory-only fallback; a redaction layer on all logging. |
| C7 | Least privilege | Scope sets are tiered; broad scopes are opt-in and flagged in the UI. |
| C8 | The standard test suite needs no live tenant | All Graph tests run against a mocked `HttpMessageHandler`. |

## 3. Architecture at a glance

```
SharePointLinkManifestBuilder.App     Avalonia 12 + MVVM. Views, ViewModels, DI composition root.
        |  (depends on)
SharePointLinkManifestBuilder.Graph   MSAL auth, GraphApiClient (HTTP), Graph-facing services,
        |                             Entra registration + consent services.
SharePointLinkManifestBuilder.Core    Domain models, URL parsing, filtering, overlap planning,
                                      manifest format/parse/merge, retry policy, redaction,
                                      job orchestration. No HTTP. No UI. No MSAL.
```

Dependency direction is one-way and enforced by project references: `App -> Graph -> Core`.
View models never touch `HttpClient` or MSAL types; they consume Core abstractions.

### Why raw Graph HTTP instead of the Graph SDK
See `docs/adr/0003-raw-graph-http-over-sdk.md`. Summary: deterministic mocked testing,
exact control over `Retry-After` / ETag / pagination semantics, and a smaller self-contained
publish. MSAL is still used for identity — authentication is never hand-rolled.

## 4. Permission model (researched against Microsoft Learn, September 2026)

### 4.1 Normal operation — delegated, permanent tenant-specific app

| Scope | Why it is required | Notes |
|---|---|---|
| `openid`, `profile`, `offline_access` | Sign-in and silent token refresh | MSAL standard |
| `User.Read` | Show the signed-in user and tenant | Least privilege |
| `Sites.Read.All` | Site discovery/search, site metadata, list document libraries | No lower-privileged alternative for cross-site discovery |
| `Files.ReadWrite.All` | Enumerate drive items, `createLink` / `invite`, upload manifests | Required because targets span SharePoint libraries and other users' OneDrive |
| `User.ReadBasic.All` | *Optional.* The User OneDrive people picker | Only requested if the user enables the User OneDrive source |
| `Sites.ReadWrite.All` | *Optional, broad.* Only where a library rejects `Files.ReadWrite.All` writes | Flagged as broad in the UI |

Read-only profile for dry runs: `User.Read`, `Sites.Read.All`, `Files.Read.All`.

Delegated access is always bounded by the signed-in user's effective SharePoint/OneDrive
permissions. Admin consent does **not** widen what the user can see.

### 4.2 Bootstrap — setup only, never used for file work

**Tier 1 (default, create-only).** `openid profile offline_access User.Read AppRegistration.Create`

`POST /applications` accepts `AppRegistration.Create` as its least-privileged delegated
permission. By sending a *complete* application object in the initial POST — `displayName`,
`signInAudience`, `isFallbackPublicClient`, `publicClient.redirectUris`, `requiredResourceAccess`
— the wizard never needs `PATCH /applications/{id}`, which would require
`Application.ReadWrite.All`. The service principal is provisioned by Microsoft as a side effect
of the official consent flow, so `POST /servicePrincipals` (also `Application.ReadWrite.All`)
is not on the happy path.

**Tier 2 (opt-in, elevated).** `Application.ReadWrite.All`
Only for Repair Registration, Replace Registration, explicit service-principal creation, deep
directory verification, and optional registration deletion. Requested only when the operator
chooses one of those actions. `Directory.ReadWrite.All` is a documented higher-privileged
alternative that this application deliberately never requests.

Directory-role reality: creating an app registration additionally depends on the signed-in
user's Entra role (default user permissions unless restricted; Application Developer; Cloud
Application Administrator; Application Administrator). The wizard detects and explains failure
rather than pre-asserting capability.

## 5. Sharing-link semantics (researched, not assumed)

- `POST /drives/{driveId}/items/{itemId}/createLink` — body `type` (`view`|`edit`|`embed`),
  `scope` (`anonymous`|`organization`|`users`), `expirationDateTime`, `retainInheritedPermissions`.
  **`201 Created` = new link. `200 OK` = an equivalent link already existed and was returned.**
  The application records those as distinct outcomes rather than claiming creation.
- v1.0 `createLink` has **no** `recipients` parameter. Recipient targeting requires
  `POST /drives/{driveId}/items/{itemId}/invite` with `recipients`, `roles`, `requireSignIn`,
  `sendInvitation`. `invite` can return `207 Multi-Status` for per-recipient failure, and the
  returned permission may carry no `link.webUrl` at all.
- Therefore "Specific people" is implemented as: `createLink` with `scope: users` to obtain the
  URL, plus an optional `invite` (default `sendInvitation: false`) to grant the named recipients.
  Both outcomes are recorded separately and honestly.
- `password` on `createLink`/`invite` is OneDrive-personal only and is not exposed.
- `embed` is OneDrive-personal only and is offered only where applicable.

## 6. Phased delivery

| Phase | Content | Commit prefix |
|---|---|---|
| 0 | Repository, solution, analyzers, .gitignore | `chore:` |
| 1 | PLAN, ARCHITECTURE, ENTRA-SETUP, GRAPH-OPERATIONS, THREAT-MODEL, ADRs | `docs:` |
| 2 | Core models + abstractions | `feat:` |
| 3 | Core logic: URL parsing, filters, overlap planning, retry policy, redaction | `feat:` |
| 4 | Manifest engine: format, parse, merge, conflict policy, CSV/JSON/Markdown safety | `feat:` |
| 5 | Graph transport: `GraphApiClient`, pagination, throttling, error normalization | `feat:` |
| 6 | Identity: MSAL, secure token storage, bootstrap, registration, consent, verification | `feat:` |
| 7 | Graph services: sites, drives, users, sharing, manifest storage | `feat:` |
| 8 | Job engine: preflight, discovery, execution, retry, cancellation, summary | `feat:` |
| 9 | Avalonia shell, navigation, resource tree, selectors | `feat:` |
| 10 | Setup wizard (8 pages), Permissions page, Diagnostics, Settings, History | `feat:` |
| 11 | Unit + mocked-Graph integration tests, headless UI launch test | `test:` |
| 12 | Scripts, packaging, GitHub Actions, publication safety gate | `build:` `ci:` `security:` |
| 13 | Remaining documentation, rollback, release process | `docs:` |

## 7. Acceptance

The 50 acceptance criteria in the product brief are tracked in
`docs/ACCEPTANCE.md`, each mapped to the code or document that satisfies it, with
live-tenant-only items explicitly marked as unverified in this environment.

## 8. Known external dependencies this repository cannot satisfy

These are configuration points, implemented and documented but intentionally unset:

1. **Bootstrap client ID** — automatic setup stays disabled until a publisher supplies one.
   No client ID is fabricated or committed. The existing-registration path is fully functional
   without it.
2. **Publisher identity** — `PLACEHOLDER-PUBLISHER` throughout.
3. **Privacy policy / terms / support URLs** — placeholders.
4. **Code-signing and Apple notarization credentials** — packaging emits clearly labelled
   unsigned artifacts.
5. **A live Microsoft 365 tenant** — no live Graph call has been executed by this build.
