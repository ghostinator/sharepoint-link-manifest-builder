# Microsoft Graph operations

Every Microsoft Graph call this application can make is listed here. All calls target the
**v1.0** endpoint (`https://graph.microsoft.com/v1.0`). No beta endpoint is used; see
§9.

Permission column shows the **delegated (work or school)** permission this application
requests. "Modifies?" means the call changes tenant configuration or content.

---

## 1. Identity and profile

| # | Purpose | Method + endpoint | Permission | Modifies? |
|---|---|---|---|---|
| 1.1 | Signed-in user and tenant for the header/Home page | `GET /me?$select=id,displayName,userPrincipalName` | `User.Read` | No |
| 1.2 | Tenant display name for manifests and the Permissions page | `GET /organization?$select=id,displayName,verifiedDomains` | `User.Read` | No |

**Pagination:** none. **Errors:** `401` -> reauthenticate; `403` -> scope missing.
`GET /organization` may be denied in restricted tenants; the app degrades to the tenant GUID
and says so rather than inventing a name.

---

## 2. SharePoint site discovery

| # | Purpose | Method + endpoint | Permission | Modifies? |
|---|---|---|---|---|
| 2.1 | Search sites by keyword | `GET /sites?search={q}&$top=50` | `Sites.Read.All` | No |
| 2.2 | Resolve a pasted site URL | `GET /sites/{hostname}:/{server-relative-path}` | `Sites.Read.All` | No |
| 2.3 | Root site of the tenant | `GET /sites/root` | `Sites.Read.All` | No |
| 2.4 | Site metadata by ID | `GET /sites/{siteId}` | `Sites.Read.All` | No |
| 2.5 | Sites the user follows (Recent) | `GET /me/followedSites` | `Sites.Read.All` | No |
| 2.6 | Document libraries of a site | `GET /sites/{siteId}/drives?$select=id,name,driveType,webUrl` | `Sites.Read.All` | No |
| 2.7 | Default library of a site | `GET /sites/{siteId}/drive` | `Sites.Read.All` | No |

**Pagination:** 2.1, 2.5, 2.6 follow `@odata.nextLink` until absent.

**Honesty constraint:** `GET /sites?search=` does **not** return every site in the tenant. It
returns what the search index exposes to the signed-in user. The UI says so and always offers
"Paste a SharePoint URL" as the exact-resolution path. The application never claims the list
is exhaustive.

---

## 3. OneDrive and drive resolution

| # | Purpose | Method + endpoint | Permission | Modifies? |
|---|---|---|---|---|
| 3.1 | The signed-in user's OneDrive | `GET /me/drive?$select=id,name,driveType,webUrl,owner` | `Files.Read.All` | No |
| 3.2 | Another user's OneDrive | `GET /users/{userId}/drive` | `Files.Read.All` | No |
| 3.3 | People picker for User OneDrive | `GET /users?$search="displayName:{q}" OR $filter=startswith(...)` | `User.ReadBasic.All` | No |
| 3.4 | Drive by ID | `GET /drives/{driveId}` | `Files.Read.All` | No |
| 3.5 | Resolve a pasted sharing URL | `GET /shares/{sharingToken}/driveItem` | `Files.Read.All` | No |

**Sharing token encoding (3.5):** `u!` + base64url(URL), with `+`->`-`, `/`->`_`, `=` trimmed.
Implemented and unit-tested in `Core/Urls/GraphShareTokenEncoder`.

**3.2 caveat, surfaced verbatim in the UI:** administrator consent does not by itself grant
access to every user's OneDrive. A user's drive may be unprovisioned, access-denied, or
blocked by policy. The application reports the normalized Graph error and **never
auto-provisions a OneDrive**.

---

## 4. Item and folder enumeration

| # | Purpose | Method + endpoint | Permission | Modifies? |
|---|---|---|---|---|
| 4.1 | Root folder of a drive | `GET /drives/{driveId}/root` | `Files.Read.All` | No |
| 4.2 | Children of a folder (the enumeration workhorse) | `GET /drives/{driveId}/items/{itemId}/children?$top=200&$select=...` | `Files.Read.All` | No |
| 4.3 | Resolve a folder by path | `GET /drives/{driveId}/root:/{path}` | `Files.Read.All` | No |
| 4.4 | Single item by ID | `GET /drives/{driveId}/items/{itemId}` | `Files.Read.All` | No |

`$select` for 4.2 is
`id,name,size,webUrl,eTag,cTag,lastModifiedDateTime,createdDateTime,folder,file,package,remoteItem,shared,parentReference`
— enough to classify, filter and identify without over-fetching.

**Pagination:** always follow `@odata.nextLink`. `$top` is a hint; Graph may return fewer.
**Recursion:** performed client-side by revisiting 4.2 for each child folder, never by fetching
a whole library up front. Non-recursive targets call 4.2 exactly once per target folder.
**Errors:** `404` -> folder deleted mid-run; `403` -> access denied on a subtree (recorded and
skipped, the job continues); `429`/`503` -> retry policy.

---

## 5. Sharing links

| # | Purpose | Method + endpoint | Permission | Modifies? |
|---|---|---|---|---|
| 5.1 | Create or obtain a sharing link | `POST /drives/{driveId}/items/{itemId}/createLink` | `Files.ReadWrite.All` | **Yes — content permissions** |
| 5.2 | Grant named recipients | `POST /drives/{driveId}/items/{itemId}/invite` | `Files.ReadWrite.All` | **Yes — content permissions** |
| 5.3 | Inspect existing permissions | `GET /drives/{driveId}/items/{itemId}/permissions` | `Files.Read.All` | No |

### 5.1 `createLink`

```jsonc
{
  "type": "view",            // view | edit | embed (embed = OneDrive personal only)
  "scope": "organization",   // anonymous | organization | users
  "expirationDateTime": "2026-12-31T23:59:59Z",  // optional
  "retainInheritedPermissions": true             // optional, default true
}
```

**Response status is product-significant:**

| Status | Meaning | Recorded as |
|---|---|---|
| `201 Created` | A new sharing link was created | `Created` |
| `200 OK` | An equivalent link already existed and was returned | `Reused` |

The application never reports "Created" for a `200`. `password` is OneDrive-personal only and
is not exposed by this product.

### 5.2 `invite`

```jsonc
{
  "recipients":    [ { "email": "..." } ],
  "roles":         [ "read" ],   // or [ "write" ]
  "requireSignIn": true,
  "sendInvitation": false,       // default false: no email unless explicitly chosen
  "expirationDateTime": "...",   // optional
  "retainInheritedPermissions": true
}
```

**`207 Multi-Status`** is returned when some recipients succeed and others fail; each failed
entry carries its own `error`. The application parses per-recipient outcomes and records
recipient rejection individually rather than failing the whole file.

The v1.0 `createLink` action has **no** `recipients` parameter. "Specific people" is therefore
`createLink` with `scope: users` for the URL, plus an optional `invite` to grant the named
people. An `invite` response permission may contain **no** `link.webUrl` — it can grant direct
access without producing a URL. The application reports exactly what it received.

---

## 6. Manifest storage

| # | Purpose | Method + endpoint | Permission | Modifies? |
|---|---|---|---|---|
| 6.1 | Read an existing manifest | `GET /drives/{driveId}/root:/{path}:/content` | `Files.Read.All` | No |
| 6.2 | Manifest metadata + ETag | `GET /drives/{driveId}/root:/{path}` | `Files.Read.All` | No |
| 6.3 | Upload a small manifest (< 4 MiB) | `PUT /drives/{driveId}/items/{parentId}:/{name}:/content` | `Files.ReadWrite.All` | **Yes — content** |
| 6.4 | Create an upload session (large) | `POST /drives/{driveId}/items/{parentId}:/{name}:/createUploadSession` | `Files.ReadWrite.All` | **Yes — content** |
| 6.5 | Upload a chunk | `PUT {uploadUrl}` with `Content-Range` | (session-authorized) | **Yes — content** |

**Concurrency:** 6.3 sends `If-Match: {eTag}` when updating a known manifest. `412 Precondition
Failed` is surfaced as a typed `ManifestConflict`, never retried blindly.
**Conflict behaviour:** governed by `ManifestConflictPolicy` (ADR-0010).
**Chunking:** upload sessions use 3,276,800-byte chunks (a multiple of 320 KiB, as Graph
requires). Chunk `PUT`s carry no `Authorization` header — the session URL is pre-authorized.

---

## 7. Entra onboarding (setup only)

| # | Purpose | Method + endpoint | Permission | Modifies? |
|---|---|---|---|---|
| 7.1 | Create the registration | `POST /applications` | `AppRegistration.Create` | **Yes — tenant** |
| 7.2 | Read a registration | `GET /applications(appId='{appId}')` | `Application.Read.All` | No |
| 7.3 | Repair a registration | `PATCH /applications/{objectId}` | `Application.ReadWrite.All` | **Yes — tenant** |
| 7.4 | Locate the service principal | `GET /servicePrincipals(appId='{appId}')` | `Application.Read.All` | No |
| 7.5 | Create a service principal (repair only) | `POST /servicePrincipals` | `Application.ReadWrite.All` | **Yes — tenant** |
| 7.6 | Inspect delegated grants (optional) | `GET /oauth2PermissionGrants?$filter=clientId eq '{spId}'` | `Directory.Read.All` | No |
| 7.7 | Delete a registration (explicit, guarded) | `DELETE /applications/{objectId}` | `Application.ReadWrite.All` | **Yes — destructive** |

The `POST` body in 7.1 carries `signInAudience`, which is `AzureADMyOrg` by default and
`AzureADMultipleOrgs` when the user has chosen multi-organization mode. The body is always
complete, so the create path never needs the `PATCH` in 7.3 and therefore never needs
`Application.ReadWrite.All`. This is also why the audience cannot be changed later by this
application — see [ADR-0011](adr/0011-multi-tenant-authority.md).

### 7.1 request body — complete, so no PATCH is needed

```jsonc
{
  "displayName": "<user-confirmed name>",
  "signInAudience": "AzureADMyOrg",
  "isFallbackPublicClient": true,
  "publicClient": { "redirectUris": [ "http://localhost" ] },
  "requiredResourceAccess": [
    {
      "resourceAppId": "00000003-0000-0000-c000-000000000000",   // Microsoft Graph
      "resourceAccess": [ { "id": "<scope GUID>", "type": "Scope" } ]
    }
  ]
}
```

Sending everything at once is what allows the default bootstrap tier to be create-only
(ADR-0005). Operations 7.3, 7.5 and 7.7 are **opt-in elevated** and are never performed
on the happy path.

**Every operation in §7 is previewed in the UI before execution and written to a local
sanitized audit entry afterwards.** Nothing here happens silently.

---

## 8. Consent (browser, not an API call)

Consent is not a Graph call. The application builds an official Microsoft URL and opens the
**system browser**:

```
https://login.microsoftonline.com/{tenantId}/v2.0/adminconsent
  ?client_id={clientId}
  &scope={space-separated scopes}
  &redirect_uri={registered loopback}
  &state={cryptographically random, validated on return}
```

`{tenantId}` is always one explicit organization: the configured tenant in single-organization
mode, or the signed-in organization in multi-organization mode. It is **never** `organizations`
or `common`. When the organization cannot be determined, the request is refused rather than
broadened — an administrator signed into several directories could otherwise consent in the
wrong one.

The application never renders a consent-like screen, never collects credentials, and validates
both `state` and the returned `tenant` before accepting the result. Success is then
**verified** by token acquisition (ADR-0006), not inferred from the redirect.

---

## 9. Beta endpoints

**None are used.** Every operation above exists in v1.0. If a future feature requires beta, the
brief's conditions apply: isolate behind an interface, label Experimental, document the reason,
and keep the core product functional without it.

---

## 10. Cross-cutting behaviour

**Pagination.** A single helper follows `@odata.nextLink` until absent, yielding pages lazily
and honouring the cancellation token between pages.

**Throttling and retry.** Retried: `429`, `503`, `504`, `502`, `408`, and transient socket
errors. Honoured first: the `Retry-After` header. Otherwise exponential backoff
`base * 2^attempt` with full jitter, capped, to a configurable attempt limit. Never retried:
`400`, `401`, `403`, `404`, `409`, `412`.

**Correlation.** Every request carries a `client-request-id` GUID; the response
`request-id`/`client-request-id` are captured into errors and diagnostics for support, and
appear in the sanitized job report.

**Sanitized logging.** Request URLs are logged with query strings stripped of sensitive
parameters. Authorization headers, tokens, and authorization codes are never logged.
