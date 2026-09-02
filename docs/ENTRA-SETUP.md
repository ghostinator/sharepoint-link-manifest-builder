# Microsoft Entra setup

This document is written for a Microsoft 365 / Entra administrator in **any** tenant. It
contains no tenant-specific values. Everything in angle brackets is supplied by your
organization.

---

## 1. Which path should you use?

| Path | Use when | Requires |
|---|---|---|
| **A — Automatic tenant setup** | You want the app to create the registration for you | A publisher-supplied bootstrap client ID **and** an authorized signed-in user |
| **B — Existing app registration** | Your organization creates app registrations centrally, or forbids self-service | A client ID your admins created (§4) |

Path B is always available and fully functional. Path A is a convenience, not a requirement.

> **Automatic setup is disabled unless a bootstrap client ID is configured.** This repository
> ships **no** client ID. See §6.

---

## 2. What gets created (Path A)

Exactly one object, plus a service principal:

- An **application registration** in your tenant:
  - `signInAudience`: `AzureADMyOrg` (single tenant — your tenant only)
  - `isFallbackPublicClient`: `true` (public client / desktop)
  - `publicClient.redirectUris`: `http://localhost` (loopback, no fixed port)
  - `requiredResourceAccess`: the delegated Microsoft Graph scopes in §3
  - **No `passwordCredentials`. No client secret. Ever.**
- A **service principal** (enterprise application), which Microsoft provisions automatically
  when consent is granted.

The wizard shows the exact object it is about to create, waits for you to confirm, and writes
a local sanitized audit entry afterwards. Nothing is created silently.

---

## 3. Permissions requested

### Normal operation (delegated — the app acts as the signed-in user)

| Scope | Type | Admin consent typically required | Why |
|---|---|---|---|
| `openid`, `profile`, `offline_access` | Delegated | No | Sign-in and silent refresh |
| `User.Read` | Delegated | No | Show who is signed in, and to which tenant |
| `Sites.Read.All` | Delegated | **Yes** | Find sites, read site metadata, list document libraries |
| `Files.ReadWrite.All` | Delegated | **Yes** | Enumerate files, create sharing links, write manifests |
| `User.ReadBasic.All` | Delegated | **Yes** | *Optional* — the User OneDrive people picker |
| `Sites.ReadWrite.All` | Delegated | **Yes** | *Optional and broad* — only if a library rejects the standard write |

Read-only profile, sufficient for browsing and dry runs:
`User.Read`, `Sites.Read.All`, `Files.Read.All`.

### Bootstrap (setup only — never used for file work)

| Tier | Scopes | When |
|---|---|---|
| Create-only (**default**) | `User.Read`, `AppRegistration.Create` | Automatic setup |
| Manage (**opt-in**) | `Application.ReadWrite.All` | Repair, Replace, service-principal creation, deep verification, deletion |

`AppRegistration.Create` is the least-privileged delegated permission for `POST /applications`.
Because the wizard submits a *complete* application object in one request, it never needs
`PATCH /applications/{id}` — which would require `Application.ReadWrite.All`. This is why the
default bootstrap tier is create-only.

`Directory.Read.All` and `Directory.ReadWrite.All` are higher-privileged alternatives that this
application **deliberately does not request**.

### What administrator consent does *not* do

Granting these permissions does **not**:

- give the app access to content the signed-in user cannot already access;
- override SharePoint or OneDrive item permissions;
- override tenant external-sharing policy, sensitivity labels, or Conditional Access;
- allow unattended, user-less access — this app has no application-only mode.

Delegated access is always the intersection of the granted scope **and** the signed-in user's
own effective permissions.

---

## 4. Path B — create the registration manually

In the Microsoft Entra admin center:

1. **App registrations -> New registration**
2. **Name:** anything your organization prefers.
3. **Supported account types:** *Accounts in this organizational directory only* (single tenant).
4. **Redirect URI:** select **Public client/native (mobile & desktop)** and enter
   `http://localhost`.
5. **Register.**
6. **Authentication ->** confirm *Allow public client flows* is **Yes**.
7. **API permissions -> Add a permission -> Microsoft Graph -> Delegated permissions**, add:
   - `User.Read`
   - `Sites.Read.All`
   - `Files.ReadWrite.All`
   - *(optional)* `User.ReadBasic.All`
8. **Grant admin consent for \<your organization\>** — or leave it, and let a user consent if
   your tenant policy permits user consent.
9. **Do not create a client secret or certificate.** The application does not use one and will
   not ask for one.
10. Copy the **Application (client) ID** and the **Directory (tenant) ID** into the app's
    setup wizard (*Existing app registration*).

Equivalent CLI, for administrators who prefer it — note this is **optional**; the application
itself never requires a terminal:

```bash
az ad app create \
  --display-name "<your app name>" \
  --sign-in-audience AzureADMyOrg \
  --is-fallback-public-client true \
  --public-client-redirect-uris "http://localhost"
```

---

## 5. Administrator consent

Consent always happens in Microsoft's own interface, in your system browser. The application
builds the official URL and opens it:

```
https://login.microsoftonline.com/<tenant-id>/v2.0/adminconsent
  ?client_id=<client-id>
  &scope=<space-separated scopes>
  &redirect_uri=http://localhost:<port>
  &state=<random>
```

The application never renders a screen that looks like Microsoft's consent page and never
collects credentials.

### If you cannot grant tenant-wide consent

The wizard does not dead-end. It offers:

- **Copy Consent Link** — send it to an authorized administrator.
- **Open Consent Page** — for an admin sitting at the same machine.
- **Save as Pending Consent** — the configuration is stored and the app shows
  *Waiting for Administrator Approval*.
- **Check Again** — an explicit, manual re-check. The app does not poll aggressively.

### Verification

Consent is **verified**, not assumed. The app acquires a real token for the tenant-specific
client ID and compares the scopes Entra actually issued against the scopes required. A redirect
that merely *looks* successful is not accepted as proof.

---

## 6. For publishers: configuring the bootstrap application

Automatic setup (Path A) needs a publisher-owned bootstrap identity. To enable it:

1. In the **publisher's own tenant**, register a **multitenant** application:
   - `signInAudience`: `AzureADMultipleOrgs`
   - Public client / native, redirect URI `http://localhost`
   - Delegated permissions: `User.Read`, `AppRegistration.Create`
   - Optionally `Application.ReadWrite.All` for the opt-in Manage tier
   - **No client secret.**
2. Supply the client ID to the application by any one of:
   - environment variable `SPLMB_BOOTSTRAP_CLIENT_ID`
   - `appsettings.json` -> `Bootstrap:ClientId`
   - the setup wizard's *Advanced* field
   - a build-time property: `dotnet build -p:BootstrapClientId=<guid>`
3. Publish a privacy policy and support URL and set them on the registration, so the consent
   screen shows a real publisher.

**Do not commit a client ID to source control.** This repository intentionally contains none;
`Bootstrap:ClientId` is empty in the sample configuration and automatic setup reports itself
as unavailable until it is set.

---

## 7. `Sites.Selected` (Advanced)

`Sites.Selected` restricts an application to specific SharePoint sites, assigned individually.
It is genuinely different from normal delegated access:

- Consent alone grants **nothing**. Each site must also be explicitly assigned to the app's
  service principal with a read or write role.
- Site assignment is an administrative operation performed against each site.
- It does **not** behave like `Sites.Read.All` scoped down; unassigned sites are simply invisible.

This application surfaces `Sites.Selected` as **Advanced**, documents the two-step model, and
does not imply it is interchangeable with normal delegated access. Graphical site-assignment
management is offered only where the chosen authorization model supports it securely.

---

## 8. Removing the application

| Action | Effect | Default |
|---|---|---|
| **Remove Local Configuration** | Deletes local settings, token cache and tenant config on this machine only | **Yes — this is the default** |
| **Delete App Registration** | Deletes the Entra object for the whole tenant | No — explicit, guarded |

Deleting the registration requires: an automatically-created registration (manually supplied
ones are not deletable by default), sufficient authorization, typing the application's display
name to confirm, and acknowledging that other installations may depend on it. **Uninstalling
the application never deletes the registration.**
