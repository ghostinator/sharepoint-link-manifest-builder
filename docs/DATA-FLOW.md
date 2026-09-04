# Data flow

Where data travels, and where it does not.

## Trust boundaries

```
  +---------------------------------------------------------------+
  |  YOUR MACHINE                                                  |
  |                                                                |
  |   +-------------------+        +---------------------------+   |
  |   |  Application UI   |        |  OS secure store          |   |
  |   |                   |<------>|  Keychain / DPAPI /       |   |
  |   |  View models      |        |  Secret Service           |   |
  |   +---------+---------+        +---------------------------+   |
  |             |                                                  |
  |             |                  +---------------------------+   |
  |             +----------------->|  Local files              |   |
  |             |                  |  settings, tenant config, |   |
  |             |                  |  profiles, history, audit,|   |
  |             |                  |  logs (redacted)          |   |
  |             |                  +---------------------------+   |
  |             |                                                  |
  |   +---------v---------+        +---------------------------+   |
  |   |  Graph transport  |        |  System browser           |   |
  |   +---------+---------+        +-------------+-------------+   |
  |             |                                |                 |
  +-------------|--------------------------------|-----------------+
                |  TLS                           |  TLS
                v                                v
     +---------------------+          +------------------------+
     | graph.microsoft.com |          | login.microsoftonline  |
     |  (Microsoft Graph)  |          |  .com (Microsoft Entra)|
     +---------------------+          +------------------------+

  No other network destination exists.
```

## Sign-in

1. The application builds a tenant-specific authority URL. **Never `/common`**, so a token from
   another directory cannot be accepted.
2. MSAL opens the **system browser**. The application does not render the sign-in page and never
   sees a credential.
3. You authenticate with Microsoft, completing any MFA or Conditional Access requirement.
4. Microsoft redirects to a loopback listener bound to `127.0.0.1` on an ephemeral port, which
   serves exactly one request.
5. MSAL exchanges the authorization code using PKCE.
6. The token is written to OS-native secure storage, or held in memory when that is unavailable.
7. The tenant in the returned token is compared with the configured tenant; a mismatch is
   rejected.

**The application never handles your password.** It receives a token, and only ever passes that
token to the Graph transport as an `Authorization` header.

## Consent

1. The application builds an official Microsoft admin-consent URL with a cryptographically
   random `state`.
2. The system browser opens it. The application renders nothing that resembles a consent screen.
3. Microsoft redirects to the loopback listener.
4. `state` is validated, and the returned tenant is compared with the expected tenant.
5. **The redirect is not treated as proof.** The application then acquires a real token and
   compares the scopes Microsoft Entra actually issued against those required.

## Browsing

```
You expand a folder
   -> View model asks IDriveService
      -> Graph transport attaches a bearer token and a correlation ID
         -> GET https://graph.microsoft.com/v1.0/drives/{id}/items/{id}/children
            <- JSON, paged
      <- mapped to display models
   <- tree shows friendly names
```

Only the fields the application actually uses are deserialized, so no unused tenant data is held
in memory. Nothing is written to disk during browsing.

## Running a job

```
Preview   enumerate -> filter -> deduplicate -> counts and warnings   (no writes)
Execute   per file: POST .../createLink   [+ POST .../invite]
Manifests read existing -> merge -> conditional PUT with If-Match
Record    counts, manifest locations and sanitized errors to local history
```

A dry run stops after Preview and performs no write of any kind.

## What leaves your machine

| Data | Sent to | Why |
|---|---|---|
| Bearer token | Microsoft Graph | Authenticating each request |
| Site, drive, folder and item identifiers | Microsoft Graph | Reading and modifying the items you selected |
| Requested link type, scope and expiry | Microsoft Graph | Creating the sharing link you asked for |
| Recipient addresses | Microsoft Graph | Only when you choose Specific people and supply them |
| Manifest content | Microsoft Graph | Writing the manifest into your own tenant |
| Correlation identifier | Microsoft Graph | Support traceability |

## What never leaves your machine

- Your password — the application never receives it
- Refresh tokens — held by the OS secure store and used only by MSAL
- Job history, saved profiles, the local audit trail and log files
- Diagnostic bundles, unless you send one yourself
- Any usage or analytics data — none is collected

## Manifests

A manifest is written **into your own tenant**, to a location you chose. It contains file names,
paths, web URLs, sharing links, drive and item identifiers, and the tenant name and identifier.

It deliberately never contains a token, an authorization header, a secret, a local file path, or
personal information beyond what the source type requires (a User OneDrive manifest names the
owning user, because otherwise it would be ambiguous).

Because a manifest contains sharing links, it is as sensitive as the content it indexes. Store it
where you would store that content.
