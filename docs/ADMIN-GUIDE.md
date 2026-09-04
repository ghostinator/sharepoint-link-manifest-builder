# Administrator guide

For the Microsoft 365 or Microsoft Entra administrator deciding whether to allow this
application, and on what terms.

## What it is

A desktop application that creates Microsoft 365 sharing links in bulk and writes manifests of
those links back into SharePoint or OneDrive. Its purpose is to give an AI system such as
Microsoft Copilot an explicit list of files.

It is a **public client** using **delegated** permissions only. There is no application-only
mode, no unattended operation, and no client secret.

## What it can and cannot do

**It can**, as the signed-in user and only where that user already has access:

- Enumerate sites, libraries, folders and files
- Create sharing links
- Write manifest files

**It cannot**:

- Access anything the signed-in user cannot already open
- Override SharePoint or OneDrive item permissions
- Override your external sharing policy, sensitivity labels or Conditional Access
- Operate without a signed-in user
- Provision a OneDrive for anybody

Granting administrator consent does not change any of the above. Consent authorizes the
application to act **as a user**, bounded by that user's own effective permissions.

## Risk assessment

The genuine risk is **over-sharing at scale**. This is a tool for creating many sharing links
quickly. Used carelessly against the wrong folder with the wrong audience, it can expose a large
amount of content in one action.

Mitigations built into the product:

- Dry run is the default and changes nothing.
- A preview must be produced before Start is available.
- "Anyone with the link" is visibly flagged as the highest-risk audience.
- Every outcome is recorded accurately; a policy-blocked link is never reported as a success.
- Overlapping selections are detected, so a parent selection does not silently process far more
  than intended.

Mitigations that remain **yours**:

- Tenant external sharing policy. This application cannot override it and does not try.
- Sensitivity labels and DLP.
- Which users you allow to consent, or which you grant the permissions to.
- Conditional Access.

Consider restricting the application to a group rather than the whole tenant, using the
enterprise application's user assignment settings.

## Permissions to expect

| Scope | Type | Why |
|---|---|---|
| `User.Read` | Delegated | Identify the signed-in user and tenant |
| `Sites.Read.All` | Delegated | Site discovery and library listing |
| `Files.ReadWrite.All` | Delegated | Enumerate files, create links, write manifests |
| `User.ReadBasic.All` | Delegated, optional | People picker for the User OneDrive source |
| `Sites.ReadWrite.All` | Delegated, optional, broad | Only if a library rejects the standard write |

A read-only deployment is supported: `User.Read`, `Sites.Read.All`, `Files.Read.All`. Browsing
and dry runs work; nothing can be created.

The application deliberately does **not** request `Directory.Read.All`,
`Directory.ReadWrite.All`, or any application-type permission.

### If you want to minimise further

- Deploy read-only first and let users evaluate it.
- Omit `User.ReadBasic.All` unless the User OneDrive source is needed.
- Never grant `Sites.ReadWrite.All` unless a specific library has actually rejected a write.
- Consider `Sites.Selected` — see the caveats in [ENTRA-SETUP.md](ENTRA-SETUP.md) section 7. It
  is genuinely different from scoped-down `Sites.Read.All`: consent alone grants nothing, and
  every site must be assigned individually.

## Deployment options

### Central registration, recommended

Create the app registration yourself, grant consent, and give users the client ID and tenant ID.
Users choose **Use an existing app registration**. You keep full control of the registration and
its permissions, and no user needs the ability to create app registrations.

### Automatic setup

The application creates the registration itself using a publisher-owned bootstrap identity. That
bootstrap identity appears in your tenant as its own enterprise application, which you can
review or block.

This path is **unavailable** unless a publisher has supplied a bootstrap client ID. It is not
present in the open-source repository.

### Blocking it entirely

Block the client ID in your Conditional Access or enterprise application settings, or simply
do not grant consent. Without consent the application cannot obtain a usable token and reports
itself as waiting for administrator approval.

## What it stores, and where

On each user's machine only, under the platform application-data directory:

| Data | Contains a credential? |
|---|---|
| Application settings | No |
| Tenant configuration (tenant ID, client ID) | No |
| Saved job profiles | No |
| Job history | No |
| Local audit of tenant changes | No |
| Log files | No — redaction is applied at the logging boundary |
| MSAL token cache | Yes — held in OS-native secure storage, never a plain file |

Nothing is sent anywhere except Microsoft identity and Microsoft Graph endpoints. Telemetry is
disabled and no telemetry pipeline is implemented.

## Auditing

**In your tenant.** Sharing-link creation appears in the Microsoft 365 unified audit log as
normal sharing activity attributed to the signed-in user. Search for `SharingSet` and related
operations. Every Graph request carries a client request identifier that appears in Microsoft 365
audit and support tooling.

**On the user's machine.** The Permissions tab in Settings shows a local audit history of every change the
application made to your tenant: registration creation, repair, consent requests and deletion,
each with what changed and whether it succeeded. This is a convenience for the user; your tenant
audit log remains authoritative.

## Removing it

| Action | Effect |
|---|---|
| User: *Remove local configuration* | Clears local settings and the token cache. Your tenant is untouched. |
| Admin: revoke consent | Delete the enterprise application, or revoke its permissions. The application stops working immediately. |
| Admin: delete the registration | Removes it for everyone. |

Uninstalling the application never deletes the registration. The application will only delete a
registration it created itself, after the user types the display name to confirm.

## Questions worth asking before approving

1. Which users need this, and can it be scoped to them rather than the tenant?
2. Is read-only sufficient for the first phase?
3. Does our external sharing policy already block the audiences we are concerned about?
4. Are the target sites labelled, and do those labels block the sharing we would not want?
5. Who reviews the manifests once they exist? A manifest is an index of links; it deserves the
   same care as the content it points to.
