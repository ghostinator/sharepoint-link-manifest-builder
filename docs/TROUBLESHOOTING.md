# Troubleshooting

Every error the application shows carries a **correlation ID**. Include it when asking for help:
it lets a maintainer, or Microsoft support, trace the exact request.

---

## Sign-in and consent

### The browser said "authentication complete", but the application reports a failure

**This is expected, and it does not mean the sign-in nearly worked.** MSAL renders its
"authentication complete" page as soon as *any* redirect arrives back from Microsoft Entra —
including a redirect that carries an error. The page looks identical either way, so it tells you
only that the browser round-trip finished, not that it succeeded.

The application logs the real cause and shows the Microsoft error code in the message. Look for
`Microsoft error code: AADSTSnnnnn`, or open the log file (Settings, then **Open data folder**,
then `logs/application.log`) and find the most recent line beginning `Could not sign in.`

| Code | Cause | Fix |
| --- | --- | --- |
| `AADSTS7000218` | The registration is not a public client, so Entra demanded a client secret. | In the registration, open **Authentication**, set *Allow public client flows* to **Yes**, and make sure `http://localhost` is registered under **Mobile and desktop applications**, not **Web**. Do **not** create a client secret — this application has none by design. |
| `AADSTS50011`, `AADSTS900971` | The loopback redirect URI is not registered. | Add `http://localhost` under **Mobile and desktop applications**. No port is needed; loopback redirects match on any port. |
| `AADSTS50020` | The account belongs to a different organization than the registration accepts. | Sign in with an account in the configured organization, or turn on multi-organization mode (see below). |
| `AADSTS700016`, `AADSTS90002` | The registration does not exist in the organization signed in to. | Check the client ID, and confirm the organization. A single-organization registration cannot be used from another organization. |
| `AADSTS65001` | Consent has not been granted. | Use **Request Missing Consent** and have an administrator approve it. |

The single most common cause is the first row: adding the redirect URI under the **Web** platform
instead of **Mobile and desktop applications**. Entra then treats the application as a
confidential client and requires a secret at the token-exchange step, which happens *after* the
browser has already reported success.

### Using one installation across several organizations

Turn on **This registration works with any work or school organization** on the sign-in page of
the setup wizard. The Directory (tenant) ID then becomes optional: the organization is taken from
whichever account you sign in with.

This requires the app registration itself to be set to **Accounts in any organizational
directory** (`AzureADMultipleOrgs`) in Microsoft Entra. That cannot be changed after the
registration is created through this application's automatic setup, because changing it requires
a far broader permission than the create-only tier this product asks for. If you already have a
single-organization registration, either change it in the Entra portal or create a new one.

Once it is on, the home page lists every organization signed in on this device, and selecting one
switches the whole application to it. Switching is silent when the cached token is still valid.

**Each organization consents separately.** Consent in one organization grants nothing in another.
Personal Microsoft accounts are not supported at all, because they have no SharePoint or OneDrive
for Business.

### "An authorized Microsoft Entra administrator must approve the requested permissions"

The permissions need administrator consent and you are not an administrator, or your tenant
requires admin consent for all applications.

**What to do.** On the Consent page or the Permissions page, use **Copy consent link** and send
it to an authorized administrator. Save the configuration as pending approval, then use **Check
again** once they confirm.

### "You signed in to a different Microsoft 365 organization than this application is configured for"

You are signed into more than one directory and picked the wrong one.

**What to do.** Sign in again and select an account in the tenant shown in the wizard. This is a
deliberate refusal: accepting a token from another tenant would be a cross-tenant data risk.

If you genuinely need to work across several organizations, turn on multi-organization mode
rather than working around this check.

### "The consent response could not be verified and was rejected"

The `state` value returned by the browser did not match the one this application generated.
Nothing was changed.

**What to do.** Start the consent step again from the application. If it recurs, check whether
something is intercepting or rewriting your browser redirects.

### "Your organization requires an additional step, such as multi-factor authentication"

Conditional Access needs interaction.

**What to do.** Sign in again and complete the prompts. If it repeats immediately, the device may
not satisfy a compliance requirement; your administrator can check the sign-in log.

### Consent seemed to succeed, but verification says permissions are missing

This is the application being honest. The redirect looked successful, but the token Microsoft
Entra actually issued did not contain every required scope.

**What to do.** Check the Permissions page for exactly which scopes are missing. Common causes: an
administrator consented to a subset; a Conditional Access policy restricts the application; or
consent was granted in a different tenant.

### Automatic setup is greyed out

Automatic setup needs a publisher-supplied bootstrap client ID, which this build does not
include.

**What to do.** Use **Use an existing app registration**. See
[ENTRA-SETUP.md](ENTRA-SETUP.md) section 4 for creating one; it takes about two minutes.

### "Your account is not permitted to create or change application registrations"

Your tenant restricts app registration, which is a common and reasonable policy.

**What to do.** Ask an administrator to create the registration, then use the
existing-registration path.

---

## Browsing

### A site does not appear in search

`GET /sites?search=` returns what the search index exposes to your account. It is not a complete
list of tenant sites, and the application does not claim otherwise.

**What to do.** Paste the site URL into the *Open URL* box. That resolves the site directly
rather than through the index.

### "That user does not have a OneDrive yet"

The user has never opened OneDrive, so it has not been provisioned.

**What to do.** The user must open OneDrive once. This application deliberately does not
provision a OneDrive on someone else's behalf.

### "You do not have permission to read that user's OneDrive"

Your own account cannot open it.

**What to do.** Nothing in this application can change that. Delegated access is bounded by your
own permissions, and administrator consent does not grant access to everyone's files. Ask the
owner to share it with you, or use an account that already has access.

### A folder shows an error with a Retry button

That folder could not be listed, usually a permissions boundary inside an otherwise readable
library.

**What to do.** Retry in case it was transient. If it persists, you do not have access to that
subtree. A job will record it as skipped and continue with everything else.

### A pasted URL resolves to the wrong place

Some SharePoint URLs carry the real folder in a query parameter rather than the path. The
application handles the common `_layouts` forms, but the SharePoint UI produces several.

**What to do.** Navigate into the folder in SharePoint and copy the address bar URL, or browse to
it in the tree instead.

---

## Running a job

### "Your organization's sharing policy does not allow this kind of link"

Tenant policy refused the request. Most often *Anyone with the link* on a site where anonymous
sharing is disabled.

**What to do.** Choose *People in the organization*, or ask an administrator about the external
sharing policy for that site. This application cannot override policy and will not pretend to.

### Everything reports "Access denied"

Either the required permission is missing, or your account cannot share those items.

**What to do.** Check the Permissions page first. If `Files.ReadWrite.All` is granted, the issue
is item-level: you may be able to read a library without being allowed to share from it.

### The job is very slow, and progress mentions throttling

Microsoft 365 is asking the application to slow down. It honours the requested wait and resumes
automatically.

**What to do.** Nothing is wrong. To reduce it: lower **Maximum concurrency**, add a **delay
between requests**, or run against fewer targets at once. Throttling is tenant-wide, so other
activity in your organization affects it too.

### Results say "Reused" but I expected "Created"

An equivalent link already existed and Microsoft Graph returned it rather than creating a
duplicate. The link works; it simply is not new.

The distinction is deliberate: reporting a reused link as created would misrepresent what the
job did.

### The job was cancelled but manifests were still written

Intentional. Links created before cancellation are real, so the successes are recorded rather
than discarded.

### "The manifest changed in SharePoint after this application read it"

Someone or something modified the manifest between the read and the write, so the write was
refused rather than overwriting their change.

**What to do.** Run the job again; it will read the current version first.

### A timestamped manifest appeared instead of the expected name

There was already a file at that path that this application did not write, so it was left
untouched and a timestamped copy was written alongside.

**What to do.** Check the existing file. If it is safe to replace, set the conflict policy to
*Replace*.

---

## Local environment

### "Secure storage could not be initialised"

Usually Linux without a running keyring.

**What to do.** Install and start a Secret Service provider (`gnome-keyring`, `kwallet`, or
similar) plus `libsecret`. Until then the application works, but you will sign in again each
launch. It will never write tokens to disk unprotected.

### The application will not start

Check for a `startup-crash-*.log` file next to the executable.

**What to do.** Confirm the .NET runtime requirement is met (self-contained builds include it).
On macOS, an unsigned application may be blocked by Gatekeeper: see
[PACKAGING.md](PACKAGING.md).

### macOS says the application is damaged or from an unidentified developer

Release artifacts are unsigned and un-notarized, and Gatekeeper blocks them by default.

**What to do.** Right-click the application and choose *Open*, then confirm. Only do this if you
trust the source and have verified the SHA-256 checksum.

### Windows SmartScreen warns about the download

Same cause: the artifact is unsigned.

**What to do.** Verify the checksum against `SHA256SUMS.txt`, then choose *More info* and *Run
anyway* if you trust the source.

---

## Getting more information

1. **Diagnostics page** — run the connectivity test, review sanitized recent errors.
2. **Open log folder** — logs are redacted, so no token appears in them.
3. **Export a diagnostic bundle** — the page lists exactly what it will include before writing.

When reporting a problem, include the application version, platform, the error message, and the
correlation ID. Do **not** include tokens, sharing links, real URLs, tenant identifiers or
personal data. See [SUPPORT.md](SUPPORT.md).

## The build fails with "Access to the path ... is denied" (MSB3026 / MSB3027)

MSBuild retries the copy ten times and then fails. The message names a permission problem, but
the permissions are usually fine — POSIX has no distinct error for "another process is holding
this file", so contention surfaces as the same `EACCES` as a genuine permission fault.

The common cause is a clone inside a cloud-synced folder: OneDrive, iCloud Drive, Dropbox or
Google Drive. A .NET build writes thousands of small files, and a sync client that wants to
upload each one will intermittently hold a freshly written assembly. The same build then
succeeds moments later, which is what makes it look random.

`scripts/common.sh` detects a clone inside a known sync root and redirects build output to
`~/.cache/splmb-build/<repo>`, printing a warning when it does. Set `SPLMB_ARTIFACTS_PATH` to
choose a different location. To build a single project by hand:

```bash
dotnet build src/SharePointLinkManifestBuilder.App -c Release -p:ArtifactsPath="$HOME/.cache/splmb-build"
```

The durable fix is to keep the clone outside the synced folder entirely.

A second cause, worth ruling out: two builds running at once against the same output directory —
an editor building in the background while a script builds in a terminal, for instance. MSBuild
assumes a single writer per `bin/`.
