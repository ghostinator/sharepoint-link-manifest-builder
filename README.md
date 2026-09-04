# SharePoint Link Manifest Builder

A cross-platform desktop application that builds explicit, reliable manifests of SharePoint and
OneDrive file links, so Microsoft Copilot or another authorized AI system can be given a precise
list of files instead of having to discover them.

Select sites, document libraries, folders, your own OneDrive, or an accessible user's OneDrive
through a point-and-click tree. Preview exactly what will be processed. Create sharing links
using the permission and audience you choose, where tenant policy allows. Write the resulting
manifests back into SharePoint or OneDrive.

> **Status.** The core flow has been run end to end against a real Microsoft 365 tenant on
> macOS: creating the app registration, signing in, browsing SharePoint and OneDrive, creating
> sharing links, and writing manifests back. Windows and Linux are built and tested in CI but
> have not been exercised against a tenant, and there is no released build yet. See
> [Limitations](#limitations) for what is and is not verified.

---

## Contents

- [What it does](#what-it-does)
- [Supported platforms](#supported-platforms)
- [Architecture](#architecture)
- [Screenshots](#screenshots)
- [Prerequisites](#prerequisites)
- [Build, test and run](#build-test-and-run)
- [Microsoft Entra setup](#microsoft-entra-setup)
- [Permissions](#permissions)
- [Security model](#security-model)
- [Limitations](#limitations)
- [Packaging](#packaging)
- [Contributing](#contributing)
- [Documentation](#documentation)

---

## What it does

- **Graphical selection.** Expand site to document library to folder to subfolder in a lazy
  tree. Nothing enumerates a tenant to draw a screen. Multi-select with tri-state checkboxes,
  or paste a SharePoint URL directly.
- **Mixed targets.** One job can combine SharePoint sites, individual libraries, specific
  folders, your own OneDrive and another user's OneDrive. Recursion is set per target.
- **Overlap handling.** Selecting a folder and its parent is detected and reconciled, so a file
  is processed once. Recursion is taken into account: a non-recursive parent does not cover a
  subfolder.
- **Honest previews.** Preview separates what was actually validated from what is merely
  expected and from what cannot be known until execution. Tenant sharing policy falls in the
  last category, and the application says so rather than showing a green tick.
- **Dry run.** Enumerate, filter and validate while creating no link and writing no file.
- **Honest results.** Microsoft Graph returns HTTP 201 for a new sharing link and 200 when an
  equivalent one already existed. This application records those as `Created` and `Reused` and
  never conflates them. A link refused by policy is recorded as blocked, not as a failure of
  unknown cause and never as a success.
- **Manifests.** Per-folder, master, or both, in plain text, Markdown, CSV or JSON. Existing
  manifests are updated in place using `(driveId, itemId)` identity, so renames and moves are
  handled correctly. A file this application did not write is never overwritten.
- **Retry and cancel.** Cancellation preserves completed results and still writes manifests for
  them. Failures caused by throttling or a transient fault can be retried; failures caused by
  policy cannot, and are not offered.

## Supported platforms

| Platform | Architectures | Secure token storage |
|---|---|---|
| Windows 10/11 | x64, ARM64 | DPAPI-protected file |
| macOS 12+ | x64 (Intel), ARM64 (Apple Silicon) | Keychain |
| Linux | x64, ARM64 | Secret Service (libsecret) keyring |

Where secure storage is unavailable — typically a Linux session with no keyring — the
application says so and keeps tokens in memory only. It never writes them to disk unprotected.

## Architecture

```
SharePointLinkManifestBuilder.App     Avalonia 12 + MVVM. Views, view models, composition root.
        |
SharePointLinkManifestBuilder.Graph   MSAL authentication, the Graph HTTP transport,
        |                             resource services, Entra onboarding and consent.
        |
SharePointLinkManifestBuilder.Core    Domain model, URL parsing, filtering, target planning,
                                      manifest engine, retry policy, redaction, job runner.
                                      No HTTP. No MSAL. No UI.
```

The dependency direction is one-way and enforced by project references. Every rule a reviewer
would want to check by reading lives in `Core` and is unit-tested without a network.
See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) and the decision records in
[docs/adr](docs/adr).

## Screenshots

Not yet included. A screenshot of this application necessarily shows a real tenant, so they
must be captured against a dedicated test tenant with synthetic content. See
[docs/screenshots/README.md](docs/screenshots/README.md).

## Prerequisites

- [.NET SDK 10.0](https://dotnet.microsoft.com/download) or newer
- A Microsoft 365 work or school account
- An application registration in your tenant, or permission to create one
  (see [Microsoft Entra setup](#microsoft-entra-setup))

On Linux, secure token storage additionally needs `libsecret` and a running Secret Service
provider such as `gnome-keyring`.

## Build, test and run

```bash
./scripts/restore.sh          # restore packages
./scripts/build.sh            # build, warnings as errors
./scripts/test.sh             # run every test; no tenant or credential needed
./scripts/format.sh --check   # verify formatting
./scripts/run.sh              # run the application
```

The scripts detect a clone inside a cloud-sync folder (OneDrive, iCloud Drive, Dropbox, Google
Drive) and redirect build output to `~/.cache/splmb-build`, because a sync client that wants to
upload each of the thousands of files a .NET build writes will intermittently hold one and fail
the copy. Set `SPLMB_ARTIFACTS_PATH` to choose a different location.

Run only one build at a time. Two concurrent builds of the same project write to the same output
directory and collide, which MSBuild reports as `MSB3026` and "Access to the path is denied" —
a message that reads like a permissions problem and is not one.

Or with the SDK directly:

```bash
dotnet restore SharePointLinkManifestBuilder.slnx
dotnet build SharePointLinkManifestBuilder.slnx --configuration Release -warnaserror
dotnet test SharePointLinkManifestBuilder.slnx --configuration Release
dotnet run --project src/SharePointLinkManifestBuilder.App
```

## Microsoft Entra setup

Two paths, both graphical. Full detail in [docs/ENTRA-SETUP.md](docs/ENTRA-SETUP.md).

**Existing app registration** — always available. Your administrators create a registration and
you supply its client ID and your tenant ID. Recommended where self-service app registration is
restricted.

**Automatic tenant setup** — the application creates the registration for you. It requires a
publisher-supplied bootstrap client ID, which **this repository does not ship**. Automatic setup
reports itself unavailable until one is configured; no client ID is fabricated or committed.

Either way:

- The desktop application is a **public client** and never uses a client secret.
- Sign-in and consent happen in your **system browser**, in Microsoft's own experience.
- Consent is **verified** by acquiring a real token, not assumed from a redirect.

Creating the registration by hand takes about two minutes:

```bash
az ad app create \
  --display-name "SharePoint Link Manifest Builder" \
  --sign-in-audience AzureADMyOrg \
  --is-fallback-public-client true \
  --public-client-redirect-uris "http://localhost"
```

Then add the delegated Microsoft Graph permissions below and grant consent. Do not create a
client secret.

## Permissions

Delegated only. The application acts as you and can never reach content your own account cannot
already open.

| Scope | Why | Admin consent |
|---|---|---|
| `User.Read` | Identify the signed-in user and tenant | No |
| `Sites.Read.All` | Find sites, read metadata, list document libraries | Yes |
| `Files.ReadWrite.All` | Enumerate files, create sharing links, write manifests | Yes |
| `User.ReadBasic.All` | *Optional.* The User OneDrive people picker | Yes |
| `Sites.ReadWrite.All` | *Optional and broad.* Only if a library rejects the standard write | Yes |

A read-only profile (`User.Read`, `Sites.Read.All`, `Files.Read.All`) supports browsing and dry
runs with no write capability at all.

**Administrator consent does not** grant access to content you cannot already open, override
SharePoint or OneDrive permissions, override external sharing policy or sensitivity labels, or
enable unattended access. This application has no application-only mode.

## Security model

- **No client secret.** Public client with Authorization Code Flow and PKCE.
- **No embedded web view.** Authentication and consent always use the system browser, so this
  application never sees a credential.
- **Tokens in OS-native secure storage**, with a visible memory-only fallback.
- **Redaction at the logging boundary**, so no call site — including third-party libraries —
  can write a token to a log.
- **Tenant-specific authorities**, never `/common`. A token issued for a different tenant than
  the configured one is rejected.
- **State validation** on the consent redirect.
- **Least privilege**, tiered and opt-in. Broad scopes are labelled broad, with the narrower
  alternative named.
- **Nothing silent.** Every change to your tenant is shown before it happens and recorded in a
  local audit history.
- **CSV formula injection, Markdown injection and path traversal** are all defended against,
  because file names come from SharePoint and are attacker-influenced.

The full analysis, with residual risks, is in [docs/THREAT-MODEL.md](docs/THREAT-MODEL.md).

## Limitations

Stated plainly rather than buried.

1. **Live-tenant validation is partial.** The core flow works against a real tenant on macOS:
   registration, sign-in, consent, browsing, link creation and manifest writing. Not exercised
   against a tenant: **Windows and Linux** (built and tested in CI only), sovereign clouds,
   throttling under load, and switching between organizations. The 478 tests run against a
   mocked transport, so anything not listed as verified above rests on the Microsoft Graph v1.0
   reference rather than observed behaviour.
2. **There is no released build.** Clone and build it yourself; see
   [Build, test and run](#build-test-and-run).
3. **Automatic setup needs a bootstrap client ID.** This repository ships none, deliberately: a
   private identifier committed to source control is a supply-chain problem. Supply your own
   under *Advanced*, or use an existing app registration.
4. **Artifacts are unsigned and un-notarized.** No signing credentials exist here, so a
   self-built application will be blocked by Gatekeeper on macOS until you allow it.
5. **Publisher metadata is a placeholder**, including the privacy policy URL shown on the
   consent screen.
6. **Site search is not exhaustive.** `GET /sites?search=` returns what the search index exposes
   to the signed-in user. The UI says so and offers URL pasting for exact resolution.
7. **Consent type cannot always be determined.** A token proves consent exists but not whether
   an administrator granted it tenant-wide. This is reported as unknown rather than guessed.
8. **Telemetry is not implemented.** The opt-in setting exists; no pipeline does.
9. **`Sites.Selected` is documented, not fully implemented.** Graphical site assignment is not
   provided.
10. **Embed links** are OneDrive-personal only and are not offered in the UI.
11. **Stale links are not tracked.** A manifest records a link at a point in time; it cannot know
    an administrator later revoked it.

## Packaging

```bash
./scripts/package.sh            # all six runtime identifiers
./scripts/package.sh osx-arm64  # one
./scripts/checksums.sh          # SHA-256 for every artifact
./scripts/sbom.sh               # dependency inventory and vulnerability report
```

Artifacts are self-contained single-file builds and are **not signed**. See
[docs/PACKAGING.md](docs/PACKAGING.md).

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Before opening a pull request:

```bash
./scripts/format.sh --check
./scripts/build.sh
./scripts/test.sh
./scripts/scan-secrets.sh
```

Never commit a token, secret, certificate, real tenant identifier, real SharePoint URL, sharing
link, or personal data. The publication safety gate scans the working tree and the full history,
and it fails closed.

## Documentation

| Document | Purpose |
|---|---|
| [PLAN.md](PLAN.md) | Goal, constraints and phased delivery |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Layering, data flow, concurrency, testing seams |
| [docs/USER-GUIDE.md](docs/USER-GUIDE.md) | Using the application |
| [docs/ADMIN-GUIDE.md](docs/ADMIN-GUIDE.md) | Deployment and governance |
| [docs/ENTRA-SETUP.md](docs/ENTRA-SETUP.md) | Registration, permissions and consent |
| [docs/GRAPH-OPERATIONS.md](docs/GRAPH-OPERATIONS.md) | Every Graph call, with permissions |
| [docs/THREAT-MODEL.md](docs/THREAT-MODEL.md) | Threats, mitigations and residual risk |
| [docs/PRIVACY.md](docs/PRIVACY.md) | What is stored, where, and what is never sent |
| [docs/DATA-FLOW.md](docs/DATA-FLOW.md) | Where data travels |
| [docs/DIAGNOSTICS.md](docs/DIAGNOSTICS.md) | Diagnostics and the sanitized bundle |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | Errors and what to do about them |
| [docs/PACKAGING.md](docs/PACKAGING.md) | Building, signing and distributing |
| [docs/RELEASE-PROCESS.md](docs/RELEASE-PROCESS.md) | Cutting a release |
| [docs/ROLLBACK.md](docs/ROLLBACK.md) | Reverting code and application versions |
| [docs/GITHUB-PUBLISHING.md](docs/GITHUB-PUBLISHING.md) | Publishing this repository safely |
| [docs/SUPPORT.md](docs/SUPPORT.md) | Getting help |
| [docs/ACCEPTANCE.md](docs/ACCEPTANCE.md) | Acceptance criteria and their status |
| [docs/adr](docs/adr) | Architecture decision records |

## Licence

GNU General Public License, version 3 or later. The full text is in [LICENSE](LICENSE), and the
reasoning is recorded in [LICENSE-SELECTION.md](LICENSE-SELECTION.md).

This is a copyleft licence: you may use, study, modify and redistribute this software, and any
distributed derivative must also be GPL-3.0-or-later and ship its source. Every third-party
dependency is MIT, BSD-3-Clause or Apache-2.0, all of which are compatible with GPL-3.0 --
Apache-2.0 is compatible with version 3 but not version 2, which is part of why version 3 was
chosen. See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
