# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **Licence: GNU General Public License version 3 or later.** `LICENSE` holds the verbatim text
  and `LICENSE-SELECTION.md` records why version 3 was chosen — two dependencies are Apache-2.0,
  which is compatible with GPLv3 but not GPLv2.
- **Back and Next buttons beside the job step tabs**, with a "Step N of 6" position label.
- **An explanation next to the Start button** saying why the job cannot start, or that dry run
  is on and will change nothing. A disabled button previously said something was wrong but never
  what.

- **An explicit opt-in for a broader bootstrap permission.** `AppRegistration.Create` is the
  least-privileged permission documented for `POST /applications` and remains the default, but
  some tenants' portal permission picker does not list it yet. Advanced now offers
  `Application.ReadWrite.All` as a deliberate, warned, logged choice rather than a silent
  fallback, since it can modify and delete every registration in the directory.

- Optional multi-organization mode. One installation can now be used across several Microsoft
  365 organizations with a registration set to *Accounts in any organizational directory*,
  using the `/organizations` authority. Never `/common`, which would also admit personal
  Microsoft accounts. Default behaviour is unchanged and remains single-organization. See
  [ADR-0011](docs/adr/0011-multi-tenant-authority.md).
- Organization switcher on the home page, listing every organization signed in on this device.
  Switching is silent when the cached token is still valid, and prompts only when the target
  organization has not consented yet.
- Administrator-consent requests now name an explicit organization even in multi-organization
  mode, and are refused rather than guessed when the organization is not yet known.

### Fixed

- **Long lists on the Preview and Results tabs consumed the whole page.** Manifests written,
  preflight warnings and preflight blockers are now collapsible and internally scrollable, so a
  job that writes one manifest per folder no longer pushes the progress card and results grid
  out of view.
- **Start job now sits beside Build preview** on the Preview tab, rather than at the bottom of
  the Results tab, so the build-then-run sequence is in one place.

- **Builds failed intermittently inside cloud-synced folders.** A clone under OneDrive, iCloud
  Drive, Dropbox or Google Drive hits `MSB3026`/`MSB3027` ("Access to the path ... is denied")
  when the sync client holds a freshly written assembly. `scripts/common.sh` now detects a clone
  inside a known sync root and redirects build output to `~/.cache/splmb-build/<repo>`, with a
  warning saying so; `SPLMB_ARTIFACTS_PATH` overrides the location. Clones outside a sync root,
  including CI, keep the standard layout.
- **A sign-in that never returned from the browser wedged the setup wizard.** Interactive
  sign-in and administrator consent both complete only when Microsoft Entra redirects back to
  the loopback listener. When Entra instead shows an error *in the browser* — a wrong-but-valid
  client ID being the common case — no redirect is ever sent, so the awaited task never
  finished. The generated `AsyncRelayCommand` reports `CanExecute` as false while it runs, which
  disabled the very button needed to retry, and `CanGoBack` is gated on `IsBusy`, which disabled
  Back as well. Correcting the IDs left no way to act on the correction. Both steps are now
  cancellable from the UI and bounded by a five-minute timeout, and a timeout is reported
  distinctly from a cancellation.
- **Automatic tenant setup could not be selected.** The method was disabled unless a bootstrap
  client ID was already configured, but the Advanced field that supplies one sits on the
  automatic path — so the only route to enabling automatic setup was reachable only after it was
  already enabled. Since this repository ships no bootstrap client ID by design, automatic setup
  was unreachable in every build. The method is now always selectable and the run is guarded
  instead, which is where the check belonged.
- **Sign-in failures were undiagnosable.** MSAL exceptions were normalized into a sanitized
  user-facing error without being logged first, so the OAuth error code, the `AADSTSnnnnn`
  code and the correlation ID were all discarded. Any unclassified failure became the single
  sentence "Microsoft Entra refused the request to sign in", with nothing in the log. The full
  diagnostic is now logged before normalization, and the Microsoft error code is shown in the
  message.
- Classified the failures that occur *after* the browser displays "authentication complete":
  a registration that is not a public client (`AADSTS7000218`), an unregistered loopback
  redirect (`AADSTS50011`, `AADSTS900971`), an account from another organization
  (`AADSTS50020`), and a registration missing from the signed-in organization
  (`AADSTS700016`, `AADSTS90002`). Each now names its own remedy instead of falling through to
  a generic message.

### Added

- Cross-platform Avalonia desktop application for Windows, macOS and Linux.
- Graphical first-run setup wizard with eight pages: welcome, method, sign-in, permission
  review, provisioning, consent, verification and completion.
- Two onboarding paths: automatic tenant setup, and an existing app registration. The
  existing-registration path is always available.
- Graphical SharePoint selector with search, URL pasting, and a lazy site to library to folder
  tree with tri-state multi-selection.
- OneDrive selector for the signed-in user's own drive and, where permitted, another user's.
- Mixed processing targets across SharePoint and OneDrive, with per-target recursion.
- Overlap detection and reconciliation that accounts for recursion.
- Dry-run mode that creates no link and writes no file.
- Preview separating validated conditions from expected ones and from those unknowable before
  execution.
- Sharing-link creation with View and Edit permissions and Organization, Specific People and
  Anyone audiences, recording the actual outcome rather than an assumed one.
- Per-folder and master manifests in plain text, Markdown, CSV and JSON.
- Manifest update mode keyed on `(driveId, itemId)`, handling renames and moves.
- Manifest conflict policies with ETag concurrency; a foreign file is never overwritten.
- Job history, saved profiles, and a local audit history of tenant modifications.
- Diagnostics page with a connectivity test and a sanitized bundle whose contents are declared
  before export.
- Publication safety gate scanning the working tree and full git history.
- 383 tests requiring no live tenant.

### Security

- Public client with Authorization Code Flow and PKCE; no client secret anywhere.
- System browser only for authentication and consent; no embedded web view.
- OS-native secure token storage with an explicit memory-only fallback.
- Redaction applied at the logging provider boundary.
- Tenant-specific authorities; cross-tenant tokens rejected.
- Cryptographically random consent state, validated on return.
- CSV formula injection, Markdown injection and path traversal defences.

### Known limitations

- No Graph operation has been executed against a live Microsoft 365 tenant.
- Automatic tenant setup is unavailable until a publisher supplies a bootstrap client ID.
- Release artifacts are unsigned and un-notarized.
- Publisher metadata, including the privacy policy URL, is a placeholder.
- Telemetry is not implemented; the opt-in setting exists but no pipeline does.
- `Sites.Selected` is documented but graphical site assignment is not implemented.

[Unreleased]: https://example.invalid/PLACEHOLDER-SOURCE/compare/main...HEAD
