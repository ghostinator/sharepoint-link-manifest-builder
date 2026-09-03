# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
