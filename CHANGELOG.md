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

- **Opening the "another user's OneDrive" panel pushed the tree off the window with no way to
  scroll.** The browsers had two `Auto` rows above their star row, so the controls could grow
  without limit and take the tree's space with them. With the expander open the tree collapsed to
  its action bar, and the top-aligned *Selected item* card spilled out of a row that no longer
  had room for it — `Grid` does not clip, so that read as misalignment rather than as the
  overflow it was. Each browser is now one scroll region covering the header, the controls, the
  tree and the details panel, with the action bar pinned below it so the primary action is
  reachable without scrolling back to find it. Applied to the SharePoint browser too, which had
  the identical structure.

- **The browsers' action bar ran into the tree above it.** The bar was a horizontal `StackPanel`,
  which overflows rather than wrapping, and it sat in a `DockPanel`, which does not clip a fill
  child that measures taller than its slot. Both were survivable when each browser had a whole
  window; neither survived being hosted inside a tab. SharePoint was worse than OneDrive purely
  because it has one more button. The bar now wraps, the tree sits in a star row that cannot
  overflow into it, and the 24px page margin the browsers carried from their standalone days is
  gone — it was being applied on top of the job page's own padding and narrowing the tree for
  nothing. The SharePoint side panel also shrinks before the tree does, rather than holding a
  fixed 340px.

- **A Conditional Access device requirement was reported as missing administrator approval.**
  `AADSTS50097` ("Device authentication is required") arrives from the consent endpoint as
  `interaction_required`, which the generic mapping read as "an administrator still needs to
  approve this". It is not that: the organization requires a managed or compliant device, and no
  amount of approving will help from a device that does not qualify. It is now classified
  separately and names the two things that do work — consent from a managed device, or an
  administrator granting it in the Entra admin center.
- **Consent reported failure when consent was not needed.** The consent endpoint can fail for
  reasons unrelated to whether the permissions are granted, while a token carrying every required
  scope has already been issued. The wizard now verifies before reporting a failure, and says so
  plainly when the step turned out to be unnecessary.

### Changed

- **Publisher metadata is filled in.** Contact, security, conduct and privacy addresses, the
  homepage, support, source and issue URLs, and the copyright and company fields now carry real
  values instead of placeholders. Privacy and terms point at `docs/PRIVACY.md` and `LICENSE`,
  which exist, rather than pages nobody has written. The update endpoint is deliberately still
  unset, which is what keeps the update check disabled until there is a release to check against.

- **Loading a profile or a past run now arrives at the job page.** Both buttons name New Link Job
  as their destination, loaded the shared draft, and then stayed where they were, leaving the
  user to work out that anything had happened.
- **Permissions and Diagnostics are sections of Settings.** Both are things you consult about the
  current configuration rather than places you go, and as top-level entries they made three of
  the ten sidebar items settings-shaped. The sidebar is now eight.
- **The update check is unavailable when no update endpoint is configured**, and says so, instead
  of being offered and then reporting that there is nothing to check against. Home's "Check for
  Updates" shortcut is gone: it navigated to About and checked nothing, which the real control on
  About already does properly.

### Added

- **A job can be saved as a profile from the job page.** Saving previously existed only on the
  Saved Profiles page; since it saves the shared draft, the sequence was build the job, navigate
  away, name it, save, navigate back.

- **Opening a OneDrive now collapses the others and expands the new one.** Each drive is a full
  folder tree, so several expanded at once buried the drive that had just been asked for beneath
  however much of the previous ones happened to be showing. Collapsing discards nothing: loaded
  children stay loaded, so re-expanding costs no further Graph calls.

### Added

- **A drive cannot be opened twice, and open drives can be closed.** Opening one that is already
  open reveals it instead of adding a second copy — two subtrees for one drive means two sets of
  checkboxes disagreeing about the same files. Identity is the drive rather than the person
  searched for, so two directory entries resolving to one drive also open once. **Close** on a
  drive row closes that drive; **Close other users' drives** closes all but your own. Neither
  removes targets already added from them: those were an explicit choice, and are removed on the
  Targets step.

- **The resource browsers are steps of a job rather than separate pages.** SharePoint Sites and
  OneDrive have left the sidebar and open as tabs inside New Link Job, from the Targets step,
  each with a **Done — back to targets** action. Selecting what to process is part of building a
  job, and having the browsers as destinations of their own made it a detour with no signposted
  way back — the Targets step said "add locations from the SharePoint Sites or OneDrive pages"
  and then offered no route home. The browsers open inside the Targets step, under a bar that
  appears below the numbered strip only while browsing, so the strip stays exactly six wide and
  choosing locations never moves the selected step.

- **Verification states a single PASS or FAIL.** The verdict came from reading three separate
  lists and inferring one, and "Not checked" in particular read as a failure when it means the
  opposite — there was no independent evidence to gather, because confirming those items would
  need directory permissions this application deliberately does not request. That section now
  says so, and the verdict that matters, taken from a real token, is stated at the top.

- **Automatic setup created a registration and then left the application disconnected.** The
  wizard signed in with the bootstrap identity, created the registration, saved the new
  configuration, and went straight to consent — never signing in to the registration it had just
  created. So consent was requested for an application the user had no grant on, the connection
  never reached Connected because no sign-in with the operating scopes had happened, and the only
  way to finish was to leave the wizard and use the Home page's sign-in button. The wizard now
  signs in to the new registration immediately after creating it, which establishes the user's
  grant and completes the connection.
- **A newly created registration is not usable immediately.** Microsoft Entra replicates it
  first, and until that finishes a sign-in fails with AADSTS700016 — "this app registration does
  not exist" — which is true only for a second or two. The sign-in now retries with backoff over
  roughly half a minute rather than reporting a correctly created registration as broken.
- **A failed consent request logged nothing.** Every failure path returned an error to the UI and
  wrote "ConsentRequested (failed)" to the audit history, leaving the log silent about why. The
  error Microsoft returned, the state-mismatch rejection, and a redirect that returns without
  granting are now each logged.

- **Automatic setup could not create a registration (`Request_BadRequest`).** Every JSON request
  body was serialized in PascalCase: the Graph transport set `PropertyNameCaseInsensitive`, which
  fixes reading, but never set a naming policy, so writes went out as `DisplayName` and
  `PublicClient` rather than `displayName` and `publicClient`. Microsoft Graph is lenient about
  scalar properties, which is why sharing links and manifest writes worked, and strict about
  complex ones, which is why a registration POST was rejected with `Invalid property
  'PublicClient'`. The transport now uses camelCase; properties that are not simply camelCase,
  such as `@microsoft.graph.conflictBehavior`, carry `[JsonPropertyName]` and are unaffected.

- **Consent verification reported a correctly consented tenant as unconsented.** It acquired a
  token silently and concluded from the failure that nobody had consented. Consent an
  administrator has granted in the directory is invisible to a silent request until some
  interactive sign-in records a grant for that user, so AADSTS65001 came back and "Check again"
  changed nothing however many times it was pressed. A user-initiated check now escalates to an
  interactive sign-in when, and only when, the failure means "no cached grant" — a refusal is a
  decision and is not re-prompted.
- **Completing the setup wizard left the application saying "Not connected".** The wizard signed
  in through the authentication service directly, bypassing the coordinator that records granted
  scopes, saves the tenant and moves the application to Connected. The wizard could therefore
  reach "Setup complete, consent granted" while every other page still showed a disconnected
  application, and the only way to actually connect was the Home page's sign-in button. The
  wizard now signs in through the coordinator, which gained an overload for the deliberately
  different scope set automatic setup needs.

- **The per-file results grid was unreachable without enlarging the window.** The results tab
  docked its progress card and manifest list to the top with no height limit, so they took
  whatever they wanted and the grid showing what the job actually did got only the leftovers.
  The tab is now a two-row grid: the top section is capped and scrolls internally, and the grid
  has the remaining space with a floor under it. The manifest list also stretches to the full
  width and starts collapsed, so the grid has room the moment a job finishes.
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
