# Acceptance criteria

Each criterion from the product brief, with its status and where it is satisfied.

**Legend.** ✅ done and verified locally · ⚠️ implemented but requires an external input or a live
tenant to verify · ❌ not done

| # | Criterion | Status | Where |
|---|---|---|---|
| 1 | The application builds | ✅ | `dotnet build`, 0 warnings, 0 errors |
| 2 | Automated tests pass | ✅ | 383 tests, all passing |
| 3 | The application launches | ✅ | `ApplicationLaunchTests`, headless Avalonia |
| 4 | First-run setup wizard exists | ✅ | `TenantSetupViewModel`, `TenantSetupView`, 8 pages |
| 5 | Automatic and existing-registration paths exist | ✅ | `SetupMethod`; existing path fully functional |
| 6 | Automatic registration works given a bootstrap identity | ⚠️ | `AppRegistrationService` implemented and unit-tested; needs a bootstrap client ID and a live tenant |
| 7 | No client secret in the desktop app | ✅ | ADR-0004; asserted by `CreateRegistration_NeverRequestsAPasswordCredential` |
| 8 | Microsoft-hosted administrator consent | ✅ | `ConsentService` builds official URLs only; system browser |
| 9 | Consent verified, not assumed | ✅ | ADR-0006; `VerifyConsentAsync` acquires a real token |
| 10 | Pending administrator approval supported | ✅ | `ConsentState.PendingAdministratorApproval`, Copy consent link, Check again |
| 11 | User can sign in with Microsoft identity | ⚠️ | `MsalAuthenticationService`; interactive flow needs a live tenant |
| 12 | User can switch tenant and account | ✅ | `ConnectionCoordinator`, Settings, Permissions |
| 13 | Graphical SharePoint selector | ✅ | `SharePointBrowserView` |
| 14 | Selecting a site loads its libraries | ✅ | `LoadSiteDrivesAsync`; `GetSiteDrives_ReturnsEveryLibrary` |
| 15 | Libraries and folders browsable graphically | ✅ | Lazy `TreeView` |
| 16 | My OneDrive browsable | ✅ | `OneDriveBrowserViewModel.LoadMyDriveAsync` |
| 17 | Accessible user OneDrive browsable | ✅ | `OpenUserDriveAsync`; unavailable cases reported specifically |
| 18 | No terminal, PowerShell, Graph Explorer or manual IDs | ✅ | IDs appear only under an Advanced expander |
| 19 | Mixed SharePoint and OneDrive targets | ✅ | `JobDraft.Targets`, `TargetSourceType` |
| 20 | Recursion configurable per target | ✅ | `ProcessingTarget.Recursive` |
| 21 | Non-recursive handles direct children only | ✅ | `NonRecursive_ProcessesOnlyDirectChildren` |
| 22 | Pagination handled | ✅ | `EnumeratePagedAsync`; multi-page tests |
| 23 | Throttling handled | ✅ | `GraphRetryPolicy`; Retry-After honoured as delay and date |
| 24 | Duplicate and overlapping targets handled | ✅ | `TargetPlanner`; recursion-aware |
| 25 | Candidate files previewable | ✅ | `PreviewAsync`, Preview tab |
| 26 | Dry run modifies nothing | ✅ | `RunAsync` returns before any write; asserted in the runner |
| 27 | View and Edit links | ✅ | `LinkPermission` |
| 28 | Organization / Specific People / Anyone handled per real behaviour | ✅ | createLink for scope; invite for recipients; 207 parsed |
| 29 | Per-folder manifests | ✅ | `BuildPerFolderManifests` |
| 30 | Master manifests | ✅ | `BuildMasterManifest`, `BuildCombinedMasterManifest` |
| 31 | Plain text enabled by default | ✅ | `ManifestConfiguration.Formats` |
| 32 | Generated manifests excluded by default | ✅ | `IsGeneratedManifestName`; asserted in discovery tests |
| 33 | Existing manifests updated safely | ✅ | `ManifestMerger`, `ManifestConflictResolver`, ETag |
| 34 | Partial failures preserve successes | ✅ | `ContinueOnError`; cancellation still writes manifests |
| 35 | Failed files retryable | ✅ | `RetryFailuresAsync`; only retryable failures offered |
| 36 | Cancellation works safely | ✅ | Cancellation tests in Core and Graph |
| 37 | Logs contain no tokens | ✅ | `RedactingLoggerProvider` wraps every provider; unit-tested |
| 38 | Tokens stored securely or in memory | ✅ | ADR-0008; probe with a real round-trip |
| 39 | No hard-coded tenant configuration | ✅ | Publication safety scan; no client ID shipped |
| 40 | Meaningful git commits | ✅ | Focused conventional commits |
| 41 | GitHub-ready workflows | ✅ | 5 workflows, actions pinned to verified SHAs |
| 42 | Publication safety scan exists | ✅ | `scripts/scan-secrets.sh`; verified against planted secrets |
| 43 | CI uses no live-tenant credentials | ✅ | No secret referenced in any workflow |
| 44 | Cross-platform packages configured | ⚠️ | Scripts and workflow for 6 RIDs; only `osx-arm64` built on this machine |
| 45 | Rollback documented | ✅ | `docs/ROLLBACK.md` |
| 46 | Another organization can follow the documentation | ✅ | `docs/ENTRA-SETUP.md`, `docs/ADMIN-GUIDE.md`; no tenant-specific values |
| 47 | No private SharePoint data in the repository | ✅ | Safety scan passes |
| 48 | No secrets in the repository | ✅ | Safety scan passes over full history |
| 49 | Keyboard accessible | ✅ | Automation names, focus ring, live regions; headless render test |
| 50 | Known limitations reported explicitly | ✅ | `README.md`, this document |

## Requiring an external input

| Need | Blocks | Consequence |
|---|---|---|
| Publisher bootstrap client ID | #6 | Automatic setup unavailable; existing-registration path unaffected |
| A live Microsoft 365 tenant | #6, #11, and real-world verification of every Graph call | Behaviour verified against the documented contract only |
| Code-signing certificate | Signed Windows artifacts | Artifacts labelled unsigned |
| Apple Developer ID and notarization | Signed macOS artifacts | Artifacts labelled unsigned |
| Publisher identity and URLs | Consent-screen branding, About page | PLACEHOLDER shown |
| Privacy policy URL | Consent screen | PLACEHOLDER shown |
| A licence decision | Legal redistribution | No licence file; default copyright applies |

## Not verified against a live tenant

Every Graph request shape was written against the Microsoft Graph v1.0 reference and is
exercised against a scripted transport. The following can only be confirmed against a real
tenant:

1. Interactive sign-in, MFA and Conditional Access behaviour
2. Actual tenant policy responses to each link audience
3. Whether `AppRegistration.Create` alone suffices to create the registration in a given tenant
4. Whether consent provisions the service principal as expected
5. Real throttling thresholds and `Retry-After` values
6. Behaviour of very large libraries
7. Sovereign-cloud endpoints
8. Secure storage on Windows and Linux (macOS Keychain is the only platform this build ran on)
