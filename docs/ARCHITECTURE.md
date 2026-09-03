# Architecture

## 1. Layering

```
+--------------------------------------------------------------+
| SharePointLinkManifestBuilder.App        (Avalonia 12, MVVM)  |
|  Views (.axaml)  ViewModels  Navigation  DI composition root  |
+---------------------------+----------------------------------+
                            | interfaces only
+---------------------------v----------------------------------+
| SharePointLinkManifestBuilder.Graph                           |
|  MSAL authentication      GraphApiClient (HttpClient)         |
|  Site/Drive/User/Sharing/ManifestStorage services             |
|  Entra registration, consent, verification                    |
+---------------------------+----------------------------------+
                            | interfaces + models
+---------------------------v----------------------------------+
| SharePointLinkManifestBuilder.Core                            |
|  Models  Abstractions  Url parsing  Filtering  Target planning|
|  Manifest format/parse/merge  Retry policy  Redaction  Jobs   |
|  No HTTP. No MSAL. No UI.                                     |
+--------------------------------------------------------------+
```

The dependency direction is enforced by project references. `Core` has no knowledge of HTTP
or Microsoft identity; it is pure, deterministic, and therefore heavily unit-tested.

### Why the seam is where it is
Every rule that a reviewer would want to check by reading — how a non-recursive target is
expanded, whether an overlapping selection processes a file twice, what a manifest looks like
byte-for-byte, whether a CSV cell can start a formula, whether a retry honours `Retry-After` —
lives in `Core` and is testable without a network. `Graph` owns exactly one hard thing:
turning those decisions into correct HTTP.

## 2. Key abstractions (all in `Core.Abstractions`)

| Interface | Responsibility |
|---|---|
| `IAuthenticationService` | Sign in/out, switch account and tenant, acquire scoped tokens |
| `ISecureTokenStorage` | Probe and provide OS-native token storage; report availability |
| `IGraphApiClient` | The single Graph transport: send, paginate, retry, normalize errors |
| `ISiteService` | Site search, resolve-by-URL, metadata, list drives |
| `IDriveService` | My/user drives, children enumeration, item resolution |
| `IUserDirectoryService` | People picker for the User OneDrive source |
| `ISharingLinkService` | `createLink` and `invite`, returning typed outcomes |
| `IManifestStorageService` | Read/upload/update manifests with ETag concurrency |
| `IAppRegistrationService` | Create and inspect the tenant-specific registration |
| `IConsentService` | Build official consent URLs; run and verify consent |
| `IFileDiscoveryService` | Streamed, cancellable, filtered enumeration |
| `ILinkJobRunner` | Preflight, execute, retry, cancel, summarize |
| `IManifestFormatter` / `IManifestParser` | Per-format read and write |
| `ISettingsStore`, `IJobHistoryStore`, `IProfileStore`, `IAuditStore` | Local non-secret persistence |
| `IDiagnosticsService` | Sanitized bundle export |
| `ISystemBrowser`, `IClipboard`, `IFileDialogs` | Platform affordances, mockable |

## 3. Data flow of a job

```
Targets (user selection)
  -> TargetPlanner        deduplicate, resolve overlap, order
  -> Preflight            auth, consent, scopes, resolvability, write access, conflicts
  -> Discovery            paged enumeration -> IAsyncEnumerable<DiscoveredFile>
  -> Filters              extension/glob/date/size/system/manifest-name rules
  -> Preview              counts, warnings, unknowns  (dry run stops here)
  -> Execution            createLink / invite per file, bounded concurrency
  -> ManifestBuilder      per-folder and/or master, per format
  -> ManifestStorage      read + ETag + conflict policy + upload (small or session)
  -> JobSummary           counts, manifest locations, sanitized errors
```

Cancellation is cooperative at every arrow. Results already produced are preserved, and the
user may still write manifests from a cancelled run's successes.

## 4. Concurrency model

- Discovery is a single ordered walk per target (`IAsyncEnumerable`), so memory stays flat on
  large libraries.
- Link creation runs through a bounded `SemaphoreSlim` (`MaxConcurrency`, default 4) with an
  optional inter-request delay. Graph throttling is a shared-tenant concern; the default is
  deliberately modest.
- The retry policy is centralized and pure: `(attempt, statusCode, Retry-After) -> delay|stop`,
  with injectable jitter so tests are deterministic.
- All UI updates marshal to the UI thread through Avalonia's dispatcher; view models never
  block on `.Result` or `.Wait()`.

## 5. Immutability of a running job

`JobConfiguration` is a `record` with init-only members and immutable collections. Once
`ILinkJobRunner.RunAsync` is entered, the configuration cannot change underneath the run —
editing in the UI produces a new configuration for the next run. This removes an entire class
of "the user changed the audience halfway through" bug.

## 6. Error strategy

Graph failures are normalized into `GraphError` at the transport boundary: sanitized HTTP
status, Graph error code, a human-readable explanation, a `GraphErrorKind` enum, and a
retry-eligibility flag. View models render `GraphErrorKind`, never a raw exception string.
Per-file failures carry full context (tenant, source, site, drive, path, requested operation)
so a retry can be reconstructed exactly.

Exceptions are never silently swallowed. The few broad `catch` sites — the UI top level and
the per-file execution loop — log with structured context and convert to a typed result.

## 7. Logging and redaction

Structured logging via `Microsoft.Extensions.Logging`. A `RedactingLogger` decorator runs
every message and scope value through `SensitiveDataRedactor`, which removes bearer tokens,
authorization headers, `code=`/`id_token`/`access_token` query values, and sharing URLs.
Tokens and authorization codes are never logged, at any level, including `Trace`.

## 8. Settings and local state

| Data | Location | Contains secrets? |
|---|---|---|
| Application settings | OS app-data dir / `settings.json` | No |
| Tenant configuration | OS app-data dir / `tenant.json` | No — client ID and tenant ID only |
| Saved profiles | OS app-data dir / `profiles/` | No |
| Job history | OS app-data dir / `history/` | No |
| Registration audit | OS app-data dir / `audit/` | No |
| MSAL token cache | OS secure storage (Keychain / DPAPI / Secret Service) | Yes — never in the files above |

## 9. Testing seams

- `IGraphApiClient` is exercised through a fake `HttpMessageHandler`, so pagination,
  throttling, ETag conflicts, `207`, and cancellation are tested against the real client.
- `Core` logic is pure and needs no fakes at all.
- `TimeProvider` and an injectable jitter function make retry/backoff deterministic.
- The Avalonia headless harness asserts that the application actually starts and builds its
  main window.
