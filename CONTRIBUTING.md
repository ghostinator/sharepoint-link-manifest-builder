# Contributing

Thank you for considering a contribution.

## Before you start

This application creates real sharing links on real content in real Microsoft 365 tenants. A
defect here can over-share an organization's documents. That shapes how changes are reviewed:
correctness and honesty about outcomes matter more than features.

## Getting set up

```bash
git clone <repository-url>
cd sharepoint-link-manifest-builder
./scripts/restore.sh
./scripts/build.sh
./scripts/test.sh
```

You need the .NET SDK 10.0 or newer. You do **not** need a Microsoft 365 tenant to build or test:
every test runs against a mocked Microsoft Graph.

## Before opening a pull request

```bash
./scripts/format.sh --check
./scripts/build.sh
./scripts/test.sh
./scripts/scan-secrets.sh
```

All four must pass. CI runs the same checks on Windows, macOS and Linux.

## Things that will block a change

**Never commit** a token, secret, certificate, private key, real tenant ID, real client ID, real
SharePoint or OneDrive URL, sharing link, email address, user principal name, generated manifest,
diagnostic bundle, or absolute home-directory path. The publication gate scans the working tree
and full history and fails closed.

**Never hard-code** anything tenant-specific. This application must work unchanged in an
unrelated organization.

**Never add a client secret.** The desktop application is a public client. If a change appears to
need a secret, the design is wrong.

**Never widen a permission for convenience.** Adding a Graph scope requires:

- the least-privileged scope that actually works,
- an entry in `GraphScopes` with a purpose and a plain-language data-access impact,
- an update to `docs/ENTRA-SETUP.md` and `docs/GRAPH-OPERATIONS.md`,
- and the narrower alternative named, if one exists.

**Never report an outcome more confidently than it was verified.** A reused link is not a created
link. A redirect that looked successful is not proof of consent. A preflight check that could not
run is not a pass. This principle runs through the whole codebase and is not negotiable.

**Never modify a tenant silently.** Every change must be shown to the user beforehand and written
to the local audit history afterwards.

## Code conventions

- Nullable reference types are on. Analyzers are on. Warnings are errors in CI.
- `.editorconfig` governs formatting; `dotnet format` settles disputes.
- Public types and interface members carry XML documentation.
- Comments explain **why**, not what. A comment restating the code is noise; a comment
  explaining a non-obvious constraint is valuable.
- Suppress an analyzer at the narrowest scope, with a justification stating the reason.
- Async all the way down, with `CancellationToken` forwarded.
- View models never reference `HttpClient`, MSAL, or a Graph DTO.
- Expected failures return `OperationResult<T>`; exceptions are for the unexpected.
- Never swallow an exception silently.

## Testing

- Logic belongs in `Core` where it can be tested without a network.
- Graph behaviour is tested by scripting an `HttpMessageHandler`, so the real client runs.
  Faking `IGraphApiClient` tests the mock, not the code.
- Anything time-dependent takes a `TimeProvider`; anything random takes an injected source.
- A test name should state the behaviour: `NonRecursive_ProcessesOnlyDirectChildren`.
- A test asserting a security property should say why in a comment.

## Commits

Conventional commits: `feat:`, `fix:`, `docs:`, `test:`, `build:`, `ci:`, `chore:`, `security:`.

Write a body explaining *why*. If a change was prompted by a defect, describe the defect: that
is the part a future reader cannot reconstruct from the diff.

Do not force-push a shared branch. Prefer `git revert` for rolling back shared history.

## Live-tenant testing

Optional and never required by CI. If you test against a tenant:

- Use a dedicated test tenant with synthetic content, never production.
- Start with a dry run.
- Never paste real output into an issue or a pull request.
- Say clearly in the pull request what you tested and what you did not.

## Reporting security issues

Do not open a public issue. See [SECURITY.md](SECURITY.md).
