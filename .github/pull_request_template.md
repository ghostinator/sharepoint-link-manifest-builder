## What this changes

<!-- One or two sentences. What behaviour is different after this change? -->

## Why

<!-- The problem being solved. Link an issue if there is one. -->

## How it was verified

<!-- What you actually ran, not what you intended to run. -->

- [ ] `./scripts/build.sh` passes
- [ ] `./scripts/test.sh` passes
- [ ] `./scripts/format.sh --check` passes
- [ ] Tested against a live Microsoft 365 tenant — describe what, or state that you did not

## Security and privacy

- [ ] No token, secret, certificate or credential is added, logged, or exported
- [ ] No tenant-specific value is hard-coded (tenant ID, client ID, hostname, site, drive, item, user)
- [ ] No new Microsoft Graph permission is requested; if one is, the least-privileged option was
      chosen and `docs/ENTRA-SETUP.md` and the in-app permission review were both updated
- [ ] `./scripts/scan-secrets.sh` passes
- [ ] Any new tenant-modifying operation is previewed to the user before it runs and written to
      the local audit history

## Behaviour that users see

- [ ] No outcome is reported more confidently than it was verified (for example, a reused
      sharing link is not reported as created)
- [ ] New UI is keyboard reachable and has an accessible name
- [ ] Status is not conveyed by colour alone

## Documentation

- [ ] `docs/GRAPH-OPERATIONS.md` updated if a Graph call was added, changed or removed
- [ ] `CHANGELOG.md` updated
- [ ] An ADR added under `docs/adr/` if this changes an architectural decision
