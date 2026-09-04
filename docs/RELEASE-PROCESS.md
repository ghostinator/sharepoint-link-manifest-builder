# Release process

## Versioning

[Semantic Versioning](https://semver.org): `MAJOR.MINOR.PATCH`.

- **MAJOR** — a breaking change to the manifest schema, the configuration format, or the
  required permission set.
- **MINOR** — new capability, backward compatible.
- **PATCH** — fixes only.

Pre-1.0, minor versions may still break things; this is called out in the changelog.

The manifest schema has its **own** version (`ManifestDefaults.SchemaVersion`), independent of
the application version. Changing it is a major application change, because an older build must
refuse a newer manifest rather than mis-parse it.

## Before releasing

```bash
./scripts/format.sh --check
./scripts/build.sh
./scripts/test.sh
./scripts/scan-secrets.sh        # full history, not --tree-only
./scripts/sbom.sh
```

Then check:

- [ ] `CHANGELOG.md` has an entry for this version, with Added / Changed / Fixed / Security
- [ ] Known limitations in `README.md` are still accurate
- [ ] `docs/GRAPH-OPERATIONS.md` matches the Graph calls actually made
- [ ] Any new permission is documented in `docs/ENTRA-SETUP.md` **and** explained in the in-app
      permission review
- [ ] Version bumped in `Directory.Build.props`
- [ ] An ADR exists for any architectural change
- [ ] Live-tenant validation status is stated honestly in the release notes

## Cutting the release

```bash
# 1. Bump the version
#    Directory.Build.props: VersionPrefix, AssemblyVersion, FileVersion

# 2. Move Unreleased to the new version in CHANGELOG.md, with a date

git add Directory.Build.props CHANGELOG.md
git commit -m "chore: release 1.2.3"

# 3. Tag. The release workflow triggers on a v-prefixed semantic version tag.
git tag -a v1.2.3 -m "SharePoint Link Manifest Builder 1.2.3"

# 4. Push (requires explicit authorization to push to a real remote)
git push origin main
git push origin v1.2.3
```

## What the workflow does

1. **Verify** — publication safety scan over full history, build with warnings as errors, tests.
2. **Package** — six runtime identifiers, each on its own platform family.
3. **Release** — collects artifacts, generates SHA-256 checksums and a dependency inventory,
   composes notes, and publishes with the GitHub CLI using the workflow token.

The release notes state plainly that artifacts are unsigned. Do not edit that out.

## After releasing

- [ ] Download one artifact and verify its checksum
- [ ] Smoke-test on at least one platform: it starts, the wizard opens, Diagnostics reports
      sensible values
- [ ] Add a new `Unreleased` section to `CHANGELOG.md`
- [ ] Close the milestone

## Signing

Not performed by this repository, which has no signing credentials. If a publisher adds signing,
insert it between Package and Release, and only then may the notes describe artifacts as signed.

**Never claim an artifact is signed unless signing actually completed.** A false claim is worse
than no claim: it invites users to skip the checksum verification that is currently their only
integrity check.

## Hotfixes

```bash
git checkout -b hotfix/1.2.4 v1.2.3
# fix, with a regression test
git commit -m "fix: <what>"
git tag -a v1.2.4 -m "SharePoint Link Manifest Builder 1.2.4"
# then merge back into main
```

## Pulling a release

If a release is found to be harmful:

1. Mark the GitHub release as a pre-release, or delete it, and say why in the notes.
2. Publish an advisory if the problem is a security issue.
3. Ship a fixed version promptly.

Do **not** silently replace the artifacts of a published tag. Someone has already verified the
old checksum; changing the bytes behind it destroys that guarantee.
