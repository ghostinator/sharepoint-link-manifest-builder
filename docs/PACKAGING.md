# Packaging and distribution

## Building artifacts

```bash
./scripts/package.sh                    # all six runtime identifiers
./scripts/package.sh osx-arm64          # just one
./scripts/checksums.sh                  # SHA-256 for everything in artifacts/packages
./scripts/sbom.sh                       # dependency inventory and vulnerability report
```

Output lands in `artifacts/packages/`.

## Supported runtime identifiers

| RID | Platform | Notes |
|---|---|---|
| `win-x64` | Windows 10/11, Intel/AMD | |
| `win-arm64` | Windows 11 on ARM | |
| `osx-x64` | macOS, Intel | |
| `osx-arm64` | macOS, Apple Silicon | |
| `linux-x64` | Linux, Intel/AMD | glibc-based distributions |
| `linux-arm64` | Linux, ARM64 | glibc-based distributions |

Builds are **self-contained** and **single-file**: the .NET runtime is included, so no runtime
installation is required.

## Signing status

> **Artifacts produced by this repository are NOT code signed and NOT notarized.**

Every published artifact includes an `UNSIGNED-ARTIFACT.txt` file saying so, and the release
notes repeat it. Do not describe an artifact from these scripts as signed, because it is not.

Consequences for users:

| Platform | What the user sees |
|---|---|
| Windows | SmartScreen warns about an unrecognised publisher |
| macOS | Gatekeeper blocks it; the user must right-click and choose *Open* |
| Linux | Nothing; Linux does not gate on signatures |

### Windows code signing

Requires an Authenticode certificate — an EV certificate gives immediate SmartScreen reputation;
an OV certificate builds reputation over time.

```powershell
signtool sign /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 `
  /f <certificate.pfx> /p <password> `
  artifacts\publish\win-x64\SharePointLinkManifestBuilder.exe
```

Never commit the certificate or its password. In CI, keep them in encrypted secrets, or better,
use a cloud signing service so the private key never reaches the runner.

### macOS signing and notarization

Requires an Apple Developer Program membership and a Developer ID Application certificate.

```bash
codesign --force --options runtime --timestamp \
  --sign "Developer ID Application: <NAME> (<TEAMID>)" \
  artifacts/publish/osx-arm64/SharePointLinkManifestBuilder

ditto -c -k --keepParent \
  artifacts/publish/osx-arm64/SharePointLinkManifestBuilder \
  SharePointLinkManifestBuilder.zip

xcrun notarytool submit SharePointLinkManifestBuilder.zip \
  --apple-id <APPLE_ID> --team-id <TEAMID> --password <APP_SPECIFIC_PASSWORD> --wait

xcrun stapler staple artifacts/publish/osx-arm64/SharePointLinkManifestBuilder.app
```

A single-file build may need `--options runtime` plus entitlements for the JIT. If notarization
rejects it, produce a proper `.app` bundle rather than a bare executable.

### Linux packaging

No signing is required. Options:

- **Tarball** — what `package.sh` produces. Extract and run.
- **AppImage** — a single portable file. Use `appimagetool` over the publish directory.
- **.deb / .rpm** — better desktop integration. `dotnet-deb`/`dotnet-rpm` or `fpm`.
- **Flatpak** — sandboxed, but note that the sandbox affects Secret Service access.

## Runtime dependencies

Self-contained builds include the .NET runtime. Additional per-platform requirements:

| Platform | Requirement | If missing |
|---|---|---|
| Windows | None | — |
| macOS | None | — |
| Linux | `libicu` | .NET fails to start unless invariant globalization is enabled |
| Linux | `libsecret` and a Secret Service provider | Secure token storage unavailable; the application says so and uses memory-only tokens |

On Debian and Ubuntu:

```bash
sudo apt-get install libicu-dev libsecret-1-0 gnome-keyring
```

## Installation

**Windows.** Extract the ZIP anywhere and run the executable. No installer is provided; add a
Start Menu shortcut manually if wanted.

**macOS.** Extract the archive. Because it is unsigned, right-click and choose *Open* the first
time, then confirm. Verify the checksum first.

**Linux.** Extract the tarball, `chmod +x` the executable, and run it. Add a `.desktop` entry if
wanted.

## Verifying a download

```bash
sha256sum -c SHA256SUMS.txt      # Linux
shasum -a 256 -c SHA256SUMS.txt  # macOS
```

```powershell
Get-FileHash .\SharePointLinkManifestBuilder-0.1.0-win-x64.zip -Algorithm SHA256
```

Because artifacts are unsigned, **the checksum is the only integrity check available**. Compare
it against the value published with the release.

## Upgrading

1. Verify the checksum of the new download.
2. Close the application.
3. Replace the extracted directory.
4. Start it again.

Settings, saved profiles, job history and sign-in details live in the application-data
directory, not in the program directory, so they survive an upgrade. Job history from an older
version is read by a newer one; see [ROLLBACK.md](ROLLBACK.md) for the reverse.

## Uninstalling

1. Delete the extracted program directory.
2. Optionally delete the application-data directory — the exact path is on the Settings page.
   *Remove tenant configuration* in the application does this cleanly.

**Uninstalling never deletes the Microsoft Entra app registration.** Deleting a registration is
a separate, explicitly confirmed action on the Permissions page, and only for a registration the
application created.

## What a publisher must supply

| Item | Needed for | Status here |
|---|---|---|
| Publisher name and URLs | About page, consent screen | PLACEHOLDER |
| Privacy policy URL | Shown on the consent screen | PLACEHOLDER |
| Bootstrap client ID | Automatic tenant setup | Not supplied |
| Authenticode certificate | Windows signing | Not supplied |
| Apple Developer ID and notarization credentials | macOS signing | Not supplied |
| Licence | Legal use and redistribution | Not chosen — see LICENSE-SELECTION.md |
