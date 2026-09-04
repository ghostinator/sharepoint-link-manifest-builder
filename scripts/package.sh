#!/usr/bin/env bash
# Publishes every supported runtime identifier and produces archives.
#
#   ./scripts/package.sh              publishes all RIDs
#   ./scripts/package.sh osx-arm64    publishes one
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

RIDS=("$@")
[[ ${#RIDS[@]} -gt 0 ]] || RIDS=("${ALL_RIDS[@]}")

mkdir -p "${ARTIFACTS}/packages"

for rid in "${RIDS[@]}"; do
  "$(dirname "${BASH_SOURCE[0]}")/publish.sh" "${rid}"

  version="$(grep -o '<VersionPrefix>[^<]*' "${REPO_ROOT}/Directory.Build.props" | head -1 | cut -d'>' -f2)"
  archive="${ARTIFACTS}/packages/SharePointLinkManifestBuilder-${version}-${rid}"

  say "Archiving ${rid}"

  if [[ "${rid}" == win-* ]]; then
    # Windows users expect a zip, but the tool that makes one varies by host. A GitHub
    # Windows runner has no `zip` on the Git Bash PATH, which is how the first release
    # failed: publishing succeeded on all six RIDs and only the archiving step died, with
    # "zip: command not found". Try each of the three tools that might be present.
    if command -v zip >/dev/null 2>&1; then
      (cd "${ARTIFACTS}/publish" && zip -qr "${archive}.zip" "${rid}")
    elif command -v 7z >/dev/null 2>&1; then
      (cd "${ARTIFACTS}/publish" && 7z a -tzip -bso0 -bsp0 "${archive}.zip" "${rid}" >/dev/null)
    elif command -v powershell.exe >/dev/null 2>&1; then
      powershell.exe -NoProfile -NonInteractive -Command \
        "Compress-Archive -Path '${ARTIFACTS}/publish/${rid}' -DestinationPath '${archive}.zip' -Force"
    else
      fail "No zip tool found (tried zip, 7z, powershell). Cannot archive ${rid}."
    fi
  else
    tar -czf "${archive}.tar.gz" -C "${ARTIFACTS}/publish" "${rid}"
  fi
done

say "Packages written to ${ARTIFACTS}/packages"
ls -1 "${ARTIFACTS}/packages"
