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
    (cd "${ARTIFACTS}/publish" && zip -qr "${archive}.zip" "${rid}")
  else
    tar -czf "${archive}.tar.gz" -C "${ARTIFACTS}/publish" "${rid}"
  fi
done

say "Packages written to ${ARTIFACTS}/packages"
ls -1 "${ARTIFACTS}/packages"
