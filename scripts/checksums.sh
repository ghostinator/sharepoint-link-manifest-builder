#!/usr/bin/env bash
# Produces SHA-256 checksums for every packaged artifact.
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

PACKAGES="${ARTIFACTS}/packages"
[[ -d "${PACKAGES}" ]] || fail "No packages found. Run scripts/package.sh first."

say "Computing SHA-256 checksums"
cd "${PACKAGES}"

if command -v sha256sum >/dev/null 2>&1; then
  sha256sum ./* > SHA256SUMS.txt
else
  # macOS ships shasum rather than sha256sum.
  shasum -a 256 ./* > SHA256SUMS.txt
fi

cat SHA256SUMS.txt
