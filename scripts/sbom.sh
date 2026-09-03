#!/usr/bin/env bash
# Produces a software bill of materials from the restored package graph.
#
# Uses only the .NET SDK, so it works without installing a third-party SBOM tool. A publisher
# who wants a signed CycloneDX or SPDX document should run their own tool in addition; this
# produces an honest dependency inventory, not a certified SBOM.
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

mkdir -p "${ARTIFACTS}/sbom"
OUTPUT="${ARTIFACTS}/sbom/dependencies.txt"

say "Generating dependency inventory"

{
  echo "SharePoint Link Manifest Builder - dependency inventory"
  echo "Generated: $(date -u +%Y-%m-%dT%H:%M:%SZ)"
  echo
  echo "== Direct and transitive packages =="
  dotnet list "${SOLUTION}" package --include-transitive
  echo
  echo "== Known vulnerabilities =="
  dotnet list "${SOLUTION}" package --vulnerable --include-transitive
  echo
  echo "== Deprecated packages =="
  dotnet list "${SOLUTION}" package --deprecated
} > "${OUTPUT}"

say "Written to ${OUTPUT}"
