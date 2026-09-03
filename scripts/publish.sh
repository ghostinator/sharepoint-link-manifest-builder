#!/usr/bin/env bash
# Publishes a self-contained, single-file build for one runtime identifier.
#
#   ./scripts/publish.sh osx-arm64
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

RID="${1:-}"
[[ -n "${RID}" ]] || fail "Usage: publish.sh <runtime-identifier>   (one of: ${ALL_RIDS[*]})"

OUTPUT="${ARTIFACTS}/publish/${RID}"

say "Publishing ${RID} to ${OUTPUT}"
rm -rf "${OUTPUT}"

dotnet publish "${APP_PROJECT}" \
  --configuration "${CONFIGURATION}" \
  --runtime "${RID}" \
  --self-contained true \
  --output "${OUTPUT}" \
  -p:PublishSingleFile=true \
  -p:IncludeNativeLibrariesForSelfExtract=true \
  -p:EnableCompressionInSingleFile=true \
  -p:DebugType=none

# Stated plainly rather than implied: nothing here is signed.
cat > "${OUTPUT}/UNSIGNED-ARTIFACT.txt" <<'NOTICE'
This build is NOT code signed and NOT notarized.

Windows SmartScreen and macOS Gatekeeper will warn about it, and on macOS the application
may refuse to open until it is explicitly allowed.

Signing requires credentials that are not present in this repository. See docs/PACKAGING.md
for what a publisher must supply. Do not describe an artifact produced by this script as
signed, because it is not.
NOTICE

say "Published ${RID}"
