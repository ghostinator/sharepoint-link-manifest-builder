#!/usr/bin/env bash
# Formats the solution, or verifies formatting when --check is passed.
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

if [[ "${1:-}" == "--check" ]]; then
  say "Verifying formatting (no changes will be made)"
  dotnet format "${SOLUTION}" --verify-no-changes --severity warn
else
  say "Formatting"
  dotnet format "${SOLUTION}" --severity warn
fi
