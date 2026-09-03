#!/usr/bin/env bash
# Builds the solution. Warnings are errors so nothing rots quietly.
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

say "Building (${CONFIGURATION})"
dotnet build "${SOLUTION}" --configuration "${CONFIGURATION}" -warnaserror "${BUILD_OUTPUT_ARGS[@]}"
