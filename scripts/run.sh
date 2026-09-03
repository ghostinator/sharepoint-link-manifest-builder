#!/usr/bin/env bash
# Runs the desktop application from source.
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

say "Starting the application"
dotnet run --project "${APP_PROJECT}" --configuration "${CONFIGURATION}" "${BUILD_OUTPUT_ARGS[@]}" -- "$@"
