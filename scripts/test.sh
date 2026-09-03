#!/usr/bin/env bash
# Runs every test. No live Microsoft 365 tenant and no credential is required.
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

mkdir -p "${ARTIFACTS}/test-results"

say "Running tests (${CONFIGURATION})"
dotnet test "${SOLUTION}" \
  --configuration "${CONFIGURATION}" \
  --logger "trx;LogFileName=results.trx" \
  --results-directory "${ARTIFACTS}/test-results" \
  --collect:"XPlat Code Coverage" \
  "$@"
