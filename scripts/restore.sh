#!/usr/bin/env bash
# Restores NuGet packages for the whole solution.
set -euo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/common.sh"

say "Restoring packages"
dotnet restore "${SOLUTION}"
