#!/usr/bin/env bash
# Shared settings for the build scripts.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SOLUTION="${REPO_ROOT}/SharePointLinkManifestBuilder.slnx"
APP_PROJECT="${REPO_ROOT}/src/SharePointLinkManifestBuilder.App/SharePointLinkManifestBuilder.App.csproj"
ARTIFACTS="${REPO_ROOT}/artifacts"
CONFIGURATION="${CONFIGURATION:-Release}"

# Runtime identifiers produced by package.sh. ARM64 Windows and Linux are included because
# .NET supports them; each artifact is labelled with the platform it was actually built for.
ALL_RIDS=(win-x64 win-arm64 osx-x64 osx-arm64 linux-x64 linux-arm64)

say() { printf '\n\033[1m==> %s\033[0m\n' "$*"; }
warn() { printf '\033[33mWARNING: %s\033[0m\n' "$*" >&2; }
fail() { printf '\033[31mERROR: %s\033[0m\n' "$*" >&2; exit 1; }
