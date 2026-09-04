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

# A .NET build writes thousands of small files. A cloud-sync client that wants to upload each
# one will intermittently hold a freshly written assembly, and MSBuild's copy into bin/ then
# fails with "Access to the path ... is denied" -- MSB3026 ten times, then MSB3027. Nothing is
# wrong with the code, and the same build succeeds moments later, which is what makes it look
# random. Building outside the synced tree removes the interference entirely.
#
# Set SPLMB_ARTIFACTS_PATH to choose the location. Otherwise a repository sitting inside a known
# sync root gets a default outside it. A repository anywhere else is left exactly as it was, so
# CI and ordinary clones keep the standard bin/obj layout.
BUILD_OUTPUT_ARGS=()
_cloud_sync_root=""
case "${REPO_ROOT}" in
  */Library/CloudStorage/*) _cloud_sync_root="iCloud, OneDrive or another File Provider" ;;
  */Dropbox/*)              _cloud_sync_root="Dropbox" ;;
  */Google?Drive/*)         _cloud_sync_root="Google Drive" ;;
  */OneDrive*/*)            _cloud_sync_root="OneDrive" ;;
esac

if [[ -n "${SPLMB_ARTIFACTS_PATH:-}" ]]; then
  BUILD_OUTPUT_ARGS=(-p:ArtifactsPath="${SPLMB_ARTIFACTS_PATH}")
elif [[ -n "${_cloud_sync_root}" ]]; then
  SPLMB_ARTIFACTS_PATH="${HOME}/.cache/splmb-build/$(basename "${REPO_ROOT}")"
  BUILD_OUTPUT_ARGS=(-p:ArtifactsPath="${SPLMB_ARTIFACTS_PATH}")
  warn "This repository is inside ${_cloud_sync_root}, where sync interference makes builds fail intermittently."
  warn "Build output is being written to ${SPLMB_ARTIFACTS_PATH} instead. Set SPLMB_ARTIFACTS_PATH to override."
fi
