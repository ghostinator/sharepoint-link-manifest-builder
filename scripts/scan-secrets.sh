#!/usr/bin/env bash
#
# Publication safety gate.
#
# Scans the working tree AND the full git history for credentials and tenant-identifying data
# before anything is pushed. It fails closed: a non-zero exit means DO NOT PUSH.
#
# History is scanned separately because .gitignore does nothing for content that was already
# committed. Deleting a file in a later commit does not remove it from the repository.
#
#   ./scripts/scan-secrets.sh              scan working tree and history
#   ./scripts/scan-secrets.sh --tree-only  scan the working tree only (faster, for pre-commit)
#
set -uo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${REPO_ROOT}"

REPORT="${REPO_ROOT}/artifacts/publication-readiness-report.txt"
mkdir -p "$(dirname "${REPORT}")"

TREE_ONLY=0
[[ "${1:-}" == "--tree-only" ]] && TREE_ONLY=1

FINDINGS=0
: > "${REPORT}"

red()   { printf '\033[31m%s\033[0m\n' "$*"; }
green() { printf '\033[32m%s\033[0m\n' "$*"; }
bold()  { printf '\033[1m%s\033[0m\n' "$*"; }

log() { echo "$*" | tee -a "${REPORT}"; }

# ---------------------------------------------------------------------------
# Synthetic values that are expected to appear in documentation and tests.
# Anything matching these is not a finding.
#
# Two entries deserve explanation, because a careless allowlist is how a scanner stops
# catching anything:
#
#   eyJzdWIiOiJ0ZXN0In0
#       The base64url payload {"sub":"test"}. This exempts ONLY a JWT whose payload is
#       literally that, which is the synthetic constant used to test the redactor. A real
#       token would never carry it. Earlier commits predate the inline SCAN-ALLOW marker, so
#       the history scan needs this rather than a history rewrite over a known non-secret.
#
#   yourdomain / yourcompany / your-tenant
#       Illustrative placeholders in the documentation.
#
#   github@ghostinator\.co
#       This project's published contact address, in SECURITY.md, CODE_OF_CONDUCT.md,
#       PRIVACY.md and the product metadata. It is deliberately public, so the scanner
#       flagging it is the rule working rather than failing. Exempted as a literal string
#       rather than by weakening the email rule, so a real tenant UPN is still caught.
# ---------------------------------------------------------------------------
ALLOWLIST='example\.sharepoint|example-my\.sharepoint|example\.(com|org|net|test|invalid)|contoso|fabrikam|adventure-works|PLACEHOLDER|your-tenant|yourcompany|yourdomain|tenant-name|localhost|00000000-0000-0000-0000-000000000000|11111111-1111|22222222-2222|33333333-3333|44444444-4444|FAKE-TEST-TOKEN|eyJzdWIiOiJ0ZXN0In0|schemas\.microsoft\.com|microsoftonline\.com|graph\.microsoft\.com|sharepoint\.com/dev|learn\.microsoft\.com|entra\.microsoft\.com|github\.com/cli|github@ghostinator\.co'

# Paths whose content is expected to describe the patterns themselves.
SELF_REFERENTIAL='^\.gitignore$|^scripts/scan-secrets\.sh$|^\.github/workflows/security\.yml$|^docs/THREAT-MODEL\.md$|^docs/GITHUB-PUBLISHING\.md$'

# ---------------------------------------------------------------------------
# Detection rules: NAME | DESCRIPTION | EXTENDED REGEX
# ---------------------------------------------------------------------------
RULES=(
  "jwt|JSON Web Token (possible access or ID token)|eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\."
  "bearer|Bearer token in an Authorization header|[Aa]uthorization:[[:space:]]*[Bb]earer[[:space:]]+[A-Za-z0-9._~+/-]{20,}"
  "github-token|GitHub personal access or app token|gh[pousr]_[A-Za-z0-9]{30,}"
  "github-pat|GitHub fine-grained personal access token|github_pat_[A-Za-z0-9_]{50,}"
  "aws-key|AWS access key identifier|AKIA[0-9A-Z]{16}"
  "slack-token|Slack token|xox[baprs]-[A-Za-z0-9-]{10,}"
  "private-key|Private key block (RSA, EC, OpenSSH, PGP)|-----BEGIN[[:space:]]+([A-Z]+[[:space:]]+)?PRIVATE[[:space:]]+KEY"
  "certificate|Certificate block|-----BEGIN[[:space:]]+CERTIFICATE-----"
  "client-secret|Client secret assignment|client_?secret[\"'\''[:space:]]*[:=][\"'\''[:space:]]*[A-Za-z0-9~._-]{8,}"
  "password|Password assignment|password[\"'\''[:space:]]*[:=][\"'\''[:space:]]*[^\"'\''[:space:]]{6,}"
  "connection-string|Connection string with credentials|(AccountKey|Password|Pwd)=[^;\"[:space:]]{8,}"
  "auth-code|OAuth authorization code in a URL|[?&]code=[A-Za-z0-9._-]{20,}"
  "refresh-token|Refresh token in a URL or assignment|refresh_token[\"'\''[:space:]]*[:=][\"'\''[:space:]]*[A-Za-z0-9._-]{20,}"
  "sharing-link|SharePoint or OneDrive sharing link|https://[A-Za-z0-9-]+(-my)?\.sharepoint\.com/:[a-z]:/"
  "tenant-url|Tenant-specific SharePoint hostname|https://[A-Za-z0-9-]+(-my)?\.sharepoint\.(com|us|cn|de)"
  "onedrive-personal|Personal OneDrive path|/personal/[A-Za-z0-9._%-]+_[A-Za-z0-9._%-]+_[A-Za-z]{2,}/"
  "email|Email address or user principal name|[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
  "home-path-unix|Absolute home-directory path|/(Users|home)/[A-Za-z0-9._-]+/"
  "home-path-windows|Absolute Windows user path|[A-Za-z]:\\\\Users\\\\[A-Za-z0-9._-]+"
  "internal-host|Internal hostname|https?://[A-Za-z0-9-]+\.(local|internal|corp|lan|intranet)(/|:|$)"
)

# Files that must never be tracked at all, regardless of content.
FORBIDDEN_PATHS=(
  '_sharepoint-links.*\.(txt|md|csv|json)$'
  '\.pfx$' '\.p12$' '\.pem$' '\.key$' '\.snk$' '\.jks$'
  'msal\.cache' 'tokens?\.json$'
  'appsettings\.(Local|Development|Production)\.json$'
  'tenant\.json$' 'job-history\.json$'
  '\.diagbundle' 'diagnostic.*\.zip$'
  'id_rsa' 'id_ed25519'
)

report_finding() {
  local rule_name="$1" description="$2" location="$3" evidence="$4"
  FINDINGS=$((FINDINGS + 1))

  {
    echo "FINDING #${FINDINGS}"
    echo "  Category : ${rule_name}"
    echo "  Concern  : ${description}"
    echo "  Location : ${location}"
    # Only ever a short redacted excerpt: printing a whole secret into a report file that
    # itself may be shared would defeat the purpose of finding it.
    echo "  Evidence : ${evidence}"
    echo
  } | tee -a "${REPORT}"
}

# Truncates and masks the middle of a matched line.
redact() {
  local line="$1"
  line="${line:0:110}"
  echo "${line}" | sed -E 's/([A-Za-z0-9._~+/-]{6})[A-Za-z0-9._~+/-]{6,}/\1********/g'
}

bold "Publication safety scan"
log "SharePoint Link Manifest Builder - publication readiness report"
log "Generated: $(date -u +%Y-%m-%dT%H:%M:%SZ)"
log "Repository: ${REPO_ROOT}"
log ""

# ---------------------------------------------------------------------------
# 1. Forbidden tracked paths
# ---------------------------------------------------------------------------
bold "1. Checking for files that must never be committed"
log "== Forbidden tracked paths =="

TRACKED="$(git ls-files 2>/dev/null || true)"

for pattern in "${FORBIDDEN_PATHS[@]}"; do
  while IFS= read -r path; do
    [[ -z "${path}" ]] && continue
    # docs/samples holds deliberately synthetic example manifests.
    [[ "${path}" == docs/samples/* ]] && continue
    report_finding "forbidden-path" "A file of this kind must never be tracked" "${path}" "matched ${pattern}"
  done < <(echo "${TRACKED}" | grep -E "${pattern}" || true)
done

[[ ${FINDINGS} -eq 0 ]] && log "  none" && green "  OK"
log ""

# ---------------------------------------------------------------------------
# 2. Working tree content
# ---------------------------------------------------------------------------
bold "2. Scanning working tree content"
log "== Working tree =="

TREE_FINDINGS_BEFORE=${FINDINGS}

for rule in "${RULES[@]}"; do
  IFS='|' read -r name description regex <<< "${rule}"

  while IFS= read -r match; do
    [[ -z "${match}" ]] && continue

    file="${match%%:*}"
    rest="${match#*:}"
    line_no="${rest%%:*}"
    content="${rest#*:}"

    [[ "${file}" =~ ${SELF_REFERENTIAL} ]] && continue
    echo "${content}" | grep -Eq "${ALLOWLIST}" && continue

    # An explicit, greppable marker for a value that is deliberately synthetic. Every use is
    # reviewable with: git grep SCAN-ALLOW
    echo "${content}" | grep -q 'SCAN-ALLOW' && continue

    report_finding "${name}" "${description}" "${file}:${line_no}" "$(redact "${content}")"
  done < <(git grep -I -n -E "${regex}" -- . ':(exclude)artifacts' 2>/dev/null || true)
done

if [[ ${FINDINGS} -eq ${TREE_FINDINGS_BEFORE} ]]; then
  log "  none"
  green "  OK"
fi
log ""

# ---------------------------------------------------------------------------
# 3. Git history
# ---------------------------------------------------------------------------
if [[ ${TREE_ONLY} -eq 0 ]]; then
  bold "3. Scanning git history"
  log "== Git history =="
  log "(.gitignore does not remove content that was already committed, so history is scanned separately.)"

  HISTORY_FINDINGS_BEFORE=${FINDINGS}
  COMMITS="$(git rev-list --all 2>/dev/null || true)"

  if [[ -z "${COMMITS}" ]]; then
    log "  no commits yet"
  else
    for rule in "${RULES[@]}"; do
      IFS='|' read -r name description regex <<< "${rule}"

      # git grep across every commit searches tree contents only, so commit metadata such as
      # the committer's own email address is not reported as a finding in file content.
      while IFS= read -r match; do
        [[ -z "${match}" ]] && continue

        commit="${match%%:*}"
        rest="${match#*:}"
        file="${rest%%:*}"
        rest2="${rest#*:}"
        line_no="${rest2%%:*}"
        content="${rest2#*:}"

        [[ "${file}" =~ ${SELF_REFERENTIAL} ]] && continue
        echo "${content}" | grep -Eq "${ALLOWLIST}" && continue
        echo "${content}" | grep -q 'SCAN-ALLOW' && continue

        report_finding "${name} (history)" "${description}" \
          "commit ${commit:0:8} ${file}:${line_no}" "$(redact "${content}")"
      done < <(git grep -I -n -E "${regex}" $(echo "${COMMITS}") 2>/dev/null | head -50 || true)
    done
  fi

  if [[ ${FINDINGS} -eq ${HISTORY_FINDINGS_BEFORE} ]]; then
    log "  none"
    green "  OK"
  fi
  log ""
else
  log "== Git history =="
  log "  SKIPPED (--tree-only). Run the full scan before pushing."
  log ""
fi

# ---------------------------------------------------------------------------
# 4. Placeholder inventory (informational, never a failure)
# ---------------------------------------------------------------------------
bold "4. Placeholder inventory"
log "== Placeholders still to be replaced by a publisher =="

PLACEHOLDER_COUNT="$(git grep -I -c 'PLACEHOLDER' -- . 2>/dev/null | wc -l | tr -d ' ')"
log "  Files containing PLACEHOLDER: ${PLACEHOLDER_COUNT}"
log "  These are expected. A publisher replaces them before distribution."
log ""

# ---------------------------------------------------------------------------
# Verdict
# ---------------------------------------------------------------------------
log "== Result =="

if [[ ${FINDINGS} -gt 0 ]]; then
  log "FAILED: ${FINDINGS} finding(s). DO NOT PUSH."
  log ""
  log "Before pushing:"
  log "  1. Remove the offending content from the working tree."
  log "  2. If it was already committed, rewrite history (git filter-repo) - deleting the file"
  log "     in a new commit does NOT remove it from the repository."
  log "  3. Treat any exposed credential as compromised. Revoke and rotate it; removing it from"
  log "     git does not un-disclose it."
  log "  4. Re-run this scan until it passes."

  red ""
  red "PUBLICATION SAFETY SCAN FAILED: ${FINDINGS} finding(s)."
  red "Report: ${REPORT}"
  exit 1
fi

log "PASSED: no secrets or tenant-identifying data detected."
log ""
log "This scan is a safety net, not a guarantee. A human should still review the diff before"
log "making a repository public."

green ""
green "PUBLICATION SAFETY SCAN PASSED"
echo "Report: ${REPORT}"
exit 0
