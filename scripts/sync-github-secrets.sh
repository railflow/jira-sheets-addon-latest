#!/usr/bin/env bash
# Reads secrets from Infisical and syncs them to GitHub Actions secrets.
# Run this whenever secrets change in Infisical.
#
# Required env vars:
#   INFISICAL_TOKEN  - Infisical service token
#   GH_TOKEN         - GitHub token with secrets:write
#
# Optional:
#   INFISICAL_PROJECT_ID  - Infisical project ID
#   INFISICAL_HOST        - Infisical host (default: https://app.infisical.com)
#   GH_REPO               - GitHub repo (default: railflow/jira-sheets-addon-latest)

set -euo pipefail

INFISICAL_HOST="${INFISICAL_HOST:-https://app.infisical.com}"
INFISICAL_PROJECT_ID="${INFISICAL_PROJECT_ID:?INFISICAL_PROJECT_ID is required}"
GH_REPO="${GH_REPO:-railflow/jira-sheets-addon-latest}"

if [[ -z "${INFISICAL_TOKEN:-}" ]]; then
  echo "Error: INFISICAL_TOKEN is not set" >&2
  exit 1
fi

# Fetch a secret value from Infisical
get_secret() {
  local name="$1"
  curl -fsSL \
    -H "Authorization: Bearer $INFISICAL_TOKEN" \
    "$INFISICAL_HOST/api/v3/secrets/raw/$name?workspaceId=$INFISICAL_PROJECT_ID&environment=production&secretPath=/" \
    2>/dev/null | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['secret']['secretValue'], end='')" 2>/dev/null || true
}

# Sync a secret only if non-empty; otherwise warn and skip
sync_secret() {
  local name="$1"
  local value="$2"
  if [[ -z "$value" ]]; then
    echo "  $name → ⚠️  not found in Infisical, skipping"
    return
  fi
  printf '%s' "$value" | gh secret set "$name" --repo "$GH_REPO"
  echo "  $name → ✨ set"
}

echo "Fetching and syncing secrets from Infisical → GitHub ($GH_REPO)..."
echo ""

sync_secret "PROXY_SECRET"          "$(get_secret "PROXY_SECRET")"
sync_secret "CLOUDFLARE_API_TOKEN"  "$(get_secret "CLOUDFLARE_API_TOKEN")"
sync_secret "CLASP_CREDENTIALS"     "$(get_secret "CLASP_CREDENTIALS")"

echo ""
echo "✅ All GitHub secrets synced from Infisical."
