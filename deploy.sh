#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PROFILE="$ROOT/ConvertLearnToDocWeb.PublishSettings"

for tool in dotnet zip curl xmllint jq; do
  command -v "$tool" >/dev/null || { echo "Required tool missing: $tool" >&2; exit 1; }
done

DEPLOY_USER="$(xmllint --xpath 'string(//publishProfile[@publishMethod="MSDeploy"]/@userName)' "$PROFILE")"
DEPLOY_PASS="$(xmllint --xpath 'string(//publishProfile[@publishMethod="MSDeploy"]/@userPWD)' "$PROFILE")"
DEPLOY_HOST="$(xmllint --xpath 'string(//publishProfile[@publishMethod="MSDeploy"]/@publishUrl)' "$PROFILE")"
if [[ -z "$DEPLOY_USER" || -z "$DEPLOY_PASS" || -z "$DEPLOY_HOST" ]]; then
  echo "Publish profile must contain MSDeploy credentials and a publishUrl." >&2
  exit 1
fi
SCM="https://${DEPLOY_HOST%:443}"

# A fresh directory prevents deleted files from surviving in the next package.
WORK="$(mktemp -d)"
trap 'rm -rf -- "$WORK"' EXIT

dotnet publish "$ROOT/src/ConvertLearnToDoc/ConvertLearnToDoc.csproj" -c Release -o "$WORK/publish"
( cd "$WORK/publish" && zip -rq "$WORK/app.zip" . )

echo "Publishing to $SCM"
HTTP_STATUS="$(curl -fsS --connect-timeout 30 --max-time 600 -u "$DEPLOY_USER:$DEPLOY_PASS" \
  -D "$WORK/headers" -o "$WORK/upload-response" -w '%{http_code}' \
  -H "Content-Type: application/zip" --data-binary "@$WORK/app.zip" \
  "$SCM/api/zipdeploy?isAsync=true")"
if [[ "$HTTP_STATUS" != 202 ]]; then
  echo "Expected an accepted deployment (HTTP 202); received HTTP $HTTP_STATUS." >&2
  exit 1
fi

# Use Azure's returned tracking URL, including its upload timestamp query.
STATUS_URL="$(awk 'tolower($1) == "location:" {sub(/\r$/, "", $2); print $2}' "$WORK/headers")"
case "$STATUS_URL" in
  "$SCM/api/deployments/"*|"$SCM:443/api/deployments/"*) ;;
  /api/deployments/*) STATUS_URL="$SCM$STATUS_URL" ;;
  *) echo "Unexpected deployment status URL; check $SCM/api/deployments." >&2; exit 1 ;;
esac

echo "Upload accepted. Waiting for Azure deployment: $STATUS_URL"
DEADLINE=$((SECONDS + 600))
while (( SECONDS < DEADLINE )); do
  RESULT="$(curl -fsS --connect-timeout 15 --max-time 30 -u "$DEPLOY_USER:$DEPLOY_PASS" "$STATUS_URL")"
  DEPLOYMENT_ID="$(jq -er '.id | strings | select(length > 0) | @uri' <<< "$RESULT")"
  STATUS_URL="$SCM/api/deployments/$DEPLOYMENT_ID"
  STATUS="$(jq -er '.status | numbers' <<< "$RESULT")"
  case "$STATUS" in
    4) echo "Deployment succeeded: $STATUS_URL"; exit 0 ;;
    3)
      echo "Deployment failed. Logs: $STATUS_URL/log" >&2
      curl -fsS --connect-timeout 15 --max-time 30 -u "$DEPLOY_USER:$DEPLOY_PASS" "$STATUS_URL/log" \
        | jq -r '.[] | .message, (.details_url // empty)' >&2 || true
      exit 1
      ;;
    0|1|2) sleep 5 ;;
    *) echo "Unexpected Azure deployment status: $STATUS" >&2; exit 1 ;;
  esac
done
echo "Timed out waiting for deployment. Check $STATUS_URL before retrying." >&2
exit 1
