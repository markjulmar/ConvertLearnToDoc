set -euo pipefail

PROFILE="./ConvertLearnToDocWeb.PublishSettings"   # path to your publish profile
APP="ConvertLearnToDocWeb"                       # e.g. my-webapp

dotnet publish ./src/ConvertLearnToDoc/ConvertLearnToDoc.csproj -c Release -o ./publish

( cd publish && zip -rq ../app.zip . )

USER="$(/usr/bin/xmllint --xpath 'string(//publishProfile[@publishMethod="MSDeploy"]/@userName)' "$PROFILE")"
PASS="$(/usr/bin/xmllint --xpath 'string(//publishProfile[@publishMethod="MSDeploy"]/@userPWD)' "$PROFILE")"
SCM="https://${APP}.scm.azurewebsites.net"

echo "Publishing to $SCM"

curl -fsS -u "$USER:$PASS" \
  -H "Content-Type: application/zip" \
  --data-binary "@app.zip" \
  "$SCM/api/zipdeploy"