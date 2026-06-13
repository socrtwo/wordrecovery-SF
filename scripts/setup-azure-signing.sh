#!/usr/bin/env bash
# setup-azure-signing.sh
#
# Adds THIS repo to the shared Azure Trusted Signing identity used across
# socrtwo's projects. It reuses the App Registration originally created for
# the-best-of-brock-cookbook — so there is NO new app, certificate, or role
# assignment to make here. All this script does is:
#   1. Locate the shared App Registration.
#   2. Add two federated credentials so THIS repo can obtain OIDC tokens:
#        - environment:release  -> used by the Windows signing job
#        - ref:refs/heads/main   -> used by the verify-oidc.yml smoke test
#   3. Print the six GitHub repo Variables to set (identical across all repos).
#
# Run this ONCE in Azure Cloud Shell (shell.azure.com) from an account that
# owns the subscription below. Safe to re-run — existing creds are left alone.
#
# Prerequisite: the-best-of-brock-cookbook/scripts/setup-azure-signing.sh has
# already been run once (that is what creates the shared App + cert role).

set -euo pipefail

# --------------------------------------------------------------------------
# Shared org-wide values (identical for every repo using this signer).
# --------------------------------------------------------------------------
readonly SUBSCRIPTION_ID="5b2f1523-5925-4630-a326-22610b97db02"
readonly TENANT_ID="99dd7dad-5263-44fc-aea5-63c220ebcd48"
readonly SIGNING_ACCOUNT="S2SetvicesCodeSigning"
readonly CERT_PROFILE="GitHubAppSigner752"
readonly TS_ENDPOINT="https://eus.codesigning.azure.net/"

readonly GITHUB_OWNER="socrtwo"
readonly GITHUB_REPO="wordrecovery-SF"

# Shared signing identity (created by the-best-of-brock-cookbook setup).
readonly APP_NAME="brock-cookbook-signer"

# --------------------------------------------------------------------------
say()  { printf '\033[36m==>\033[0m %s\n' "$*"; }
die()  { printf '\033[31m!!\033[0m %s\n' "$*" >&2; exit 1; }

command -v az >/dev/null || die "az CLI not found. Use Azure Cloud Shell."

say "Setting active subscription"
az account set --subscription "$SUBSCRIPTION_ID"

say "Locating shared App Registration '$APP_NAME'"
APP_ID="$(az ad app list --display-name "$APP_NAME" --query '[0].appId' -o tsv)"
if [[ -z "$APP_ID" || "$APP_ID" == "null" ]]; then
  die "App '$APP_NAME' not found. Run the setup script in the-best-of-brock-cookbook first."
fi
say "Found shared App: $APP_ID"

add_federated_cred() {
  local name="$1" subject="$2" existing
  existing="$(az ad app federated-credential list --id "$APP_ID" \
    --query "[?name=='$name'].name" -o tsv)"
  if [[ -n "$existing" ]]; then
    say "Federated credential '$name' already present"
    return
  fi
  say "Adding federated credential '$name' (subject: $subject)"
  az ad app federated-credential create --id "$APP_ID" --parameters "{
    \"name\":        \"$name\",
    \"issuer\":      \"https://token.actions.githubusercontent.com\",
    \"subject\":     \"$subject\",
    \"audiences\":   [\"api://AzureADTokenExchange\"],
    \"description\": \"GitHub Actions OIDC for $GITHUB_OWNER/$GITHUB_REPO ($name)\"
  }" >/dev/null
}

add_federated_cred \
  "gh-wordrecovery-SF-release-env" \
  "repo:$GITHUB_OWNER/$GITHUB_REPO:environment:release"

add_federated_cred \
  "gh-wordrecovery-SF-main-branch" \
  "repo:$GITHUB_OWNER/$GITHUB_REPO:ref:refs/heads/main"

cat <<EOF

All done. Set these as **GitHub repo Variables** (Settings -> Secrets and
variables -> Actions -> Variables tab) for $GITHUB_OWNER/$GITHUB_REPO:

  AZURE_CLIENT_ID        $APP_ID
  AZURE_TENANT_ID        $TENANT_ID
  AZURE_SUBSCRIPTION_ID  $SUBSCRIPTION_ID
  TS_ENDPOINT            $TS_ENDPOINT
  TS_ACCOUNT_NAME        $SIGNING_ACCOUNT
  TS_CERT_PROFILE        $CERT_PROFILE

No secrets needed — auth is federated via OIDC, zero client secrets.

Next:
  1. Paste those six values into GitHub repo Variables.
  2. Create an Environment called 'release'
       (Settings -> Environments -> New environment -> 'release').
  3. Run Actions -> Verify OIDC -> Run workflow to confirm federation.
  4. Tag a release (git tag vX.Y.Z && git push --tags); the release
     workflow signs the Windows .exe automatically.
EOF
