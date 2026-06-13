# Code Signing

Windows `.exe` binaries produced by the release workflow are signed with
**Azure Trusted Signing** using GitHub Actions OIDC federated credentials. No
client secrets live in the repo, and no signing certificate is ever downloaded
to the CI runner — GitHub Actions gets a short-lived Azure token, the signing
happens inside Azure, and the signed binary comes back to the runner.

This repo reuses a **shared signing identity** (one Azure App Registration and
one Trusted Signing certificate profile) that is used across socrtwo's
projects, so there is no new certificate to buy or manage here.

macOS and mobile artifacts in releases remain **unsigned** (Apple signing
requires an Apple Developer Program membership; see the org-level notes).

## One-time setup

1. Open [shell.azure.com](https://shell.azure.com) (Bash) in the Azure
   subscription that owns the Trusted Signing account.

2. Clone the repo and run the setup script:

   ```bash
   git clone https://github.com/socrtwo/wordrecovery-SF.git
   cd wordrecovery-SF
   bash scripts/setup-azure-signing.sh
   ```

   The script is idempotent. It locates the shared App Registration
   (`brock-cookbook-signer`, created once by the-best-of-brock-cookbook) and
   adds two federated credentials scoped to **this** repo:
   - `environment:release` — used by the Windows signing job.
   - `ref:refs/heads/main` — used by the OIDC smoke-test workflow.

   It does **not** create a new app, certificate, or role assignment — the
   shared identity already holds the **Artifact Signing Certificate Profile
   Signer** role on the cert profile.

3. The script prints six values. Paste them into the GitHub repo as
   **Variables** (not Secrets) under
   **Settings → Secrets and variables → Actions → Variables**:

   | Variable | Value |
   |---|---|
   | `AZURE_CLIENT_ID`       | (printed by the script — the shared App's client ID) |
   | `AZURE_TENANT_ID`       | `99dd7dad-5263-44fc-aea5-63c220ebcd48` |
   | `AZURE_SUBSCRIPTION_ID` | `5b2f1523-5925-4630-a326-22610b97db02` |
   | `TS_ENDPOINT`           | `https://eus.codesigning.azure.net/` |
   | `TS_ACCOUNT_NAME`       | `S2SetvicesCodeSigning` |
   | `TS_CERT_PROFILE`       | `GitHubAppSigner752` |

4. Create the `release` environment in
   **Settings → Environments → New environment → `release`**.
   The Windows signing job is gated on this environment; the federated
   credential is scoped to it, so signing only happens for jobs that opt in.

## Graceful no-op until configured

The signing steps are guarded by `if: vars.AZURE_TENANT_ID != ''`. Until you
set the repo Variables above, the release workflow still runs and produces the
**unsigned** `.exe` exactly as before — it simply skips the signing steps. Once
the Variables are set, the same workflow starts signing automatically.

## Verify it works (no signing consumed)

Go to **Actions → Verify OIDC → Run workflow**. It authenticates to Azure via
OIDC and prints the identity without consuming any signing quota. If it fails
with `AADSTS70021: No matching federated identity record found`, re-run
`scripts/setup-azure-signing.sh` (it's idempotent).

## Verify the signature on a released .exe

```powershell
# PowerShell (Windows)
Get-AuthenticodeSignature .\YourApp-Setup.exe | Format-List
```

Expect `Status: Valid`, an issuer including `Microsoft ID Verified CS EOC CA`,
and a timestamp from `timestamp.acs.microsoft.com`. On macOS/Linux you can use
`osslsigncode verify -in YourApp-Setup.exe`.
