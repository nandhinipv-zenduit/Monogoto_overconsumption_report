# ==========================================================
# Monogoto SIM Usage Audit — monthly overconsumption report
# ==========================================================
# Runs main.py, which pulls the Monogoto SIM inventory + a 31-day usage
# report, joins it to Zoho Analytics, builds an Excel workbook, and emails
# it via Gmail OAuth2.
#
# SECRETS REQUIRED (Settings -> Secrets and variables -> Actions):
#   MONOGOTO_USERNAME                     Monogoto console login
#   MONOGOTO_PASSWORD
#   GMAIL_USERNAME                        sending mailbox, e.g. reports@zenduit.com
#   GMAIL_CLIENT_ID                       OAuth client (…apps.googleusercontent.com)
#   GMAIL_CLIENT_SECRET                   OAuth client secret (GOCSPX-…)
#   GMAIL_REFRESH_TOKEN                   minted by get_gmail_refresh_token.py
#   ZOHO_CLIENT_ID_ANALYTICS              Analytics-scoped, NOT the CRM ones
#   ZOHO_CLIENT_SECRET_ANALYTICS
#   ZOHO_CLIENT_REFRESH_TOKEN_ANALYTICS
#
# GMAIL_PASS is no longer read — delete that secret once this workflow runs
# green, so nobody is misled into "fixing" the app password later.
# ==========================================================

name: Monogoto SIM Usage Audit

on:
  # 04:00 UTC on the 1st of each month = 09:30 IST. The script's window is
  # T-31 -> T-1, so running on the 1st covers (near enough) the month just
  # ended. GitHub's scheduler queues cron jobs under load and can start
  # them late by tens of minutes — fine here, nothing downstream is timing-
  # sensitive, but it's why this isn't pinned to a precise hour.
  schedule:
    - cron: "0 4 1 * *"

  workflow_dispatch:
    inputs:
      report_timeout_seconds:
        description: "Seconds to wait for the Monogoto report to finish (default 900)"
        required: false
        default: "900"

# Never let a manual run collide with a scheduled one — two runs would
# generate duplicate Monogoto report templates and email the team twice.
concurrency:
  group: monogoto-overconsumption
  cancel-in-progress: false

permissions:
  contents: read

jobs:
  run-report:
    runs-on: ubuntu-latest
    # /things is capped at 50 per page and fetched sequentially, and the
    # usage report is polled for up to 15 minutes, so this job is slow by
    # design. 90 minutes leaves headroom without letting a hung run burn
    # the full 6-hour default.
    timeout-minutes: 90

    steps:
      - name: Check out repository
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"
          cache: pip

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          if [ -f requirements.txt ]; then
            pip install -r requirements.txt
          else
            pip install aiohttp yarl requests pandas XlsxWriter
          fi

      - name: Verify required secrets are present
        # Fails in 2 seconds with the names of what's missing, instead of
        # 20 minutes later inside the script — and never echoes a value.
        env:
          REQUIRED: >-
            MONOGOTO_USERNAME MONOGOTO_PASSWORD
            GMAIL_USERNAME GMAIL_CLIENT_ID GMAIL_CLIENT_SECRET GMAIL_REFRESH_TOKEN
            ZOHO_CLIENT_ID_ANALYTICS ZOHO_CLIENT_SECRET_ANALYTICS ZOHO_CLIENT_REFRESH_TOKEN_ANALYTICS
          MONOGOTO_USERNAME: ${{ secrets.MONOGOTO_USERNAME }}
          MONOGOTO_PASSWORD: ${{ secrets.MONOGOTO_PASSWORD }}
          GMAIL_USERNAME: ${{ secrets.GMAIL_USERNAME }}
          GMAIL_CLIENT_ID: ${{ secrets.GMAIL_CLIENT_ID }}
          GMAIL_CLIENT_SECRET: ${{ secrets.GMAIL_CLIENT_SECRET }}
          GMAIL_REFRESH_TOKEN: ${{ secrets.GMAIL_REFRESH_TOKEN }}
          ZOHO_CLIENT_ID_ANALYTICS: ${{ secrets.ZOHO_CLIENT_ID_ANALYTICS }}
          ZOHO_CLIENT_SECRET_ANALYTICS: ${{ secrets.ZOHO_CLIENT_SECRET_ANALYTICS }}
          ZOHO_CLIENT_REFRESH_TOKEN_ANALYTICS: ${{ secrets.ZOHO_CLIENT_REFRESH_TOKEN_ANALYTICS }}
        run: |
          missing=""
          for name in $REQUIRED; do
            if [ -z "${!name}" ]; then missing="$missing $name"; fi
          done
          if [ -n "$missing" ]; then
            echo "::error::Missing repository secrets:$missing"
            exit 1
          fi
          echo "All 9 required secrets are set."

      - name: Run Monogoto overconsumption report
        env:
          # --- Monogoto ---
          MONOGOTO_USERNAME: ${{ secrets.MONOGOTO_USERNAME }}
          MONOGOTO_PASSWORD: ${{ secrets.MONOGOTO_PASSWORD }}
          MONOGOTO_REPORT_TIMEOUT_SECONDS: ${{ github.event.inputs.report_timeout_seconds || '900' }}

          # --- Gmail (OAuth2 — no app password) ---
          GMAIL_USERNAME: ${{ secrets.GMAIL_USERNAME }}
          GMAIL_CLIENT_ID: ${{ secrets.GMAIL_CLIENT_ID }}
          GMAIL_CLIENT_SECRET: ${{ secrets.GMAIL_CLIENT_SECRET }}
          GMAIL_REFRESH_TOKEN: ${{ secrets.GMAIL_REFRESH_TOKEN }}

          # --- Zoho Analytics ---
          # Analytics-specific names on purpose: the script falls back to
          # generic ZOHO_* vars, but if those hold CRM credentials the
          # exports fail with INVALID_OAUTHSCOPE (errorCode 8540).
          ZOHO_CLIENT_ID_ANALYTICS: ${{ secrets.ZOHO_CLIENT_ID_ANALYTICS }}
          ZOHO_CLIENT_SECRET_ANALYTICS: ${{ secrets.ZOHO_CLIENT_SECRET_ANALYTICS }}
          ZOHO_CLIENT_REFRESH_TOKEN_ANALYTICS: ${{ secrets.ZOHO_CLIENT_REFRESH_TOKEN_ANALYTICS }}
          ZOHO_ORG_ID: ${{ vars.ZOHO_ORG_ID || '67409019' }}
          ZOHO_ANALYTICS_EXPORT_TIMEOUT_SECONDS: "300"

          # Keep a copy on disk so the workbook survives an email failure.
          SAVE_REPORT_PATH: report/monogoto_overconsumption_report.xlsx

          PYTHONUNBUFFERED: "1"
        run: python main.py

      - name: Upload report as a run artifact
        # if: always() — the whole point is to keep the workbook when the
        # email step is what failed. The file is written before send_email()
        # is called, so it exists even on an SMTP error.
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: monogoto-overconsumption-report
          path: report/*.xlsx
          if-no-files-found: warn
          retention-days: 30
