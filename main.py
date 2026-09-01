import os
import io
import re
import csv
import json
import time
import gzip
import zipfile
import asyncio
import aiohttp
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from aiohttp import ClientTimeout
from yarl import URL
import smtplib
from email.message import EmailMessage
from io import BytesIO

# ==========================================================
# WHAT THIS IS
# ==========================================================
# Monogoto overconsumption report — sibling to daily_sim_report.py (1NCE)
# and telenor_sim_report.py (Telenor / Ericsson IoT Accelerator), same
# overall shape: env-var credentials, a 31-day trailing usage window
# (T-31 -> T-1) compared against the Zenduit data plan, an in-memory
# Excel report, and Gmail delivery.
#
# THE PLATFORM: Monogoto (https://docs.monogoto.io, public docs, no login
# needed to read them).
#   1) Auth: POST https://console.monogoto.io/Auth with
#      {"UserName": ..., "Password": ...} -> JWT `token` + `CustomerId`.
#      Base host is console.monogoto.io (no separate api.monogoto.io).
#      Some endpoints additionally require an `apikey: {CustomerId}`
#      header alongside the bearer token, so both are captured at auth.
#   2) SIM inventory: GET /things, paginated with limit+offset.
#      limit is CAPPED AT 50 by the API, so a few thousand SIMs means a
#      few hundred sequential pages. Response is a BARE JSON ARRAY (no
#      envelope, no total count). ICCID lives in `ExternalUniqueId`.
#   3) Usage: the async REPORT flow — unlike Telenor's per-day report
#      files, Monogoto's report API accepts an arbitrary date RANGE in one
#      call, so the whole 31-day window is fetched in a single report
#      instead of a per-day loop:
#         POST /report-template              -> ReportTemplateId_...
#         POST /report-history/{templateId}  -> ReportHistoryId_...
#         GET  /report-history/byTemplate    -> poll until csvPath appears
#         GET  /report-history/downloadReport/csv/{urlencoded csvPath}
#      Dates are EPOCH MILLISECONDS. The CSV is per-SIM (one row per
#      thing x roaming partner), so rows are summed per ICCID here.
#   4) AccountId / Account Name mapping: identical Zoho Analytics logic to
#      the 1NCE/Telenor scripts (no Zenduit website/API calls).
#
# TWO PARSING QUIRKS IN THE MONOGOTO CSV (handled below):
#   a) No ICCID column — it's embedded in "Thing Name"
#      ("ICCID 8912372646888991, 8912372646888991"), extracted with a
#      regex. IMSI is a real column and is kept as a fallback join key.
#   b) The "Data" column is a unit-suffixed string (" 140.25 MB"), not a
#      bare number; _data_string_to_mb() normalises everything to MB.
#      Anything unitless is ASSUMED to already be MB and a warning listing
#      sample raw values is printed so it can be confirmed on first run.
#
# THINGS TO CONFIRM ON YOUR FIRST RUN:
#   - GroupBy is "servingNetwork" (the value shown working in Monogoto's
#     docs); it yields one row per SIM per roaming partner, which is why
#     rows are summed per ICCID.
#   - Monogoto publishes no rate limits — page fetching for /things is
#     serialised with a small polite delay rather than fanned out.
#   - Whether a single report call really covers a 31-day range end to
#     end, or whether Monogoto caps report windows shorter than that — the
#     poll loop and REPORT_TIMEOUT_SECONDS give it room to run long; if the
#     report never finishes, split the range into chunks and tell me.
# ==========================================================

# ==========================================================
# GMAIL CONFIG (GLOBAL) — OAuth 2.0, not an app password.
# ==========================================================
# WHY THIS CHANGED
#   smtp.gmail.com used to be fed a 16-character "app password". Those get
#   auto-revoked whenever the account password changes, 2FA is re-enrolled,
#   or a Workspace admin tightens policy — which shows up here as:
#       smtplib.SMTPAuthenticationError: (535, b'5.7.8 Username and
#       Password not accepted ... BadCredentials')
#   OAuth refresh tokens don't expire on password changes, so the job stops
#   breaking every time someone rotates a password.
#
# WHAT GOOGLE ACTUALLY NEEDS
#   A client ID + client secret ALONE cannot authenticate SMTP — they only
#   identify the app, they prove nothing about the mailbox. The mailbox
#   consent is carried by a REFRESH TOKEN, minted once by the sending
#   account. At run time the three are exchanged for a short-lived access
#   token, which is what SMTP's AUTH XOAUTH2 accepts.
#
# ENV VARS EXPECTED (set these as GitHub Actions secrets):
#   GMAIL_USERNAME       the sending mailbox, e.g. reports@zenduit.com
#   GMAIL_CLIENT_ID      OAuth client ID     (…apps.googleusercontent.com)
#   GMAIL_CLIENT_SECRET  OAuth client secret (GOCSPX-…)
#   GMAIL_REFRESH_TOKEN  refresh token minted by GMAIL_USERNAME for the
#                        https://mail.google.com/ scope
#   GMAIL_PASS           no longer used — delete the secret once this runs
#
# HOW TO MINT THE REFRESH TOKEN (once, ~3 minutes):
#   Run get_gmail_refresh_token.py, shipped alongside this file. Or by hand
#   in Google Cloud Console: create an OAuth client of type "Desktop app",
#   then in OAuth Playground (⚙ -> "Use your own OAuth credentials") consent
#   as the sending account to scope https://mail.google.com/ and exchange
#   the auth code for a refresh token.
#
# GOTCHAS THAT WILL BITE
#   - Scope must be https://mail.google.com/ (full). gmail.send is enough
#     for the Gmail API but NOT for SMTP XOAUTH2.
#   - While the OAuth consent screen is in "Testing", refresh tokens expire
#     after 7 days. Publish the app (or keep it Internal on the Workspace)
#     or this job dies weekly.
#   - The refresh token belongs to the account that consented; it must be
#     the same mailbox as GMAIL_USERNAME or Google returns 535 again.
# ==========================================================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GMAIL_TOKEN_URL = "https://oauth2.googleapis.com/token"

EMAIL_SENDER = os.getenv("GMAIL_USERNAME")
GMAIL_CLIENT_ID = os.getenv("GMAIL_CLIENT_ID")
GMAIL_CLIENT_SECRET = os.getenv("GMAIL_CLIENT_SECRET")
GMAIL_REFRESH_TOKEN = os.getenv("GMAIL_REFRESH_TOKEN")

EMAIL_TO = [
    "nandhinipv@zenduit.com", "nikithavinod@zenduit.com", "abidali@gofleet.com",
      "rizamae@gofleet.com",
]

_missing_gmail = [
    name for name, value in (
        ("GMAIL_USERNAME", EMAIL_SENDER),
        ("GMAIL_CLIENT_ID", GMAIL_CLIENT_ID),
        ("GMAIL_CLIENT_SECRET", GMAIL_CLIENT_SECRET),
        ("GMAIL_REFRESH_TOKEN", GMAIL_REFRESH_TOKEN),
    ) if not value
]
if _missing_gmail:
    raise RuntimeError(
        f"❌ Gmail OAuth credentials not found in environment variables: {_missing_gmail}. "
        f"This script no longer uses GMAIL_PASS (app password) — see the Gmail config "
        f"block at the top of this file for how to mint GMAIL_REFRESH_TOKEN."
    )

# ==========================================================
# MONOGOTO CONFIG
# ==========================================================
MONOGOTO = {
    "base_url": "https://console.monogoto.io",
    "username": os.getenv("MONOGOTO_USERNAME"),
    "password": os.getenv("MONOGOTO_PASSWORD"),
}
if not MONOGOTO["username"] or not MONOGOTO["password"]:
    raise RuntimeError("❌ Monogoto credentials not found in environment variables "
                        "(MONOGOTO_USERNAME / MONOGOTO_PASSWORD)")

THINGS_PAGE_LIMIT = 50      # hard API cap - larger values are rejected
THINGS_PAGE_DELAY = 0.05    # small politeness delay between pages (no documented rate limit)
REPORT_POLL_SECONDS = 5
REPORT_TIMEOUT_SECONDS = int(os.getenv("MONOGOTO_REPORT_TIMEOUT_SECONDS", "900"))

# ==========================================================
# ZOHO ANALYTICS CONFIG — same source as the 1NCE/Telenor scripts use for
# Zenduit-side data (no calls to Zenduit's own website/API for Monogoto).
# ==========================================================
ZOHO_ANALYTICS_DOMAIN = os.getenv("ZOHO_ANALYTICS_DOMAIN", "analyticsapi.zoho.com")
ZOHO_ACCOUNTS_DOMAIN = os.getenv("ZOHO_ACCOUNTS_DOMAIN", "accounts.zoho.com")
ZOHO_ANALYTICS_API = f"https://{ZOHO_ANALYTICS_DOMAIN}/restapi/v2"
ZOHO_ORG_ID = os.getenv("ZOHO_ORG_ID", "67409019")
ZOHO_ANALYTICS_WORKSPACE_ID = "953790000013364003"
ZOHO_ANALYTICS_ZENDUIT_DEVICES_VIEW_ID = "953790000054827175"   # "Zenduit Devices" view
ZOHO_ANALYTICS_ACCOUNTS_VIEW_ID = "953790000013364024"          # "Accounts" view
ZOHO_ANALYTICS_EXPORT_TIMEOUT_SECONDS = int(os.getenv("ZOHO_ANALYTICS_EXPORT_TIMEOUT_SECONDS", "300"))


def _pick_env(*names):
    """First-match-wins env var lookup that also reports which var supplied
    the value — matters because falling back to generic ZOHO_* (CRM) vars
    authenticates fine but gets rejected by Analytics with INVALID_OAUTHSCOPE."""
    for n in names:
        v = os.environ.get(n)
        if v:
            return v, n
    return None, None


ZOHO_ANALYTICS_CLIENT_ID, _ZOHO_CLIENT_ID_SRC = _pick_env(
    "ZOHO_CLIENT_ID_ANALYTICS", "ZOHO_ANALYTICS_CLIENT_ID", "ZOHO_CLIENT_ID")
ZOHO_ANALYTICS_CLIENT_SECRET, _ZOHO_CLIENT_SECRET_SRC = _pick_env(
    "ZOHO_CLIENT_SECRET_ANALYTICS", "ZOHO_ANALYTICS_CLIENT_SECRET", "ZOHO_CLIENT_SECRET")
ZOHO_ANALYTICS_REFRESH_TOKEN, _ZOHO_REFRESH_TOKEN_SRC = _pick_env(
    "ZOHO_CLIENT_REFRESH_TOKEN_ANALYTICS", "ZOHO_ANALYTICS_REFRESH_TOKEN", "ZOHO_REFRESH_TOKEN")

_ZOHO_ANALYTICS_SPECIFIC = {
    "ZOHO_CLIENT_ID_ANALYTICS", "ZOHO_ANALYTICS_CLIENT_ID",
    "ZOHO_CLIENT_SECRET_ANALYTICS", "ZOHO_ANALYTICS_CLIENT_SECRET",
    "ZOHO_CLIENT_REFRESH_TOKEN_ANALYTICS", "ZOHO_ANALYTICS_REFRESH_TOKEN",
}


def _report_zoho_credential_sources():
    srcs = {
        "client_id": _ZOHO_CLIENT_ID_SRC,
        "client_secret": _ZOHO_CLIENT_SECRET_SRC,
        "refresh_token": _ZOHO_REFRESH_TOKEN_SRC,
    }
    print(f"Zoho credentials sourced from env vars: {srcs}")
    unset = [k for k, v in srcs.items() if v is None]
    if unset:
        print(f"⚠️ No env var found at all for: {unset}. Copy the Zoho env vars from "
              f"daily_sim_report.py / telenor_sim_report.py's run config if this is a "
              f"different run config.")
        return
    generic = {k: v for k, v in srcs.items() if v not in _ZOHO_ANALYTICS_SPECIFIC}
    if generic:
        print(f"⚠️ These fell back to GENERIC Zoho env vars rather than Analytics-specific "
              f"ones: {generic}. If those hold Zoho CRM credentials, the export calls below "
              f"will fail with INVALID_OAUTHSCOPE (errorCode 8540).")


# ==========================================================
# SHARED HELPERS
# ==========================================================
def _numeric_fillna(series, fill=0):
    return pd.to_numeric(series, errors="coerce").fillna(fill)


def _find_col(df, candidates):
    lower_map = {str(c).lower(): c for c in df.columns if c is not None}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def _epoch_ms(dt):
    return int(dt.replace(tzinfo=timezone.utc).timestamp() * 1000)


_UNIT_TO_MB = {
    "B": 1.0 / (1024 * 1024),
    "KB": 1.0 / 1024,
    "MB": 1.0,
    "GB": 1024.0,
    "TB": 1024.0 * 1024.0,
}
_DATA_RE = re.compile(r"^\s*([0-9][0-9,]*\.?[0-9]*)\s*([KMGT]?B)?\s*$", re.IGNORECASE)
_unitless_samples = []


def _data_string_to_mb(value):
    """Monogoto's report CSV writes data usage as a unit-suffixed string
    (" 140.25 MB"), and the unit varies with magnitude, so a plain
    to_numeric() would silently produce NaN for every row. Returns MB."""
    if value is None:
        return 0.0
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", "-"):
        return 0.0
    m = _DATA_RE.match(s)
    if not m:
        return 0.0
    number = float(m.group(1).replace(",", ""))
    unit = (m.group(2) or "").upper()
    if not unit:
        if len(_unitless_samples) < 5:
            _unitless_samples.append(s)
        return number
    return number * _UNIT_TO_MB.get(unit, 1.0)


_ICCID_IN_TEXT = re.compile(r"(\d{15,22})")


def _iccid_from_thing_name(value):
    """The report CSV has no ICCID column; Thing Name carries it, e.g.
    'ICCID 8912372646888991, 8912372646888991'."""
    if value is None:
        return None
    m = _ICCID_IN_TEXT.search(str(value))
    return m.group(1) if m else None


# ==========================================================
# MONOGOTO AUTH — POST /Auth, returns a JWT plus the CustomerId that some
# endpoints need as a separate `apikey` header.
# ==========================================================
async def get_monogoto_token():
    async with aiohttp.ClientSession() as session:
        async with session.post(
            f"{MONOGOTO['base_url']}/Auth",
            json={"UserName": MONOGOTO["username"], "Password": MONOGOTO["password"]},
            headers={"Accept": "application/json"},
            timeout=ClientTimeout(total=30),
        ) as r:
            text = await r.text()
            if r.status != 200:
                raise RuntimeError(f"Monogoto auth failed: {r.status} — {text[:400]}")
            try:
                data = json.loads(text)
            except json.JSONDecodeError:
                raise RuntimeError(f"Monogoto auth returned non-JSON: {text[:400]}")

    token = data.get("token") or data.get("Token")
    customer_id = data.get("CustomerId") or data.get("customerId")
    if not token:
        raise RuntimeError(f"Monogoto auth: no token in response. Keys present: {list(data.keys())}")
    if not customer_id:
        print(f"⚠️ Monogoto auth returned no CustomerId (keys: {list(data.keys())}). "
              f"Endpoints requiring the `apikey` header may fail.")
    print(f"✅ Monogoto auth OK (CustomerId={customer_id}, role={data.get('Role')})")
    return token, customer_id


def _mono_headers(token, customer_id=None, accept="application/json"):
    h = {"Authorization": f"Bearer {token}", "Accept": accept}
    if customer_id:
        h["apikey"] = str(customer_id)
    return h


# ==========================================================
# MONOGOTO THINGS (= SIMs) — GET /things, limit/offset paginated.
# Response is a bare JSON array; ICCID is ExternalUniqueId.
# ==========================================================
async def fetch_all_things(token, customer_id):
    headers = _mono_headers(token, customer_id)
    rows = []
    offset = 0

    async with aiohttp.ClientSession() as session:
        while True:
            qs = [
                f"limit={THINGS_PAGE_LIMIT}",
                f"offset={offset}",
                "sortBy%5BThingName%5D=ASC",
            ]
            url = URL(f"{MONOGOTO['base_url']}/things?{'&'.join(qs)}", encoded=True)
            async with session.get(url, headers=headers, timeout=ClientTimeout(total=60)) as r:
                text = await r.text()
                if r.status != 200:
                    raise RuntimeError(f"Monogoto /things failed at offset {offset}: "
                                       f"{r.status} — {text[:400]}")
                try:
                    page = json.loads(text)
                except json.JSONDecodeError:
                    raise RuntimeError(f"Monogoto /things returned non-JSON at offset {offset}: "
                                       f"{text[:400]}")

            if isinstance(page, dict):
                page = page.get("things") or page.get("data") or page.get("dbResponse") or []
            if not isinstance(page, list):
                raise RuntimeError(f"Monogoto /things: unexpected payload type {type(page)} "
                                   f"at offset {offset}")

            rows.extend(page)

            if offset == 0 and page:
                print("--- Sample Thing record (1 of first page) ---")
                print(json.dumps(page[0], indent=2, default=str)[:2500])
                print("-" * 80)

            if len(page) < THINGS_PAGE_LIMIT:
                break
            offset += THINGS_PAGE_LIMIT
            if offset % (THINGS_PAGE_LIMIT * 20) == 0:
                print(f"  ...fetched {len(rows)} things so far")
            if THINGS_PAGE_DELAY:
                await asyncio.sleep(THINGS_PAGE_DELAY)

    print(f"✅ Total Monogoto SIMs fetched: {len(rows)}")
    if not rows:
        return pd.DataFrame(columns=["ICCID", "IMSI", "MSISDN", "SIM_Status", "ThingId",
                                      "ThingName", "Monogoto_Lifetime_Bytes"])

    df = pd.DataFrame(rows)
    out = pd.DataFrame()

    iccid_col = _find_col(df, ["ExternalUniqueId", "ICCID", "iccid"])
    if not iccid_col:
        print(f"⚠️ No ICCID column found on Thing records. Columns present: {list(df.columns)}")
        out["ICCID"] = None
    else:
        out["ICCID"] = df[iccid_col].astype(str).str.strip()

    imsi_col = _find_col(df, ["ActiveMobileSubscriber", "IMSI", "imsi"])
    out["IMSI"] = df[imsi_col].astype(str).str.strip() if imsi_col else None

    msisdn_col = _find_col(df, ["AddressSignal", "MSISDN", "msisdn"])
    out["MSISDN"] = df[msisdn_col] if msisdn_col else None

    state_col = _find_col(df, ["State", "state", "BillingState"])
    out["SIM_Status"] = df[state_col] if state_col else None

    out["ThingId"] = df.get("ThingId")
    out["ThingName"] = df.get("ThingName")

    # `Data` on a Thing is a cumulative/lifetime byte counter, kept for
    # reference - it is NOT the window's usage (that comes from the report).
    data_col = _find_col(df, ["Data"])
    out["Monogoto_Lifetime_Bytes"] = pd.to_numeric(df[data_col], errors="coerce") if data_col else None

    out = out[out["ICCID"].notna() & (out["ICCID"].astype(str).str.strip() != "")]
    return out.drop_duplicates(subset=["ICCID"], keep="first")


# ==========================================================
# MONOGOTO USAGE — the async report flow, run ONCE over the full 31-day
# window (T-31 -> T-1) rather than per day, since Monogoto's report API
# accepts an arbitrary date range (unlike Telenor's per-day report files).
# ==========================================================
async def fetch_usage_report(token, customer_id, start_date, end_date):
    headers = _mono_headers(token, customer_id)
    base = MONOGOTO["base_url"]

    start_dt = datetime(start_date.year, start_date.month, start_date.day)
    end_dt = datetime(end_date.year, end_date.month, end_date.day) + timedelta(days=1) - timedelta(seconds=1)
    start_ms, end_ms = _epoch_ms(start_dt), _epoch_ms(end_dt)
    report_name = f"zenduit_overconsumption_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}"
    empty = pd.DataFrame(columns=["ICCID", "IMSI", "Monogoto_MB_Usage"])

    async with aiohttp.ClientSession() as session:
        # --- Step 1: create the report template -----------------------
        payload = {
            "Name": report_name,
            "Description": f"Zenduit overconsumption report for {start_date.isoformat()} to {end_date.isoformat()}",
            "GroupBy": "servingNetwork",
            "rateByCurrentPricePlan": True,
            "ReportType": "rating",
            "ReportPeriod": {"startDate": start_ms, "endDate": end_ms},
        }
        async with session.post(f"{base}/report-template", headers=headers, json=payload,
                                timeout=ClientTimeout(total=60)) as r:
            text = (await r.text()).strip()
            if r.status not in (200, 201):
                print(f"⚠️ Monogoto create report-template failed: {r.status} — {text[:300]}")
                return empty
        template_id = text.strip('"')
        print(f"  Monogoto report template created: {template_id} "
              f"(window {start_ms} -> {end_ms}, i.e. {start_date} to {end_date})")

        # --- Step 2: kick off report generation -----------------------
        async with session.post(f"{base}/report-history/{quote(template_id, safe='')}",
                                headers=headers, timeout=ClientTimeout(total=60)) as r:
            text = (await r.text()).strip()
            if r.status not in (200, 201, 202):
                print(f"⚠️ Monogoto generate-report failed: {r.status} — {text[:300]}")
                return empty
        history_id = text.strip('"')
        print(f"  Monogoto report generation started: {history_id}")

        # --- Step 3: poll until the CSV path shows up -----------------
        csv_path = None
        deadline = time.time() + REPORT_TIMEOUT_SECONDS
        last_log = 0.0
        while time.time() < deadline:
            async with session.get(f"{base}/report-history/byTemplate", headers=headers,
                                   timeout=ClientTimeout(total=60)) as r:
                text = await r.text()
                if r.status != 200:
                    print(f"⚠️ Monogoto report-history/byTemplate failed: {r.status} — {text[:300]}")
                    return empty
                try:
                    body = json.loads(text)
                except json.JSONDecodeError:
                    print(f"⚠️ report-history/byTemplate returned non-JSON: {text[:300]}")
                    return empty

            entries = body.get("dbResponse") if isinstance(body, dict) else body
            entries = entries or []
            match = next((e for e in entries if e.get("ReportHistoryId") == history_id
                          and e.get("csvPath")), None)
            if not match:
                match = next((e for e in entries if e.get("ReportTemplateId") == template_id
                              and e.get("csvPath")), None)
            if match:
                csv_path = match["csvPath"]
                print(f"  Monogoto report ready: {csv_path}")
                break

            if time.time() - last_log >= 30:
                print(f"  ...waiting for Monogoto report {history_id} to finish "
                      f"({int(deadline - time.time())}s left)")
                last_log = time.time()
            await asyncio.sleep(REPORT_POLL_SECONDS)

        if not csv_path:
            print(f"⚠️ Monogoto report {history_id} did not finish within "
                  f"{REPORT_TIMEOUT_SECONDS}s. Usage will be empty this run. If a 31-day "
                  f"window is consistently too slow, split it into chunks (e.g. weekly) "
                  f"and sum the results instead of one call.")
            return empty

        # --- Step 4: download the CSV ---------------------------------
        dl_url = f"{base}/report-history/downloadReport/csv/{quote(csv_path, safe='')}"
        async with session.get(dl_url, headers=_mono_headers(token, customer_id, accept="*/*"),
                               timeout=ClientTimeout(total=180)) as r:
            raw = await r.read()
            if r.status != 200:
                print(f"⚠️ Monogoto CSV download failed: {r.status} — {raw[:300]!r}")
                return empty
            print(f"  downloaded Monogoto report CSV "
                  f"(Content-Type: {r.headers.get('Content-Type')}, {len(raw)} bytes)")

    # Defensive: handle zip/gzip in case the CSV ever arrives compressed.
    if raw[:4] == b"PK\x03\x04":
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            names = [n for n in zf.namelist() if not n.endswith("/")]
            print(f"  report was a ZIP containing: {names}")
            raw = zf.read(names[0]) if names else b""
    elif raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)

    text = raw.decode("utf-8-sig", errors="replace")
    rows = list(csv.DictReader(io.StringIO(text), restkey="_extra_values"))
    if not rows:
        print("⚠️ Monogoto report CSV parsed to zero rows.")
        return empty

    print(f"--- Monogoto report CSV: {len(rows)} rows. Columns: {list(rows[0].keys())} ---")
    print(f"--- Sample row: {rows[0]} ---")

    df = pd.DataFrame(rows)
    name_col = _find_col(df, ["Thing Name", "ThingName", "Thing_Name"])
    imsi_col = _find_col(df, ["IMSI", "imsi"])
    data_col = _find_col(df, ["Data", "Data (MB)", "Data_MB", "Total Data"])

    if not data_col or not (name_col or imsi_col):
        print(f"⚠️ Couldn't identify the Thing Name/IMSI + Data columns in the Monogoto report "
              f"(name_col={name_col!r}, imsi_col={imsi_col!r}, data_col={data_col!r}). "
              f"Columns seen: {list(df.columns)}. Usage will be empty this run.")
        return empty

    out = pd.DataFrame()
    out["ICCID"] = df[name_col].map(_iccid_from_thing_name) if name_col else None
    out["IMSI"] = df[imsi_col].astype(str).str.strip() if imsi_col else None
    out["Monogoto_MB_Usage"] = df[data_col].map(_data_string_to_mb)

    if _unitless_samples:
        print(f"⚠️ Some Data values had no unit suffix and were ASSUMED to be MB. "
              f"Samples: {_unitless_samples}. Tell me if they're actually bytes/KB "
              f"and I'll change the conversion.")

    matched = out["ICCID"].notna().sum()
    print(f"  ICCID parsed out of Thing Name for {matched}/{len(out)} report rows")

    # One row per SIM per roaming partner -> collapse to one row per SIM.
    if matched:
        grouped = (out.dropna(subset=["ICCID"])
                      .groupby("ICCID", as_index=False)
                      .agg(Monogoto_MB_Usage=("Monogoto_MB_Usage", "sum")))
        grouped["IMSI"] = None
        return grouped

    grouped = (out.dropna(subset=["IMSI"])
                  .groupby("IMSI", as_index=False)
                  .agg(Monogoto_MB_Usage=("Monogoto_MB_Usage", "sum")))
    grouped["ICCID"] = None
    return grouped


# ==========================================================
# ZOHO ANALYTICS — identical to the 1NCE/Telenor scripts. The ONLY source
# of Zenduit-side data (no calls to Zenduit's own website/API).
# ==========================================================
def get_zoho_analytics_token():
    _report_zoho_credential_sources()
    r = requests.post(
        f"https://{ZOHO_ACCOUNTS_DOMAIN}/oauth/v2/token",
        data={
            "refresh_token": ZOHO_ANALYTICS_REFRESH_TOKEN,
            "client_id": ZOHO_ANALYTICS_CLIENT_ID,
            "client_secret": ZOHO_ANALYTICS_CLIENT_SECRET,
            "grant_type": "refresh_token",
        },
        timeout=30,
    )
    data = r.json()
    if r.status_code != 200 or "access_token" not in data:
        raise RuntimeError(f"Zoho Analytics OAuth failed | status={r.status_code} | response={data}")
    granted = data.get("scope")
    print(f"Zoho token acquired. Granted scope: {granted or '(not reported by Zoho)'}")
    if granted and "ZohoAnalytics" not in granted:
        print("⚠️ The granted scope contains no ZohoAnalytics.* entries — exports below will "
              "fail with INVALID_OAUTHSCOPE. The refresh token needs Analytics scopes.")
    return data["access_token"]


def _zoho_headers(token):
    return {"Authorization": f"Zoho-oauthtoken {token}", "ZANALYTICS-ORGID": str(ZOHO_ORG_ID)}


def _zoho_create_job(token, view_id):
    url = f"{ZOHO_ANALYTICS_API}/bulk/workspaces/{ZOHO_ANALYTICS_WORKSPACE_ID}/views/{view_id}/data"
    params = {"CONFIG": json.dumps({"responseFormat": "csv"}, separators=(",", ":"))}
    resp = requests.get(url, params=params, headers=_zoho_headers(token), timeout=60)
    if resp.status_code >= 400:
        raise RuntimeError(f"Zoho Analytics create-export-job (view {view_id}) failed: "
                            f"{resp.status_code}: {resp.text[:500]}")
    job_id = (resp.json().get("data") or {}).get("jobId")
    if not job_id:
        raise RuntimeError(f"Zoho Analytics: no jobId returned for view {view_id}: {resp.text[:300]}")
    return job_id


def _zoho_wait(token, job_id, timeout_seconds=ZOHO_ANALYTICS_EXPORT_TIMEOUT_SECONDS):
    url = f"{ZOHO_ANALYTICS_API}/bulk/workspaces/{ZOHO_ANALYTICS_WORKSPACE_ID}/exportjobs/{job_id}"
    start = time.time()
    deadline = start + timeout_seconds
    last_log = start
    while time.time() < deadline:
        resp = requests.get(url, params={"responseFormat": "json"}, headers=_zoho_headers(token), timeout=60)
        resp.raise_for_status()
        info = resp.json().get("data") or {}
        if info.get("jobStatus") == "JOB COMPLETED" or str(info.get("jobCode")) == "1004":
            return info["downloadUrl"]
        if str(info.get("jobCode")) in ("1003", "1005"):
            raise RuntimeError(f"Zoho Analytics export job {job_id} failed: {info}")
        now = time.time()
        if now - last_log >= 30:
            print(f"  ...still waiting on Zoho Analytics export job {job_id} "
                  f"({now - start:.0f}s elapsed, timeout at {timeout_seconds}s)")
            last_log = now
        time.sleep(2)
    raise RuntimeError(f"Zoho Analytics export job {job_id} timed out after {timeout_seconds}s.")


def _zoho_download(token, url):
    resp = requests.get(url, headers={**_zoho_headers(token), "Accept-Encoding": "identity"}, timeout=180)
    resp.raise_for_status()
    text = resp.text
    return text[1:] if text and text[0] == "﻿" else text


def _flatten(obj, parent="", out=None):
    out = {} if out is None else out
    for k, v in obj.items():
        key = f"{parent}_{k}" if parent else k
        if isinstance(v, dict):
            _flatten(v, key, out)
        else:
            out[key] = v
    return out


def _parse_zoho_export(text):
    if text.lstrip()[:1] in ("{", "["):
        obj = json.loads(text)
        recs = obj.get("data") if isinstance(obj, dict) and "data" in obj else obj
        if isinstance(recs, dict):
            recs = [recs]
        return [_flatten(r) for r in recs if isinstance(r, dict)]
    return list(csv.DictReader(text.splitlines()))


def fetch_analytics_view(token, view_id, label, retries=2, timeout_seconds=ZOHO_ANALYTICS_EXPORT_TIMEOUT_SECONDS):
    """Runs a Zoho Analytics bulk export job and returns parsed rows (empty
    list on failure). Retries the whole job so a flaky export can't take
    down the run — the caller falls back to an empty lookup instead."""
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            job_id = _zoho_create_job(token, view_id)
            print(f"  Zoho Analytics export job created for '{label}': {job_id} (attempt {attempt}/{retries})")
            download_url = _zoho_wait(token, job_id, timeout_seconds=timeout_seconds)
            rows = _parse_zoho_export(_zoho_download(token, download_url))
            print(f"--- {label}: {len(rows)} rows. Columns found: "
                  f"{list(rows[0].keys()) if rows else '(no rows returned)'} ---")
            return rows
        except Exception as e:
            last_err = e
            print(f"⚠️ Zoho Analytics export for '{label}' failed on attempt {attempt}/{retries}: {e}")
    print(f"⚠️ Zoho Analytics export for '{label}' failed after {retries} attempt(s) — giving up. "
          f"Last error: {last_err}")
    return []


_ZENDUIT_DEVICE_COLUMNS = ["ICCID", "AccountId", "Device_Serial", "Zenduit_Data_Plan",
                           "CompanyName", "Zenduit_Usage_MB", "Zenduit_BillingStatus"]


def fetch_zenduit_devices_analytics(token):
    rows = fetch_analytics_view(token, ZOHO_ANALYTICS_ZENDUIT_DEVICES_VIEW_ID, "Zenduit Devices (analytics)")
    if not rows:
        print("⚠️ Zenduit Devices (analytics) export returned no data — "
              "device/account info will be empty for every SIM this run.")
        return pd.DataFrame(columns=_ZENDUIT_DEVICE_COLUMNS)

    df = pd.DataFrame(rows)
    col_map = {
        "ICCID": _find_col(df, ["ICCID", "SIM", "Sim", "SIM Number", "Iccid"]),
        "AccountId": _find_col(df, ["AccountId", "Account_Id", "AccountID", "Account Id",
                                     "CustomerId", "Customer_Id"]),
        "Device_Serial": _find_col(df, ["Device_Serial", "DeviceSerial", "Serial", "Device Serial"]),
        "Zenduit_Data_Plan": _find_col(df, ["Zenduit_Data_Plan", "DataPlan", "Data Plan",
                                             "Data_Plan", "Plan"]),
        "CompanyName": _find_col(df, ["CompanyName", "Company_Name", "Company Name", "Company"]),
        "Zenduit_Usage_MB": _find_col(df, ["Zenduit_Usage_MB", "Data_Usage", "Usage",
                                            "Zenduit_Usage", "Usage_MB", "Data Usage", "DataUsage"]),
        "Zenduit_BillingStatus": _find_col(df, ["Zenduit_BillingStatus", "BillingStatus",
                                                 "Billing_Status", "Billing Status", "Status"]),
    }
    print(f"Column match — Zenduit Devices (analytics): {col_map}")

    if not col_map["ICCID"]:
        print("⚠️ Couldn't find an ICCID/SIM column in the Zenduit Devices view — returning an empty lookup.")
        return pd.DataFrame(columns=_ZENDUIT_DEVICE_COLUMNS)

    missing = [k for k, v in col_map.items() if v is None and k != "ICCID"]
    if missing:
        print(f"⚠️ These fields weren't found and will be blank/0 for every SIM: {missing}.")

    out = pd.DataFrame()
    out["ICCID"] = df[col_map["ICCID"]].astype(str).str.strip()
    out["AccountId"] = df[col_map["AccountId"]].astype(str).str.strip() if col_map["AccountId"] else None
    out["Device_Serial"] = df[col_map["Device_Serial"]].astype(str).str.strip() if col_map["Device_Serial"] else None
    out["Zenduit_Data_Plan"] = (
        pd.to_numeric(df[col_map["Zenduit_Data_Plan"]], errors="coerce").fillna(0)
        if col_map["Zenduit_Data_Plan"] else 0
    )
    out["CompanyName"] = df[col_map["CompanyName"]] if col_map["CompanyName"] else None
    out["Zenduit_Usage_MB"] = (
        pd.to_numeric(df[col_map["Zenduit_Usage_MB"]], errors="coerce").fillna(0)
        if col_map["Zenduit_Usage_MB"] else 0
    )
    out["Zenduit_BillingStatus"] = df[col_map["Zenduit_BillingStatus"]] if col_map["Zenduit_BillingStatus"] else None

    total_rows, unique_iccids = len(out), out["ICCID"].nunique()
    if total_rows != unique_iccids:
        print(f"ℹ️ Zenduit Devices (analytics) has {total_rows} rows for only {unique_iccids} unique "
              f"ICCIDs — keeping the highest-Zenduit_Data_Plan row per ICCID.")
    out = out.sort_values("Zenduit_Data_Plan", ascending=False).drop_duplicates(subset=["ICCID"], keep="first")
    return out


def fetch_account_name_lookup(token):
    rows = fetch_analytics_view(token, ZOHO_ANALYTICS_ACCOUNTS_VIEW_ID, "Accounts (analytics)")
    if not rows:
        print("⚠️ Accounts (analytics) export returned no data — Analytics_Customer_Name will be empty.")
        return pd.DataFrame(columns=["AccountId", "Analytics_Customer_Name"])

    df_accounts = pd.DataFrame(rows)
    account_id_col = _find_col(df_accounts, ["Id", "AccountId", "Account_Id", "AccountID"])
    account_name_col = _find_col(df_accounts, ["Account Name", "AccountName", "Account_Name", "Name",
                                                "CustomerName", "Customer_Name"])
    print(f"Column match — accounts id: {account_id_col!r} | accounts name: {account_name_col!r}")

    if not (account_id_col and account_name_col):
        print("⚠️ Couldn't confidently identify the id/name columns in the Accounts view — "
              "returning an empty lookup.")
        return pd.DataFrame(columns=["AccountId", "Analytics_Customer_Name"])

    df_lookup = df_accounts[[account_id_col, account_name_col]].rename(
        columns={account_id_col: "AccountId", account_name_col: "Analytics_Customer_Name"}
    )
    df_lookup["AccountId"] = df_lookup["AccountId"].astype(str).str.strip()
    return df_lookup.drop_duplicates(subset=["AccountId"], keep="first")


# ==========================================================
# GMAIL OAUTH2 — mint a short-lived access token from the client
# id/secret/refresh token, then authenticate SMTP with AUTH XOAUTH2.
# ==========================================================
def get_gmail_access_token():
    """Exchange the long-lived refresh token for a ~1-hour access token.

    Called once per run, immediately before sending, so the token can't go
    stale while the (slow) Monogoto report is being generated."""
    resp = requests.post(
        GMAIL_TOKEN_URL,
        data={
            "client_id": GMAIL_CLIENT_ID,
            "client_secret": GMAIL_CLIENT_SECRET,
            "refresh_token": GMAIL_REFRESH_TOKEN,
            "grant_type": "refresh_token",
        },
        timeout=30,
    )
    try:
        data = resp.json()
    except ValueError:
        raise RuntimeError(f"Gmail OAuth token endpoint returned non-JSON "
                           f"({resp.status_code}): {resp.text[:300]}")

    if resp.status_code != 200 or "access_token" not in data:
        err = data.get("error")
        hint = ""
        if err == "invalid_grant":
            hint = (" — GMAIL_REFRESH_TOKEN is revoked or expired. This happens when the "
                    "sending account's password changed, the token was unused for 6 months, "
                    "or the OAuth consent screen is still in 'Testing' (those tokens die "
                    "after 7 days). Re-run get_gmail_refresh_token.py and update the secret.")
        elif err in ("invalid_client", "unauthorized_client"):
            hint = (" — GMAIL_CLIENT_ID / GMAIL_CLIENT_SECRET don't match, or don't match the "
                    "client that issued the refresh token. All three secrets must come from "
                    "the same OAuth client.")
        raise RuntimeError(f"Gmail OAuth failed | status={resp.status_code} | "
                           f"response={data}{hint}")

    scope = data.get("scope", "")
    if scope and "https://mail.google.com/" not in scope:
        print(f"⚠️ Granted Gmail scope is '{scope}' — SMTP XOAUTH2 requires the full "
              f"https://mail.google.com/ scope. gmail.send alone will be rejected with 535.")
    print(f"✅ Gmail access token acquired (expires in {data.get('expires_in', '?')}s)")
    return data["access_token"]


def _xoauth2_string(user, access_token):
    """Google's SASL XOAUTH2 payload: user=<addr>^Aauth=Bearer <tok>^A^A
    where ^A is 0x01. smtplib base64-encodes what the auth callback returns,
    so this returns the raw (unencoded) string."""
    return f"user={user}\x01auth=Bearer {access_token}\x01\x01"


# ==========================================================
# EMAIL — same shape as the 1NCE/Telenor scripts' send_email(), but the
# login step is OAuth2 rather than an app password.
# ==========================================================
def send_email(overconsumption_count, unmapped_count, excel_buffer):
    msg = EmailMessage()
    msg["Subject"] = "Monthly Monogoto SIM Usage Audit – Overconsumption Report"
    msg["From"] = EMAIL_SENDER
    msg["To"] = ", ".join(EMAIL_TO)
    msg.set_content(f"""
Hello Team,

Please find the monthly Monogoto SIM usage audit report attached.

Summary:
- Overconsumption SIMs: {overconsumption_count}
- ICCIDs without active customer in Zenduit Devices (analytics): {unmapped_count}

Regards,
Nandhiv
""")
    msg.add_attachment(
        excel_buffer.read(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="monogoto_overconsumption_report.xlsx",
    )

    access_token = get_gmail_access_token()
    auth_string = _xoauth2_string(EMAIL_SENDER, access_token)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=60) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()          # re-issued after STARTTLS so XOAUTH2 is advertised
        try:
            # initial_response_ok=True (the default) means smtplib sends the
            # payload with the AUTH verb and base64-encodes it for us.
            server.auth("XOAUTH2", lambda challenge=None: auth_string)
        except smtplib.SMTPAuthenticationError as e:
            raise RuntimeError(
                f"Gmail rejected the OAuth2 token: {e}. The token minted fine, so the "
                f"client id/secret/refresh token are consistent — the usual causes are "
                f"(a) the refresh token was consented by a DIFFERENT account than "
                f"GMAIL_USERNAME={EMAIL_SENDER}, or (b) the scope isn't the full "
                f"https://mail.google.com/, or (c) a Workspace admin has SMTP/IMAP "
                f"access disabled for this mailbox."
            ) from e
        server.send_message(msg)

    print("📧 Email sent via Gmail OAuth2 (Excel attached from memory)")


# ==========================================================
# MAIN
# ==========================================================
async def main():
    print("🔹 Getting Monogoto token...")
    token, customer_id = await get_monogoto_token()

    print("🔹 Fetching all Monogoto things (SIMs)... "
          "(API caps pages at 50, so this is sequential and can take a while)")
    df_base = await fetch_all_things(token, customer_id)

    today = datetime.now(timezone.utc).date()
    start_dt = today - timedelta(days=31)
    end_dt = today - timedelta(days=1)

    def zoho_analytics_chain():
        try:
            za_token = get_zoho_analytics_token()
            d_devices = fetch_zenduit_devices_analytics(za_token)
            d_account_name_lookup = fetch_account_name_lookup(za_token)
        except Exception as e:
            print(f"⚠️ Zoho Analytics lookups unavailable this run: {e}")
            d_devices = pd.DataFrame(columns=_ZENDUIT_DEVICE_COLUMNS)
            d_account_name_lookup = pd.DataFrame(columns=["AccountId", "Analytics_Customer_Name"])
        return d_devices, d_account_name_lookup

    print(f"🔹 Fetching Monogoto usage ({start_dt} to {end_dt}) and Zoho Analytics data concurrently...")
    df_usage, (df_z_devices, df_account_name_lookup) = await asyncio.gather(
        fetch_usage_report(token, customer_id, start_dt, end_dt),
        asyncio.to_thread(zoho_analytics_chain),
    )
    print(f"✅ Zenduit devices (analytics) fetched: {len(df_z_devices)}")

    # Join usage on ICCID where the report gave us one, else fall back to IMSI.
    if "ICCID" in df_usage.columns and df_usage["ICCID"].notna().any():
        df_base = df_base.merge(
            df_usage[["ICCID", "Monogoto_MB_Usage"]].dropna(subset=["ICCID"]), on="ICCID", how="left"
        )
    elif "IMSI" in df_usage.columns and df_usage["IMSI"].notna().any():
        df_base = df_base.merge(
            df_usage[["IMSI", "Monogoto_MB_Usage"]].dropna(subset=["IMSI"]), on="IMSI", how="left"
        )
    else:
        df_base["Monogoto_MB_Usage"] = 0

    # Step 1: Monogoto (ICCID) <-> Zoho Analytics "Zenduit Devices" view.
    df_base = df_base.merge(df_z_devices, on="ICCID", how="left")
    df_base["Zenduit_Usage_MB"] = _numeric_fillna(df_base["Zenduit_Usage_MB"])

    # Step 2: AccountId -> Account Name via the Zoho Analytics Accounts table.
    df_base = df_base.merge(df_account_name_lookup, on="AccountId", how="left")
    print(f"✅ Account Name lookup merged: "
          f"{df_base['Analytics_Customer_Name'].notna().sum()} / {len(df_base)} SIMs matched")

    df_base["Monogoto_MB_Usage"] = _numeric_fillna(df_base["Monogoto_MB_Usage"])
    df_base["Zenduit_Data_Plan"] = _numeric_fillna(df_base["Zenduit_Data_Plan"])

    df_base["Consumption"] = ""
    df_base.loc[
        df_base["Monogoto_MB_Usage"] > df_base["Zenduit_Data_Plan"],
        "Consumption"
    ] = "Overconsumption"

    device_iccids = set(df_z_devices["ICCID"].dropna().astype(str).str.strip())
    df_no_customer_in_zenduone = df_base[
        df_base["ICCID"].notna() &
        ~df_base["ICCID"].astype(str).str.strip().isin(device_iccids)
    ]

    df_account_summary = (
        df_base[df_base["AccountId"].notna()]
        .groupby(["AccountId", "Analytics_Customer_Name"], dropna=False)
        .agg(
            iccid_count=("ICCID", "nunique"),
            monogoto_usage_mb=("Monogoto_MB_Usage", "sum"),
            zenduit_usage_mb=("Zenduit_Usage_MB", "sum"),
        )
        .reset_index()
    )

    print("🔹 Writing Excel file...")
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        df_base.to_excel(writer, sheet_name="base_combined", index=False)
        df_no_customer_in_zenduone.to_excel(writer, sheet_name="no_customer_in_zenduone", index=False)
        df_account_summary.to_excel(writer, sheet_name="account_usage_summary", index=False)
    excel_buffer.seek(0)
    print("✅ DONE → Excel generated")

    # Persist the workbook to disk when asked. CI sets SAVE_REPORT_PATH and
    # uploads the file as a run artifact, so a failure in the email step
    # (expired token, Gmail outage) no longer throws away a run that already
    # did all the slow Monogoto/Zoho work — you can download the report from
    # the Actions run and send it by hand.
    save_path = os.getenv("SAVE_REPORT_PATH")
    if save_path:
        os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
        with open(save_path, "wb") as fh:
            fh.write(excel_buffer.getvalue())
        print(f"💾 Report also written to {save_path} ({len(excel_buffer.getvalue())} bytes)")

    overconsumption_count = df_base[df_base["Consumption"] == "Overconsumption"].shape[0]
    unmapped_iccid_count = df_no_customer_in_zenduone.shape[0]
    print(f"📊 Overconsumption count: {overconsumption_count}")
    print(f"📊 ICCIDs without active customer: {unmapped_iccid_count}")

    print("🔹 Sending email...")
    send_email(
        overconsumption_count=overconsumption_count,
        unmapped_count=unmapped_iccid_count,
        excel_buffer=excel_buffer,
    )


if __name__ == "__main__":
    asyncio.run(main())
