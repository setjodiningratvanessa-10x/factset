"""
Weekly FactSet Consensus updater for TXG.
Fetches consensus estimates from FactSet API and writes to Google Sheets
via a Google Apps Script web endpoint (no Google Cloud account needed).
"""

import json
import os
import time
import datetime
import requests
from jose import jwt as jose_jwt

# ── Configuration ────────────────────────────────────────────────────────────

TICKER = "TXG-US"
SHEET_NAME = "Weekly Consensus"

FACTSET_TOKEN_URL = "https://auth.factset.com/as/token.oauth2"
FACTSET_FORMULA_URL = "https://api.factset.com/formula-api/v1/cross-sectional"

QTR_PERIODS = ["2026/1F", "2026/2F", "2026/3F", "2026/4F"]
ANN_PERIODS = [2026, 2027]

# (display label, FDS estimate code)
# None code = calculated or always "na"
DATA_ROWS = [
    # Instrument Revenue section
    ("Instrument Revenue",        None),               # section header
    ("  Chromium",                "PRODLINE_SALES_4"),
    ("  Spatial",                 "PRODLINE_SALES_5"),
    ("  Visium",                  None),
    ("  Xenium",                  None),
    ("Instrument Revenue (Total)","PRODLINE_SALES_1"),
    # Consumables Revenue section
    ("Consumables Revenue",       None),               # section header
    ("  Chromium",                "PRODLINE_SALES_6"),
    ("  Spatial",                 "PRODLINE_SALES_7"),
    ("  Visium",                  None),
    ("  Xenium",                  None),
    ("Consumables Revenue (Total)","PRODLINE_SALES_2"),
    ("Services Revenue",          "PRODLINE_SALES_3"),
    ("Total Revenue",             "SALES"),
    # P&L section
    (None, None),                                      # blank separator
    ("COGS",                      "COS"),
    ("Gross Profit",              "GROSS_INC"),
    ("Gross Margin %",            None),               # calculated
    (None, None),
    ("R&D",                       "RD_EXP"),
    ("SG&A",                      "SGA"),
    ("Total Opex",                None),               # calculated
    ("EBIT",                      "EBIT"),
    (None, None),
    ("Net Income",                "NET_INC"),
]


# ── FactSet Auth ─────────────────────────────────────────────────────────────

def get_factset_token() -> str:
    client_id = os.environ["FACTSET_CLIENT_ID"]
    jwk = json.loads(os.environ["FACTSET_JWK"])

    now = int(time.time())
    payload = {
        "iss": client_id,
        "sub": client_id,
        "aud": FACTSET_TOKEN_URL,
        "iat": now,
        "exp": now + 300,
        "jti": f"{client_id}-{now}",
    }
    token = jose_jwt.encode(
        payload,
        {k: v for k, v in jwk.items()},
        algorithm="RS256",
        headers={"kid": jwk["kid"]},
    )

    resp = requests.post(
        FACTSET_TOKEN_URL,
        data={
            "grant_type": "client_credentials",
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": token,
            "client_id": client_id,
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


# ── FactSet Estimates ─────────────────────────────────────────────────────────

def fetch_estimate(access_token: str, fds_code: str, period: str, as_of_date: str, freq: str) -> float | str:
    # FQL syntax: keywords unquoted, strings in single quotes
    formula = f"FE_ESTIMATE({fds_code},MEAN,{freq},'{period}','{as_of_date}')"
    resp = requests.post(
        FACTSET_FORMULA_URL,
        json={"data": {"ids": [TICKER], "formulas": [formula]}},
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        timeout=30,
    )
    if resp.status_code != 200:
        print(f"WARN {fds_code}/{period} HTTP {resp.status_code}: {resp.text[:300]}", flush=True)
        return "na"
    try:
        body = resp.json()
        value = body["data"][0]["result"][0]
        return round(float(value), 3) if value is not None else "na"
    except (KeyError, IndexError, TypeError, ValueError) as e:
        print(f"WARN parse error {fds_code}/{period}: {e} | {str(body)[:300]}", flush=True)
        return "na"


def fetch_all_estimates(access_token: str, as_of_date: str, prior_date: str) -> dict:
    codes = {row[1] for row in DATA_ROWS if row[1] is not None}
    estimates = {}

    for code in codes:
        estimates[code] = {}
        for period in QTR_PERIODS:
            estimates[code][f"curr_{period}"] = fetch_estimate(access_token, code, period, as_of_date, "QTR_ROLL")
            estimates[code][f"prior_{period}"] = fetch_estimate(access_token, code, period, prior_date, "QTR_ROLL")
        for period in ANN_PERIODS:
            estimates[code][f"curr_{period}"] = fetch_estimate(access_token, code, str(period), as_of_date, "ANN")
            estimates[code][f"prior_{period}"] = fetch_estimate(access_token, code, str(period), prior_date, "ANN")

    return estimates


# ── Sheet Data Builder ────────────────────────────────────────────────────────

def get_vals(estimates: dict, code: str | None, prefix: str) -> list:
    if code is None:
        return ["na"] * 6
    row = estimates.get(code, {})
    return (
        [row.get(f"{prefix}_{p}", "na") for p in QTR_PERIODS] +
        [row.get(f"{prefix}_{p}", "na") for p in ANN_PERIODS]
    )


def calc_diff(curr: list, prior: list) -> list:
    result = []
    for c, p in zip(curr, prior):
        try:
            result.append(round(float(c) - float(p), 3))
        except (TypeError, ValueError):
            result.append("na")
    return result


def build_sheet_rows(as_of_date: str, prior_date: str, estimates: dict) -> list[list]:
    qtr_labels = ["2026 Q1", "2026 Q2", "2026 Q3", "2026 Q4", "FY 2026", "FY 2027"]

    rows = [
        ["Weekly FactSet Consensus"] + [""] * 21,
        [""] * 22,
        [""] * 22,
        [""] * 22,
        # Row 5: dates
        ["", "Date", "", as_of_date, "", "", "", "", "", prior_date] + [""] * 12,
        # Row 6: ticker
        ["", "Ticker for FactSet Formulas", "", TICKER] + [""] * 18,
        # Row 7: timing periods
        ["", "Timing for FactSet Formulas", ""] +
        QTR_PERIODS + [str(p) for p in ANN_PERIODS] + [""] +
        QTR_PERIODS + [str(p) for p in ANN_PERIODS] + [""] * 3,
        [""] * 22,
        # Section headers
        ["", "", "", "FactSet Consensus"] + [""] * 7 + ["", "Change since Last Week"] + [""] * 10,
        ["", "", "", f"As of {as_of_date}"] + [""] * 5 + ["", f"As of {prior_date}"] + [""] * 12,
        # Column labels
        ["", "", ""] + qtr_labels + [""] + qtr_labels + [""] + qtr_labels,
    ]

    saved = {}

    for label, code in DATA_ROWS:
        if label is None:
            rows.append([""] * 22)
            continue

        curr = get_vals(estimates, code, "curr")
        prior = get_vals(estimates, code, "prior")

        # Calculated rows
        if label == "Gross Margin %":
            rev = saved.get("Total Revenue", ["na"] * 6)
            gp = saved.get("Gross Profit", ["na"] * 6)
            curr = []
            for r, g in zip(rev, gp):
                try:
                    curr.append(round(float(g) / float(r), 4))
                except (TypeError, ValueError, ZeroDivisionError):
                    curr.append("na")
            prior = ["na"] * 6

        elif label == "Total Opex":
            gp = saved.get("Gross Profit", ["na"] * 6)
            ebit = saved.get("EBIT", ["na"] * 6)
            curr = []
            for g, e in zip(gp, ebit):
                try:
                    curr.append(round(float(g) - float(e), 3))
                except (TypeError, ValueError):
                    curr.append("na")
            prior = ["na"] * 6

        diff = calc_diff(curr, prior)
        saved[label.strip()] = curr

        rows.append(["", label, ""] + curr + [""] + prior + [""] + diff)

    rows.append([""] * 22)
    rows.append([
        "",
        "Note: Individual product and platform consensus figures take average of figures from available analysts "
        "(not all analysts have estimates by product and platform); numbers may not tie to total revenue.",
    ] + [""] * 20)

    # Normalize all rows to the same width
    width = max(len(r) for r in rows)
    return [r + [""] * (width - len(r)) for r in rows]


# ── Google Sheets via Apps Script ─────────────────────────────────────────────

def push_to_sheet(rows: list[list], sheet_name: str):
    web_app_url = os.environ["GOOGLE_APPS_SCRIPT_URL"]
    secret_token = os.environ["GOOGLE_APPS_SCRIPT_TOKEN"]

    resp = requests.post(
        web_app_url,
        params={"token": secret_token},
        json={"sheet_name": sheet_name, "rows": rows},
        timeout=60,
    )
    resp.raise_for_status()
    if resp.text.strip() != "OK":
        raise RuntimeError(f"Apps Script returned: {resp.text}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    today = datetime.date.today()
    this_monday = today - datetime.timedelta(days=today.weekday())
    as_of_date = this_monday.strftime("%Y-%m-%d")
    prior_date = (this_monday - datetime.timedelta(days=7)).strftime("%Y-%m-%d")

    print(f"Fetching TXG consensus as of {as_of_date} (prior: {prior_date})")

    access_token = get_factset_token()
    print("FactSet token obtained.", flush=True)

    # Diagnostic: test Formula API and Estimates API to see which one works
    import sys

    # Test 1: Formula API
    try:
        test_formula = f"FE_ESTIMATE(SALES,MEAN,ANN,'2026','{as_of_date}')"
        r1 = requests.post(
            FACTSET_FORMULA_URL,
            json={"data": {"ids": [TICKER], "formulas": [test_formula]}},
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            timeout=30,
        )
        sys.stderr.write(f"FORMULA_API status={r1.status_code}\n")
        sys.stderr.write(f"FORMULA_API body={r1.text[:600]}\n")
        sys.stderr.flush()
    except Exception as ex:
        sys.stderr.write(f"FORMULA_API exception: {ex}\n")
        sys.stderr.flush()

    # Test 2: Dedicated Estimates API
    try:
        r2 = requests.post(
            "https://api.factset.com/content/factset-estimates/v2/consensus-estimates",
            json={
                "ids": [TICKER],
                "metrics": ["SALES", "EPS"],
                "periodicity": "ANNUAL",
                "fiscalPeriodStart": "2026",
                "fiscalPeriodEnd": "2027",
                "currency": "USD",
            },
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            timeout=30,
        )
        sys.stderr.write(f"ESTIMATES_API status={r2.status_code}\n")
        sys.stderr.write(f"ESTIMATES_API body={r2.text[:600]}\n")
        sys.stderr.flush()
    except Exception as ex:
        sys.stderr.write(f"ESTIMATES_API exception: {ex}\n")
        sys.stderr.flush()

    estimates = fetch_all_estimates(access_token, as_of_date, prior_date)
    print(f"Fetched estimates for {len(estimates)} metrics.", flush=True)

    rows = build_sheet_rows(as_of_date, prior_date, estimates)
    push_to_sheet(rows, SHEET_NAME)
    print("Google Sheet updated successfully.")


if __name__ == "__main__":
    main()
