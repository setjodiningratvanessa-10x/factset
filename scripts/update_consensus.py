"""
Weekly FactSet Consensus updater for TXG.
Fetches consensus estimates from FactSet API and writes to Google Sheets.
"""

import json
import os
import time
import datetime
import requests
import gspread
from google.oauth2.service_account import Credentials
from jose import jwt as jose_jwt

# ── Configuration ────────────────────────────────────────────────────────────

TICKER = "TXG"
SPREADSHEET_ID = "1svUgGTr8DVUGIV3vwKO8EMJTPGklmnz19eUdJp_HTJA"
SHEET_NAME = "Weekly Consensus"

FACTSET_TOKEN_URL = "https://auth.factset.com/as/token.oauth2"
FACTSET_FORMULA_URL = "https://api.factset.com/formula-api/v1/cross-sectional"

# Quarters and annual periods to fetch
QTR_PERIODS = ["2026/1F", "2026/2F", "2026/3F", "2026/4F"]
ANN_PERIODS = [2026, 2027]

# Rows: (label, fds_code, is_product_line)
# fds_code None means it's a calculated row (formula in sheet)
METRICS = [
    # Product segment estimates
    ("Instrument Revenue", None, False),           # section header
    ("Chromium",          "PRODLINE_SALES_4", True),
    ("Spatial",           "PRODLINE_SALES_5", True),
    ("Visium",            None,               True),   # always na
    ("Xenium",            None,               True),   # always na
    ("Instrument Revenue","PRODLINE_SALES_1", True),
    ("Consumables Revenue", None, False),          # section header
    ("Chromium",          "PRODLINE_SALES_6", True),
    ("Spatial",           "PRODLINE_SALES_7", True),
    ("Visium",            None,               True),
    ("Xenium",            None,               True),
    ("Consumables Revenue","PRODLINE_SALES_2", True),
    ("Services Revenue",  "PRODLINE_SALES_3", True),
    ("Total Revenue",     "SALES",            True),
    # P&L
    ("COGS",              "COS",              False),
    ("Gross Profit",      "GROSS_INC",        False),
    ("Gross Margin %",    None,               False),  # calculated
    ("R&D",               "RD_EXP",           False),
    ("SG&A",              "SGA",              False),
    ("Total Opex",        None,               False),  # calculated
    ("EBIT",              "EBIT",             False),
    ("Net Income",        "NET_INC",          False),
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
    headers = {"kid": jwk["kid"]}
    private_key = {k: v for k, v in jwk.items() if k != "use"}
    token = jose_jwt.encode(payload, private_key, algorithm="RS256", headers=headers)

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


# ── FactSet Formula API ───────────────────────────────────────────────────────

def fetch_estimate(access_token: str, fds_code: str, period: str, as_of_date: str, freq: str = "QTR_ROLL") -> float | str:
    """Call FE_ESTIMATE via the FactSet Formula API."""
    formula = f'FE_ESTIMATE("{fds_code}","MEAN","{freq}","{period}","{as_of_date}")'
    payload = {
        "data": {
            "ids": [TICKER],
            "formulas": [formula],
        }
    }
    resp = requests.post(
        FACTSET_FORMULA_URL,
        json=payload,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        timeout=30,
    )
    if resp.status_code != 200:
        return "na"
    try:
        result = resp.json()
        value = result["data"][0]["result"][0]
        return round(float(value), 3) if value is not None else "na"
    except (KeyError, IndexError, TypeError, ValueError):
        return "na"


def fetch_all_estimates(access_token: str, as_of_date: str) -> dict:
    """Fetch all estimates for current week and prior week."""
    prior_date = (datetime.datetime.strptime(as_of_date, "%Y-%m-%d") - datetime.timedelta(days=7)).strftime("%Y-%m-%d")

    results = {}
    codes = {m[1] for m in METRICS if m[1] and m[1] not in (None,)}

    for code in codes:
        results[code] = {}
        for period in QTR_PERIODS:
            for date_key, date_val in [("current", as_of_date), ("prior", prior_date)]:
                val = fetch_estimate(access_token, code, period, date_val, "QTR_ROLL")
                results[code][f"{date_key}_{period}"] = val
        for period in ANN_PERIODS:
            for date_key, date_val in [("current", as_of_date), ("prior", prior_date)]:
                val = fetch_estimate(access_token, code, str(period), date_val, "ANN")
                results[code][f"{date_key}_{period}"] = val

    return results, prior_date


# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_sheet():
    creds_json = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(
        creds_json,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = spreadsheet.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=SHEET_NAME, rows=50, cols=30)
    return ws


def build_sheet_data(as_of_date: str, prior_date: str, estimates: dict) -> list[list]:
    """Build the 2D array to write to the sheet."""
    def v(code, period_key):
        if code is None:
            return "na"
        return estimates.get(code, {}).get(period_key, "na")

    def diff(code, period_key_curr, period_key_prior):
        curr = v(code, period_key_curr)
        prior = v(code, period_key_prior)
        if curr == "na" or prior == "na":
            return "na"
        try:
            return round(float(curr) - float(prior), 3)
        except (TypeError, ValueError):
            return "na"

    qtr_labels = ["2026 Q1", "2026 Q2", "2026 Q3", "2026 Q4", "FY 2026", "FY 2027"]
    all_periods_curr = [f"current_{p}" for p in QTR_PERIODS] + [f"current_{p}" for p in ANN_PERIODS]
    all_periods_prior = [f"prior_{p}" for p in QTR_PERIODS] + [f"prior_{p}" for p in ANN_PERIODS]

    rows = []
    rows.append(["Weekly FactSet Consensus"] + [""] * 20)
    rows.append([""] * 21)
    rows.append([""] * 21)
    rows.append([""] * 21)

    # Row 5: dates
    rows.append(["", "Date", "", as_of_date, "", "", "", "", "", prior_date] + [""] * 11)

    # Row 6: ticker
    rows.append(["", "Ticker", "", TICKER] + [""] * 17)

    # Row 7: timing headers
    rows.append(["", "Timing", ""] + QTR_PERIODS + [str(p) for p in ANN_PERIODS] +
                [""] + QTR_PERIODS + [str(p) for p in ANN_PERIODS] + [""] * 3)

    rows.append([""] * 21)

    # Header row
    rows.append(["", "", "", "FactSet Consensus"] + [""] * 8 + [""] +
                ["Change since Last Week"] + [""] * 8)

    rows.append(["", "", "",
                 f"As of {as_of_date}"] + [""] * 5 +
                [""] +
                [f"As of {prior_date}"] + [""] * 11)

    # Column headers
    rows.append(["", "", ""] + qtr_labels + [""] + qtr_labels + [""] * 3)

    # Data rows
    metric_map = {
        "Instrument Revenue (header)": None,
        "Chromium (instr)": "PRODLINE_SALES_4",
        "Spatial (instr)": "PRODLINE_SALES_5",
        "Visium": None,
        "Xenium": None,
        "Instrument Revenue": "PRODLINE_SALES_1",
        "Consumables Revenue (header)": None,
        "Chromium (cons)": "PRODLINE_SALES_6",
        "Spatial (cons)": "PRODLINE_SALES_7",
        "Consumables Revenue": "PRODLINE_SALES_2",
        "Services Revenue": "PRODLINE_SALES_3",
        "Total Revenue": "SALES",
    }

    data_rows = [
        ("Instrument Revenue", None),
        ("  Chromium", "PRODLINE_SALES_4"),
        ("  Spatial", "PRODLINE_SALES_5"),
        ("  Visium", None),
        ("  Xenium", None),
        ("Instrument Revenue (Total)", "PRODLINE_SALES_1"),
        ("Consumables Revenue", None),
        ("  Chromium", "PRODLINE_SALES_6"),
        ("  Spatial", "PRODLINE_SALES_7"),
        ("  Visium", None),
        ("  Xenium", None),
        ("Consumables Revenue (Total)", "PRODLINE_SALES_2"),
        ("Services Revenue", "PRODLINE_SALES_3"),
        ("Total Revenue", "SALES"),
        ["SEPARATOR"],
        ("COGS", "COS"),
        ("Gross Profit", "GROSS_INC"),
        ("Gross Margin %", None),  # calculated below
        ["SEPARATOR"],
        ("R&D", "RD_EXP"),
        ("SG&A", "SGA"),
        ("Total Opex", None),      # calculated below
        ("EBIT", "EBIT"),
        ["SEPARATOR"],
        ("Net Income", "NET_INC"),
    ]

    rev_curr = None
    gp_curr = None

    for item in data_rows:
        if item == ["SEPARATOR"]:
            rows.append([""] * 21)
            continue

        label, code = item
        curr_vals = [v(code, p) for p in all_periods_curr]
        prior_vals = [v(code, p) for p in all_periods_prior]

        # Special calculated rows
        if label == "Gross Margin %" and rev_curr and gp_curr:
            curr_vals = []
            prior_vals = []
            for i in range(6):
                try:
                    curr_vals.append(round(gp_curr[i] / rev_curr[i], 4) if isinstance(gp_curr[i], float) and isinstance(rev_curr[i], float) else "na")
                except (TypeError, ZeroDivisionError):
                    curr_vals.append("na")
            prior_vals = ["na"] * 6
        elif label == "Total Opex":
            curr_vals = ["na"] * 6
            prior_vals = ["na"] * 6

        diff_vals = []
        for c, p in zip(curr_vals, prior_vals):
            if c == "na" or p == "na":
                diff_vals.append("na")
            else:
                try:
                    diff_vals.append(round(float(c) - float(p), 3))
                except (TypeError, ValueError):
                    diff_vals.append("na")

        row = ["", label, ""] + curr_vals + [""] + prior_vals + [""] + diff_vals
        rows.append(row)

        # Save for GM% calc
        if label == "Total Revenue":
            rev_curr = curr_vals
        if label == "Gross Profit":
            gp_curr = curr_vals

    rows.append([""] * 21)
    rows.append(["", "Note: Individual product and platform consensus figures take average of figures from available analysts (not all analysts have estimates by product and platform); numbers may not tie to total revenue."] + [""] * 19)

    # Pad all rows to same width
    max_width = max(len(r) for r in rows)
    rows = [r + [""] * (max_width - len(r)) for r in rows]

    return rows


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    # Use the most recent Monday
    today = datetime.date.today()
    days_since_monday = today.weekday()
    this_monday = today - datetime.timedelta(days=days_since_monday)
    as_of_date = this_monday.strftime("%Y-%m-%d")

    print(f"Fetching consensus as of {as_of_date} for {TICKER}...")

    access_token = get_factset_token()
    print("FactSet token obtained.")

    estimates, prior_date = fetch_all_estimates(access_token, as_of_date)
    print(f"Estimates fetched. Prior week: {prior_date}")

    ws = get_sheet()
    data = build_sheet_data(as_of_date, prior_date, estimates)

    ws.clear()
    ws.update(data, value_input_option="USER_ENTERED")
    print(f"Google Sheet updated: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


if __name__ == "__main__":
    main()
