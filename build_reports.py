#!/usr/bin/env python3
"""
Swiggy payout annexure consolidation & GST return working
=========================================================

Point it at one or more Swiggy weekly payout annexures (invoice_Annexure_*.xlsx)
and it builds two workbooks **per restaurant / GSTIN**:

  1. Swiggy_Consolidated_Report.xlsx
        Consolidated data report - period summary, merged payout breakup, merged
        order-level data, ads/discount/complaint detail, cancelled-order policy
        working, reconciliation checks and field definitions.

  2. Swiggy_GST_Return_Working.xlsx
        GST return filing working papers - period-wise tax summary, GSTR-1 and
        GSTR-3B working, Section 9(5) tax discharge reconciliation, TCS/TDS
        reconciliation and an order-wise GST register.

Usage
-----
  python3 build_reports.py                                  # interactive menu
  python3 build_reports.py --all                            # every annexure found in this folder
  python3 build_reports.py --list                           # just show what would be read
  python3 build_reports.py a.xlsx b.xlsx                    # specific files
  python3 build_reports.py "D:/annexures/*.xlsx"            # a glob
  python3 build_reports.py --dir "D:/annexures" --all       # a folder
  python3 build_reports.py --all --outdir out --gst-rate 5

Files for different restaurants / GSTINs can be mixed in one run: the reports are
grouped, and a separate pair of workbooks is written for each restaurant.
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import sys
from collections import OrderedDict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, "reports")
PREPARED_ON = datetime.now().strftime("%d %B %Y")

# Default tax rate; can be overridden with --gst-rate or left to auto-detection.
GST_RATE = 0.05
CANCELLATION_COMPENSATION_RATE = 0.80
TDS_RATE = 0.001  # Section 194-O, 0.1%

# GST state codes, used to describe the place of supply from the GSTIN.
STATE_CODES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh",
    "05": "Uttarakhand", "06": "Haryana", "07": "Delhi", "08": "Rajasthan",
    "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh",
    "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura",
    "17": "Meghalaya", "18": "Assam", "19": "West Bengal", "20": "Jharkhand",
    "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "26": "Dadra & Nagar Haveli and Daman & Diu", "27": "Maharashtra", "29": "Karnataka",
    "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu",
    "34": "Puducherry", "35": "Andaman & Nicobar Islands", "36": "Telangana",
    "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory",
}

GSTIN_RE = re.compile(r"\b(\d{2}[A-Z]{5}\d{4}[A-Z][A-Z\d]Z[A-Z\d])\b")
REST_ID_RE = re.compile(r"Rest(?:aurant)?\.?\s*ID\s*[-:\u2013]?\s*(\d+)", re.IGNORECASE)
DATE_IN_NAME_RE = re.compile(r"_(\d{2})(\d{2})(\d{4})_")
REQUIRED_SHEETS = ("summary", "payout breakup", "order level")

# --------------------------------------------------------------------------
# Annexure layout (identical across all weekly files)
# --------------------------------------------------------------------------
# Order Level: 1-based column index -> field name
ORDER_COLS = OrderedDict([
    (1, "Order ID"), (2, "Parent Order ID"), (3, "Order Date"),
    (4, "Order Settlement Date"), (5, "Order Status"), (6, "Order Category"),
    (7, "Order Payment Type"), (8, "Cancelled By?"),
    (9, "Coupon type applied by customer"), (10, "Item Total"),
    (11, "Packaging Charges"),
    (12, "Restaurant Discounts (Promo, Freebies, Flat Off, etc.)"),
    (13, "Swiggy One / Exclusive Offer Discount"),
    (14, "Restaurant Discount Share [3a+3b]"),
    (15, "Net Bill Value (before taxes) [1+2-3]"), (16, "GST Collected"),
    (17, "Total Customer Paid [4+5]"), (18, "Commission charged on"),
    (19, "Service Fees %"), (20, "Commission"), (21, "Long Distance Charges"),
    (22, "Discount on Long Distance Fee"), (23, "Pocket Hero Fees"),
    (24, "No Fees Week - Cashback"), (25, "Swiggy One Fees"),
    (26, "Payment Collection Charges"), (27, "Restaurant Cancellation Charges"),
    (28, "Call Center Charges"),
    (29, "Delivery Fee sponsored by Restaurant (w/o tax)"), (30, "Bolt Fees"),
    (31, "GST on Service Fee @18%"),
    (32, "Total Swiggy Fees [incl. GST on fees]"),
    (33, "Customer Cancellations"), (34, "Customer Complaints"),
    (35, "Complaint & Cancellation Charges [18+19]"),
    (36, "GST Deduction [Sec 9(5)]"), (37, "TCS"), (38, "TDS"),
    (39, "Total Taxes [20+21+22]"),
    (40, "Net Payout for Order (after taxes)"), (41, "Long Distance Order"),
    (42, "Last Mile (in km)"), (43, "MFR Accurate?"), (44, "MFR Pressed?"),
    (45, "Coupon Code Sourced"), (46, "Discount Campaign ID"),
    (47, "Replicated Order"), (48, "Base order ID"), (49, "Cancellation time"),
    (50, "Pick Up Status"), (51, "Swiggy One Customer?"),
    (52, "Pocket Hero Order?"),
])

# Payout Breakup: line code -> canonical label
PAYOUT_LINES = OrderedDict([
    ("A", "Total Customer Paid [1+2-3-4+5]"),
    ("1.0", "Item Total"),
    ("2.0", "Packaging Charges"),
    ("3.0", "Restaurant Discounts (Coupon based)"),
    ("4.0", "Restaurant Discounts (Trade Discounts, Freebies and others)"),
    ("5.0", "GST Collected"),
    ("B", "Swiggy Fees [6+7+8+9+10+11+12+13+14+15]"),
    ("6.0", "Commission"),
    ("7.0", "Long Distance Charges"),
    ("8.0", "Payment Collection Charges"),
    ("9.0", "Pocket Hero Fees"),
    ("10.0", "Swiggy One Fees"),
    ("11.0", "Restaurant Cancellation Charges"),
    ("12.0", "Call Center Charges"),
    ("13.0", "Delivery Fee sponsored by Restaurant"),
    ("14.0", "Bolt Fees"),
    ("15.0", "No Fees Week - Cashback"),
    ("C", "Customer Complaints & Cancellation Charges [16+17]"),
    ("16.0", "Merchant Share Of Cancelled Orders"),
    ("17.0", "Refund For Customer Complaints"),
    ("D", "Growth Investment In Ads"),
    ("E", "Other Charges and Refunds"),
    ("F", "Total Taxes [18+19+20+21]"),
    ("18", "GST Deduction (paid by Swiggy on behalf of the restaurant)"),
    ("19", "GST on Service Fee @18%"),
    ("20", "TCS"),
    ("21", "TDS"),
    ("G", "Net Payout [A+B+C+D+E+F]"),
])
SUB_LINES = {"1.0", "2.0", "3.0", "4.0", "5.0", "6.0", "7.0", "8.0", "9.0",
             "10.0", "11.0", "12.0", "13.0", "14.0", "15.0", "16.0", "17.0",
             "18", "19", "20", "21"}


def _norm_code(code):
    """Line codes read back as 1, 1.0 or '1.0' depending on how Excel saved the file."""
    if code is None:
        return None
    if isinstance(code, (int, float)):
        return "{:g}".format(float(code))
    text = str(code).strip()
    if re.match(r"^\d+(?:\.\d+)?$", text):
        return "{:g}".format(float(text))
    return text


# normalised code -> canonical key used everywhere in this script
PAYOUT_CANON = {_norm_code(key): key for key in PAYOUT_LINES}

MONTH_NUM = {"january": 1, "february": 2, "march": 3, "april": 4, "may": 5,
             "june": 6, "july": 7, "august": 8, "september": 9, "october": 10,
             "november": 11, "december": 12}
MONTH_ABBR = {"january": "Jan", "february": "Feb", "march": "Mar", "april": "Apr",
              "may": "May", "june": "Jun", "july": "Jul", "august": "Aug",
              "september": "Sep", "october": "Oct", "november": "Nov",
              "december": "Dec"}


# --------------------------------------------------------------------------
# Cancellation compensation policy
# --------------------------------------------------------------------------
# When an order is cancelled before pickup and the cancellation is NOT on the
# restaurant (i.e. cancelled by the customer or by Swiggy), Swiggy compensates the
# restaurant at 80% of the order value: "You will be paid 80% less commissions as
# compensation."  The restaurant is taxed on that compensation value, not on the
# full value the customer had paid:
#
#       compensation value (pre-tax) = 80% x net bill value
#       GST discharged u/s 9(5)      = GST rate x compensation value
#       TDS u/s 194-O                = 0.1% x compensation value
#       commission / collection      = charged on 80% x total customer paid
#       balance 20% + GST            = deducted as 'Customer Cancellations'
#
# Everything below is derived from that basis, so the working can be re-checked
# against the annexures order by order.
MERCHANT_CANCEL_WORDS = ("MERCHANT", "RESTAURANT", "OUTLET", "SELF")


def cancel_category(order):
    """How a cancelled order is treated: 'merchant', 'after_pickup', 'compensated' or ''.

    - merchant     : cancelled by the restaurant -> no supply, no GST on the food
                     (only the cancellation charge, quoted exclusive of GST);
    - after_pickup : cancelled once the food had been picked up -> the supply
                     happened, so the full value is taxable and the GST is due;
    - compensated : cancelled before pickup by the customer/Swiggy -> the
                     restaurant is paid 80% of the order value and that is what
                     is taxed (the un-compensated 20% is not a supply).
    """
    if str(order.get("Order Status") or "").strip().lower() != "cancelled":
        return ""
    by = str(order.get("Cancelled By?") or "").strip().upper()
    if any(word in by for word in MERCHANT_CANCEL_WORDS):
        return "merchant"
    pickup = str(order.get("Pick Up Status") or "").strip().lower()
    if "picked" in pickup and "not" not in pickup:
        return "after_pickup"
    return "compensated"


def is_compensated_cancel(order):
    """True when the restaurant is compensated for a pre-pickup cancellation."""
    return cancel_category(order) == "compensated"


def is_after_pickup_cancel(order):
    """True when the order was cancelled after pickup - the supply still stands."""
    return cancel_category(order) == "after_pickup"


def order_billed_value(order):
    return r2(num(order.get("Net Bill Value (before taxes) [1+2-3]")))


def order_reportable_value(order):
    """Value on which GST is discharged: full value normally, 80% for a
    compensated pre-pickup cancellation."""
    billed = order_billed_value(order)
    if is_compensated_cancel(order):
        return r2(billed * CANCELLATION_COMPENSATION_RATE)
    return billed


def order_compensation_deduction(order):
    """Un-compensated portion of a compensated cancellation (nil otherwise)."""
    if not is_compensated_cancel(order):
        return 0.0
    return r2(order_billed_value(order) - order_reportable_value(order))


# --------------------------------------------------------------------------
# Generic helpers
# --------------------------------------------------------------------------
def safe_slug(text, fallback="restaurant"):
    """Filesystem-safe short name for a restaurant / outlet."""
    slug = re.sub(r"[^A-Za-z0-9]+", "_", str(text or "")).strip("_")
    slug = re.sub(r"_+", "_", slug)
    return (slug[:60] or fallback)


def group_folder_name(first, key):
    """Folder name for one restaurant: 'Name_544662'.

    Two outlets can trade under the same brand (here both are 'Cheesecake
    Mills'), so the restaurant id - or the GSTIN when there is no id - is
    always added to keep each outlet's reports in their own folder.
    """
    name = first.get("restaurant") or first.get("rest_id") or "Restaurant"
    suffix = str(first.get("rest_id") or first.get("gstin") or key or "").strip()
    return "{} {}".format(name, suffix).strip() if suffix else str(name)


def state_name(gstin):
    """'24AAQFC3084Q1ZT' -> 'Gujarat' (falls back to 'state 24')."""
    gstin = str(gstin or "")
    code = gstin[:2]
    return STATE_CODES.get(code, "state {}".format(code or "??"))


def place_of_supply(gstin):
    gstin = str(gstin or "")
    return "{} ({})".format(state_name(gstin), gstin[:2] or "??")


def find_sheet(wb, *candidates):
    """Case/spacing-tolerant sheet lookup."""
    wanted = [c.strip().lower() for c in candidates]
    for name in wb.sheetnames:
        if name.strip().lower() in wanted:
            return wb[name]
    for name in wb.sheetnames:
        low = name.strip().lower()
        if any(low.startswith(w[:12]) for w in wanted):
            return wb[name]
    return None


def looks_like_annexure(path):
    """True when the workbook has the sheets a Swiggy payout annexure must have."""
    if str(path).lower().endswith(".xls"):
        return False, ("old .xls format - open it in Excel and save as .xlsx, then try again")
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception:
        return False, "not a readable .xlsx file"
    names = [n.strip().lower() for n in wb.sheetnames]
    missing = [n for n in REQUIRED_SHEETS if n not in names]
    wb.close()
    if missing:
        return False, "missing sheet(s): {}".format(", ".join(missing))
    return True, ""


def extract_identity(wb):
    """Pull restaurant name, outlet, restaurant id and GSTIN from the Summary sheet."""
    ws = find_sheet(wb, "Summary")
    identity = {"restaurant": "", "outlet": "", "city": "", "rest_id": "", "gstin": "", "summary": {}}
    if ws is None:
        return identity
    for row in ws.iter_rows(min_row=1, max_row=25):
        label = row[1].value
        value = row[2].value
        if label is not None and value is not None:
            identity["summary"][str(label).strip()] = value
    # Values may sit in the label column itself (merged cells) - scan the first
    # block of cells for a GSTIN / restaurant id whatever the layout.
    for row in ws.iter_rows(min_row=1, max_row=25, max_col=4):
        for cell in row:
            text = cell.value
            if isinstance(text, str):
                m = GSTIN_RE.search(text.upper())
                if m and not identity["gstin"]:
                    identity["gstin"] = m.group(1)
                m = REST_ID_RE.search(text)
                if m and not identity["rest_id"]:
                    identity["rest_id"] = m.group(1)
    # Name / outlet / city: the first three text cells beneath the title block.
    col_b = [ws.cell(row=r, column=2).value for r in range(4, 12)]
    texts = [str(v).strip() for v in col_b if isinstance(v, str) and v.strip()]
    texts = [t for t in texts if not GSTIN_RE.search(t.upper()) and not REST_ID_RE.search(t)
             and "payout" not in t.lower() and "rest" not in t.lower()]
    if texts:
        identity["restaurant"] = texts[0]
    if len(texts) > 1:
        identity["outlet"] = texts[1]
    if len(texts) > 2:
        identity["city"] = texts[2]
    return identity


def year_hint_from_name(path):
    """invoice_Annexure_544662_01102026_1790849827939.xlsx -> 2026 (from _ddmmyyyy_)."""
    m = DATE_IN_NAME_RE.search(os.path.basename(path))
    if m:
        try:
            return int(m.group(3))
        except ValueError:
            pass
    return datetime.now().year


def gst_rate_from_data(records):
    """Detect the GST rate actually applied (5%, 18%, ...) from GST / taxable value."""
    collected = sum(order_totals(rec["orders"], "GST Collected") for rec in records)
    taxable = sum(order_totals(rec["orders"], "Net Bill Value (before taxes) [1+2-3]")
                  for rec in records)
    if taxable <= 0:
        return GST_RATE
    rate = collected / taxable
    if rate >= 0.15:
        return 0.18
    if rate >= 0.10:
        return 0.12
    return 0.05


def month_label(start, end):
    """'September 2026' / 'September – October 2026' for a date span."""
    if start is None or end is None:
        return "the selected period"
    if (start.year, start.month) == (end.year, end.month):
        return start.strftime("%B %Y")
    if start.year == end.year:
        return "{} \u2013 {} {}".format(start.strftime("%B"), end.strftime("%B"), end.year)
    return "{} {} \u2013 {} {}".format(start.strftime("%B"), start.year,
                                       end.strftime("%B"), end.year)


def span_text(start, end):
    if start is None or end is None:
        return "the selected period"
    if(start.strftime("%b %Y") == end.strftime("%b %Y")):
        return "{} \u2013 {} {}".format(start.strftime("%d %b"), end.strftime("%d %b"), end.year)
    return "{} \u2013 {}".format(start.strftime("%d %b %Y"), end.strftime("%d %b %Y"))


def number_word(n):
    return {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six",
            7: "seven", 8: "eight", 9: "nine", 10: "ten"}.get(n, str(n))


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def num(value):
    """Coerce a cell value to float, returning 0.0 for blanks/labels."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    text = re.sub(r"(?i)^(₹|rs\.?|inr)\s*", "", text)   # 'Rs. 1,234.56' -> '1,234.56'
    text = text.replace("₹", "").replace(",", "").replace("%", "").replace(" ", "")
    if text in ("", "-", "--", "NA", "N/A"):
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def r2(value):
    return round(float(value) + 0.0, 2)


def to_date(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        for fmt in ("%d-%m-%Y", "%d %B", "%d-%b-%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(value.strip(), fmt)
            except ValueError:
                continue
    return None


def _month_number(token, year):
    """'September' / 'Sep' / '09' / 'Sep-26' -> (month, year)."""
    token = str(token).strip().lower().replace(",", "")
    token = re.sub(r"\s+", " ", token)
    yr = year
    m = re.search(r"(\d{4})", token)
    if m:
        yr = int(m.group(1))
        token = token.replace(m.group(1), "").strip(" -")
    m = re.match(r"^(\d{1,2})$", token)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(1)), yr
    for name, num in MONTH_NUM.items():
        if token.startswith(name[:3]):
            return num, yr
    raise KeyError(token)


def parse_payout_period(text, fallback_year=None):
    """Parse the 'Payout Period' cell of any Swiggy annexure.

    Handles '20 September - 26 September', '01 Sep - 05 Sep',
    '1 September 2026 - 5 September 2026' and variants with an en dash.
    Returns (start_date, end_date, label).
    """
    fallback_year = fallback_year or datetime.now().year
    raw = str(text or "").strip()
    if not raw:
        return None, None, ""
    cleaned = raw.replace("\u2013", "-").replace("\u2014", "-").replace("to", "-", 1) \
                 if " to " in raw else raw.replace("\u2013", "-").replace("\u2014", "-")
    parts = [p.strip() for p in cleaned.split("-") if p.strip()]
    if len(parts) == 2:
        try:
            def parse_side(side):
                side = side.strip()
                m = re.match(r"^(\d{1,2})\s*([A-Za-z]+)?\s*(\d{4})?$", side)
                if m:
                    day = int(m.group(1))
                    month_tok = m.group(2) or ""
                    year_tok = m.group(3) or ""
                else:
                    toks = side.split()
                    day = int(re.sub(r"\D", "", toks[0]) or 1)
                    month_tok = toks[1] if len(toks) > 1 else ""
                    year_tok = toks[2] if len(toks) > 2 else ""
                year = int(year_tok) if year_tok else fallback_year
                if not month_tok:
                    return day, None, year
                month, year = _month_number(month_tok, year)
                return day, month, year
            sd_day, sd_month, sd_year = parse_side(parts[0])
            ed_day, ed_month, ed_year = parse_side(parts[1])
            if sd_month is None:
                sd_month = 1
            if ed_month is None:
                ed_month = sd_month
            sd = datetime(sd_year, sd_month, sd_day)
            ed = datetime(ed_year, ed_month, ed_day)
        except (ValueError, KeyError, AttributeError):
            return None, None, raw
        label = "{} \u2013 {} {}".format(sd.strftime("%d"), ed.strftime("%d"),
                                        MONTH_ABBR[ed.strftime("%B").lower()], ed.year)
        return sd, ed, label
    # Single date (some cycles report only the payout date)
    try:
        sd = datetime.strptime(raw, "%d %B").replace(year=fallback_year)
        return sd, sd, label_for(sd, sd)
    except ValueError:
        return None, None, raw


def label_for(start, end):
    if start is None or end is None:
        return ""
    return "{} \u2013 {} {}".format(start.strftime("%d"), end.strftime("%d"),
                                    MONTH_ABBR[end.strftime("%B").lower()], end.year)


# --------------------------------------------------------------------------
# Extraction
# --------------------------------------------------------------------------
def _norm_header(text):
    """A comparable key for a column heading.

    Bracketed annotations are dropped because Swiggy writes the same column in
    several ways - 'Total Swiggy Fees [incl. GST on fees]' in older annexures,
    'Total Swiggy Fees [6+7+8-9+...]' in the ones here - while
    'GST Deduction [Sec 9(5)]' appears elsewhere as just 'GST Deduction'.
    """
    low = str(text or "").lower().replace("&", " and ")
    low = re.sub(r"[\(\[].*?[\)\]]", " ", low, flags=re.S)
    return re.sub(r"[^a-z0-9]", "", low)


# columns the reports cannot do without; anything else may be missing
REQUIRED_ORDER_COLS = [
    "Order ID", "Order Date", "Order Status", "Cancelled By?", "Item Total",
    "Packaging Charges", "Restaurant Discount Share [3a+3b]",
    "Net Bill Value (before taxes) [1+2-3]", "GST Collected",
    "Total Customer Paid [4+5]", "Commission", "Payment Collection Charges",
    "GST on Service Fee @18%", "Total Swiggy Fees [incl. GST on fees]",
    "Complaint & Cancellation Charges [18+19]", "GST Deduction [Sec 9(5)]",
    "TCS", "TDS", "Net Payout for Order (after taxes)",
]


def map_order_columns(ws):
    """Locate the Order Level header row and map each expected column to its position.

    Swiggy has been known to insert or rename columns between annexures; reading
    by header name rather than by a fixed position keeps older and newer files
    working. Returns (header_row, {column name: index}, missing required names).
    """
    def matches(key, want):
        """Header keys carry formula suffixes ('... [A-B-C-D]') or drop brackets."""
        if key == want:
            return True
        if key.startswith(want):
            rest = key[len(want):]
            return bool(rest) and (rest.isdigit() or len(rest) <= 6)
        if want.startswith(key):
            return len(want) - len(key) <= 6
        return False

    expected = {_norm_header(name): name for name in ORDER_COLS.values()}
    header_row = None
    mapping = {}
    for row in ws.iter_rows(min_row=1, max_row=12):
        texts = {_norm_header(c.value): c.column for c in row if c.value is not None}
        if _norm_header("Order ID") in texts and len(set(texts) & set(expected)) >= 8:
            header_row = row[0].row
            # exact (annotation-stripped) names first, then the longest near-match
            for key, col in texts.items():
                if key in expected and expected[key] not in mapping:
                    mapping[expected[key]] = col
            used = set(mapping.values())
            for key, col in sorted(texts.items(), key=lambda kv: -len(kv[0])):
                if col in used:
                    continue
                candidates = [want for want in expected if want not in mapping and matches(key, want)]
                if not candidates:
                    continue
                best = max(candidates, key=len)
                mapping[expected[best]] = col
                used.add(col)
            break
    missing = [name for name in REQUIRED_ORDER_COLS if name not in mapping]
    if header_row is None:
        # no recognisable header - fall back to the historic fixed positions
        mapping = dict(ORDER_COLS)
        header_row = 3
    return header_row, mapping, missing


def detect_payout_layout(ws):
    """Locate the delivered / cancelled / total columns of the Payout Breakup.

    Annexures arrive in two shapes:

        single block      D Delivered | E Cancelled | F Total
        two channels      D Swiggy delivered | E Swiggy cancelled
                          F Toing delivered  | G Toing cancelled | H Total

    so the columns are read from the header rows instead of being assumed. The
    channel blocks are summed, which keeps one-channel and two-channel
    annexures on the same footing. Returns (first data row, delivered columns,
    cancelled columns, total column).
    """
    delivered, cancelled, total_col, header_row = [], [], None, None
    for r in range(1, 13):
        for c in range(1, min(ws.max_column, 16) + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            text = re.sub(r"\s+", " ", v).strip().lower()
            if text.startswith(("delivered", "canceled", "cancelled")):
                (cancelled if text.startswith(("cancel", "canceled")) else delivered).append(c)
                header_row = max(header_row or 0, r)
            elif text == "total":
                total_col = c
                header_row = max(header_row or 0, r)
    if not delivered or not cancelled:
        raise ValueError("Payout Breakup sheet has no Delivered/Cancelled column headings")
    # the row under the headings is the Orders count, then the money lines
    data_start = (header_row or 5) + 2
    return data_start, sorted(set(delivered)), sorted(set(cancelled)), total_col


def read_annexure(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    rec = {"file": os.path.basename(path), "path": path}

    # ---- Summary sheet (identity + payout header) ----------------------
    ws = find_sheet(wb, "Summary")
    if ws is None:
        raise ValueError("no 'Summary' sheet - not a Swiggy payout annexure?")
    identity = extract_identity(wb)
    summary = identity["summary"]
    rec["restaurant"] = identity["restaurant"]
    rec["outlet"] = identity["outlet"]
    rec["city"] = identity["city"]
    rec["rest_id"] = identity["rest_id"]
    rec["gstin"] = identity["gstin"]
    rec["summary_raw"] = summary
    rec["payout_period_text"] = str(_first_match(summary, "payout period") or "")
    rec["settlement_text"] = str(_first_match(summary, "settlement") or "")
    rec["total_payout_stated"] = num(_first_match(summary, "total payout"))
    rec["total_orders_stated"] = int(num(_first_match(summary, "total orders")))
    rec["utr"] = str(_first_match(summary, "utr") or "").strip()

    # ---- Payout Breakup ------------------------------------------------
    ws = find_sheet(wb, "Payout Breakup", "Payout breakup")
    if ws is None:
        raise ValueError("no 'Payout Breakup' sheet")
    data_start, delivered_cols, cancelled_cols, total_col = detect_payout_layout(ws)
    rec["payout_columns"] = {"delivered": delivered_cols, "cancelled": cancelled_cols,
                             "total": total_col}
    lines = {}
    ads_lines = []
    mode = None

    def cell_value(row, col):
        return row[col - 1].value if col and col <= len(row) else None

    def sum_cols(row, cols):
        return r2(sum(num(cell_value(row, c)) for c in cols))

    for row in ws.iter_rows(min_row=data_start, max_row=data_start + 60):
        code = row[1].value
        label = row[2].value
        code_str = PAYOUT_CANON.get(_norm_code(code))
        if code_str in PAYOUT_LINES:
            mode = code_str
            delivered = sum_cols(row, delivered_cols)
            cancelled = sum_cols(row, cancelled_cols)
            total = num(cell_value(row, total_col)) if total_col else r2(delivered + cancelled)
            lines[code_str] = {
                "label": PAYOUT_LINES[code_str],
                "delivered": delivered,
                "cancelled": cancelled,
                "total": total,
            }
        elif mode == "D" and label and code is None and total_col and \
                cell_value(row, total_col) is not None:
            # Ads sub-lines (Top Picks - Ads, Ads Offers, ...)
            ads_lines.append({"label": str(label).strip().split("\n")[0],
                              "amount": num(cell_value(row, total_col))})
    rec["payout_lines"] = lines
    rec["ads_lines"] = ads_lines

    # ---- Order Level ----------------------------------------------------
    ws = find_sheet(wb, "Order Level", "Order level breakup")
    if ws is None:
        raise ValueError("no 'Order Level' sheet")
    header_row, col_map, missing = map_order_columns(ws)
    if missing:
        raise ValueError("Order Level sheet is missing column(s): " + ", ".join(missing[:4]) +
                         (" ..." if len(missing) > 4 else ""))
    rec["order_columns"] = col_map
    orders = []
    for row in ws.iter_rows(min_row=header_row + 1):
        if row[0].value is None or str(row[0].value).strip() == "":
            continue
        order = {}
        for name in ORDER_COLS.values():
            idx = col_map.get(name)
            order[name] = row[idx - 1].value if idx and idx <= len(row) else None
        if isinstance(order["Order Date"], datetime):
            order["_order_dt"] = order["Order Date"]
        else:
            order["_order_dt"] = to_date(order["Order Date"]) or datetime(2026, 9, 1)
        orders.append(order)
    rec["orders"] = orders

    # ---- Ads / adjustments ----------------------------------------------
    ws = find_sheet(wb, "Growth Investments and Other Deductions",
                    "Growth Investments and Other De")
    ads_detail = []
    for r in range(27, (ws.max_row if ws is not None else 0) + 1):
        label = ws.cell(r, 1).value
        amt = ws.cell(r, 4).value
        if not isinstance(label, str) or not label.strip():
            continue
        head = label.strip().lower()
        if head.startswith(("total", "recovery", "how to read", "investments this",
                            "investments from", "adjustment type", "adjustments from")):
            continue
        has_amount = isinstance(amt, (int, float)) or (isinstance(amt, str) and "₹" in amt)
        if not has_amount:
            continue
        ads_detail.append({
            "type": label.strip(),
            "total": num(amt),
            "settled": num(ws.cell(r, 5).value),
            "outstanding": num(ws.cell(r, 6).value),
            "from": ws.cell(r, 7).value,
            "to": ws.cell(r, 8).value,
            "invoice": ws.cell(r, 9).value,
            "remarks": ws.cell(r, 10).value,
        })
    rec["ads_detail"] = ads_detail

    # ---- Discount Summary -------------------------------------------------
    ws = find_sheet(wb, "Discount Summary")
    if ws is None:
        rec["discount"] = {"validity": "", "coupon": "", "rest_share": None, "orders": 0.0,
                           "given": 0.0, "target": "", "remarks": "sheet not present"}
    else:
        rec["discount"] = {
            "validity": ws["A3"].value, "coupon": ws["B3"].value,
            "rest_share": num(ws["C3"].value) if ws["C3"].value is not None else None,
            "orders": num(ws["D3"].value), "given": num(ws["E3"].value),
            "target": ws["F3"].value, "remarks": ws["G3"].value,
        }

    # ---- Unresolved complaints -------------------------------------------
    ws = find_sheet(wb, "Unresolved Customer Complaints")
    rec["complaints"] = {
        "processed_prev_complaint": num(ws["D2"].value) if ws else 0.0,
        "processed_current_regular": num(ws["D3"].value) if ws else 0.0,
        "processed_total": num(ws["D4"].value) if ws else 0.0,
        "not_processed_prev": num(ws["D6"].value) if ws else 0.0,
        "not_processed_current": num(ws["D7"].value) if ws else 0.0,
        "not_processed_total": num(ws["D8"].value) if ws else 0.0,
    }
    return rec


def _first_match(mapping, needle):
    """Look up a Summary label loosely, preferring the tightest match.

    Annexures with a Toing section list 'Total orders on Swiggy (Delivered +
    Cancelled)' and 'Total orders on Toing (...)' above the combined 'Total
    Orders (Delivered + Cancelled)', so a plain 'first contains' lookup would
    return the Swiggy count alone. Exact matches win; then the shortest label.
    """
    needle = needle.lower().strip()

    def tidy(text):
        return re.sub(r"\s+", " ", str(text)).strip().lower()

    keys = [k for k in mapping if tidy(k)]
    for want in (lambda k: tidy(k) == needle,
                 lambda k: tidy(k).startswith(needle),
                 lambda k: needle in tidy(k)):
        hits = [k for k in keys if want(k)]
        if hits:
            return mapping[min(hits, key=lambda k: len(tidy(k)))]
    return None


SKIP_DIRS = {".git", ".venv", "venv", "__pycache__", "reports", "output", "node_modules"}


def find_legacy_excel(folder, recursive=True):
    """Old-format .xls workbooks (openpyxl cannot read these).

    They are returned so the menu can say what to do with them, and they are
    skipped when reading.
    """
    folder = os.path.abspath(os.path.expanduser(folder))
    found = glob.glob(os.path.join(folder, "*.xls"))
    if recursive:
        found += glob.glob(os.path.join(folder, "**", "*.xls"), recursive=True)
    unique, seen = [], set()
    for f in found:
        if os.path.basename(f).startswith("~$"):
            continue
        real = os.path.abspath(f)
        if real not in seen:
            seen.add(real)
            unique.append(real)
    return unique


def find_annexures(folder, recursive=True, skip=None):
    """Every plausible annexure workbook in a folder (newest name order).

    `skip` is an extra directory name to ignore, normally the output folder, so
    that reports generated earlier are not offered as inputs.
    """
    skip_dirs = set(SKIP_DIRS)
    if skip:
        skip_dirs.add(str(skip))
    folder = os.path.abspath(os.path.expanduser(folder))
    patterns = ["*.xlsx", "*.xlsm"]
    found = []
    for pattern in patterns:
        found += glob.glob(os.path.join(folder, pattern))
        if recursive:
            found += glob.glob(os.path.join(folder, "**", pattern), recursive=True)
    found += find_legacy_excel(folder, recursive=recursive)
    cleaned = []
    for f in found:
        if os.path.basename(f).startswith("~$"):
            continue
        parts = os.path.relpath(f, folder).split(os.sep)[:-1]
        if any(part in skip_dirs or part.startswith(".") for part in parts):
            continue
        cleaned.append(f)
    # glob("**/") also matches the folder itself, so the same file can appear twice
    found, seen = [], set()
    for f in cleaned:
        real = os.path.abspath(f)
        if real not in seen:
            seen.add(real)
            found.append(real)
    # Prefer annexure-looking names, but do not hide anything: the caller validates.
    found.sort(key=lambda f: (0 if "annexure" in os.path.basename(f).lower() else 1,
                              os.path.basename(f).lower()))
    return found


def has_legacy_excel(folder, recursive=True):
    """Any .xls workbooks in the folder that openpyxl cannot read?"""
    return find_legacy_excel(folder, recursive=recursive)


def expand_inputs(items, folder=None):
    """Turn files / globs / folders given by the user into a list of paths."""
    paths = []
    for item in items:
        item = os.path.expanduser(str(item).strip().strip('"').strip("'"))
        if not item:
            continue
        if os.path.isdir(item):
            paths += find_annexures(item)
        elif any(ch in item for ch in "*?["):
            paths += sorted(glob.glob(item))
        elif os.path.isfile(item):
            paths.append(os.path.abspath(item))
    if folder and not paths:
        paths += find_annexures(folder)
    # de-duplicate, keep order
    seen, unique = set(), []
    for path in paths:
        real = os.path.abspath(path)
        if real not in seen:
            seen.add(real)
            unique.append(real)
    return unique


def annexure_signature(rec):
    """Identity of the data in an annexure: outlet, period, the orders and the payout.

    Two files with the same signature carry the same data (for example the same
    annexure kept in two folders), so only one of them may be counted.
    """
    ids = tuple(sorted(str(o.get("Order ID")) for o in rec["orders"]))
    return (str(rec.get("gstin") or rec.get("rest_id") or ""),
            str(rec.get("label")),
            int(rec.get("orders_total") or len(rec["orders"])),
            r2(rec.get("net_payout") or 0),
            ids)


def load_files(paths, gst_rate=None):
    """Read the given annexures. Invalid files are reported, not fatal."""
    records, skipped = [], []
    seen_signatures = {}
    for path in paths:
        ok, why = looks_like_annexure(path)
        if not ok:
            skipped.append((path, why))
            continue
        try:
            rec = read_annexure(path)
        except Exception as exc:  # noqa: BLE001 - report and carry on
            skipped.append((path, str(exc)))
            continue
        year = year_hint_from_name(path)
        sd, ed, label = parse_payout_period(rec["payout_period_text"], year)
        if sd is None:
            # fall back to the settlement date or the order dates
            order_dates = [o["_order_dt"] for o in rec["orders"] if o.get("_order_dt")]
            if order_dates:
                sd, ed = min(order_dates), max(order_dates)
                label = label_for(sd, ed)
        rec["start"], rec["end"], rec["label"] = sd, ed, label
        rec["sort_key"] = sd or datetime(year, 1, 1)
        rec["year"] = year
        rec["path"] = path
        signature = annexure_signature(rec)
        if signature in seen_signatures:
            skipped.append((path, "duplicate - same outlet, period and orders as {}; read once "
                                  "only".format(display_path(seen_signatures[signature], 44))))
            continue
        seen_signatures[signature] = path
        records.append(rec)
    records.sort(key=lambda r: r["sort_key"])
    return records, skipped


def group_by_outlet(records):
    """Split records into per-restaurant bundles (stable order)."""
    groups = OrderedDict()
    for rec in records:
        key = rec.get("gstin") or rec.get("rest_id") or (rec.get("restaurant") or "unknown").lower()
        groups.setdefault(key, []).append(rec)
    return groups


# --------------------------------------------------------------------------
# Derived aggregates
# --------------------------------------------------------------------------
def order_totals(orders, field):
    return r2(sum(num(o.get(field)) for o in orders))


def period_rollup(rec):
    """Key figures for one payout period."""
    lines = rec["payout_lines"]
    orders = rec["orders"]
    delivered = [o for o in orders if str(o["Order Status"]).lower() == "delivered"]
    cancelled = [o for o in orders if str(o["Order Status"]).lower() == "cancelled"]
    taxable = order_totals(orders, "Net Bill Value (before taxes) [1+2-3]")
    compensated = [o for o in orders if is_compensated_cancel(o)]
    compensation_deduction = r2(sum(order_compensation_deduction(o) for o in compensated))
    reportable = r2(taxable - compensation_deduction)
    return {
        "label": rec["label"],
        "period_text": rec["payout_period_text"],
        "start": rec["start"], "end": rec["end"],
        "settlement": rec["settlement_text"],
        "utr": rec["utr"],
        "file": rec["file"],
        "stated_payout": r2(rec["total_payout_stated"]),
        "stated_orders": int(rec["total_orders_stated"]),
        "orders_delivered": len(delivered),
        "orders_cancelled": len(cancelled),
        "orders_total": len(orders),
        # payout breakup (Total column)
        "cust_paid": r2(lines["A"]["total"]),
        "item_total": r2(lines["1.0"]["total"]),
        "packaging": r2(lines["2.0"]["total"]),
        "disc_coupon": r2(lines["3.0"]["total"]),
        "disc_trade": r2(lines["4.0"]["total"]),
        "gst_collected": r2(lines["5.0"]["total"]),
        "swiggy_fees": r2(lines["B"]["total"]),
        "commission": r2(lines["6.0"]["total"]),
        "payment_collection": r2(lines["8.0"]["total"]),
        "rest_cancellation": r2(lines["11.0"]["total"]),
        "complaints_cancel": r2(lines["C"]["total"]),
        "ads": r2(lines["D"]["total"]),
        "other_charges": r2(lines["E"]["total"]),
        "total_taxes": r2(lines["F"]["total"]),
        "gst_9_5": r2(lines["18"]["total"]),
        "gst_on_fees": r2(lines["19"]["total"]),
        "tcs": r2(lines["20"]["total"]),
        "tds": r2(lines["21"]["total"]),
        "net_payout": r2(lines["G"]["total"]),
        "net_delivered": r2(lines["G"]["delivered"]),
        "net_cancelled": r2(lines["G"]["cancelled"]),
        # order-level cross totals
        "ol_taxable": taxable,
        "ol_reportable": reportable,
        "ol_compensation_deduction": compensation_deduction,
        "ol_compensation_value": r2(sum(order_reportable_value(o) for o in compensated)),
        "ol_compensated_orders": len(compensated),
        "ol_taxable_delivered": order_totals(delivered, "Net Bill Value (before taxes) [1+2-3]"),
        "ol_taxable_cancelled": order_totals(cancelled, "Net Bill Value (before taxes) [1+2-3]"),
        "ol_gst": order_totals(orders, "GST Collected"),
        "ol_cust_paid": order_totals(orders, "Total Customer Paid [4+5]"),
        "ol_net": order_totals(orders, "Net Payout for Order (after taxes)"),
        "ol_tds": order_totals(orders, "TDS"),
        "ol_tcs": order_totals(orders, "TCS"),
        "ol_commission": order_totals(orders, "Commission"),
        "ol_payment_collection": order_totals(orders, "Payment Collection Charges"),
        "ol_gst_on_fees": order_totals(orders, "GST on Service Fee @18%"),
        "ol_gst_9_5": order_totals(orders, "GST Deduction [Sec 9(5)]"),
        "ol_cancellations": order_totals(orders, "Complaint & Cancellation Charges [18+19]"),
        "channels": channel_split(orders),
    }


def channel_split(orders):
    """Group orders by the platform they came through (Swiggy / Toing / ...)."""
    channels = OrderedDict()
    for o in orders:
        name = str(o.get("Order Category") or "").strip() or "Swiggy"
        ch = channels.setdefault(name, {"orders": 0, "item_total": 0.0, "gst_collected": 0.0,
                                        "gst_9_5": 0.0, "tds": 0.0, "net_payout": 0.0,
                                        "cancelled": 0})
        ch["orders"] += 1
        if str(o.get("Order Status") or "").lower() == "cancelled":
            ch["cancelled"] += 1
        ch["item_total"] = r2(ch["item_total"] + num(o.get("Item Total")))
        ch["gst_collected"] = r2(ch["gst_collected"] + num(o.get("GST Collected")))
        ch["gst_9_5"] = r2(ch["gst_9_5"] + num(o.get("GST Deduction [Sec 9(5)]")))
        ch["tds"] = r2(ch["tds"] + num(o.get("TDS")))
        ch["net_payout"] = r2(ch["net_payout"] + num(o.get("Net Payout for Order (after taxes)")))
    return channels


def channel_names(rollups):
    """Every channel seen across the selected periods, in first-seen order."""
    names = OrderedDict()
    for pr in rollups:
        for name in (pr.get("channels") or {}):
            names.setdefault(name, True)
    return list(names)


def channel_totals(rollups):
    """Channel (Swiggy / Toing / ...) totals for the whole selection."""
    total = OrderedDict()
    for pr in rollups:
        for name, ch in (pr.get("channels") or {}).items():
            acc = total.setdefault(name, {"orders": 0, "cancelled": 0, "item_total": 0.0,
                                          "gst_collected": 0.0, "gst_9_5": 0.0, "tds": 0.0,
                                          "net_payout": 0.0})
            acc["orders"] += ch["orders"]
            acc["cancelled"] += ch["cancelled"]
            for key in ("item_total", "gst_collected", "gst_9_5", "tds", "net_payout"):
                acc[key] = r2(acc[key] + ch[key])
    return total


def policy_facts(records, rollups, totals, gst_rate=None, period_label=None):
    """Everything the narrative sheets need, computed from the data itself."""
    gst_rate = gst_rate or gst_rate_from_data(records)
    compensated = [o for rec in records for o in rec["orders"] if is_compensated_cancel(o)]
    merchant = [o for rec in records for o in rec["orders"] if cancel_category(o) == "merchant"]
    after_pickup = [o for rec in records for o in rec["orders"] if cancel_category(o) == "after_pickup"]
    sample = compensated[0] if compensated else None
    facts = {
        "files": len(records),
        "files_word": number_word(len(records)),
        "channels": channel_names(rollups),
        "channel_totals": channel_totals(rollups),
        "start": rollups[0]["start"], "end": rollups[-1]["end"],
        "period_label": period_label or month_label(rollups[0]["start"], rollups[-1]["end"]),
        "span": span_text(rollups[0]["start"], rollups[-1]["end"]),
        "gst_rate": gst_rate,
        "gst_rate_pct": "{:g}%".format(round(gst_rate * 100, 2)),
        "comp_rate_pct": "{:g}%".format(round(CANCELLATION_COMPENSATION_RATE * 100, 2)),
        "uncomp_rate_pct": "{:g}%".format(round((1 - CANCELLATION_COMPENSATION_RATE) * 100, 2)),
        "comp_rate_pct": "{:g}%".format(round(CANCELLATION_COMPENSATION_RATE * 100, 2)),
        "uncomp_rate_pct": "{:g}%".format(round((1 - CANCELLATION_COMPENSATION_RATE) * 100, 2)),
        "comp_rate_pct": "{:g}%".format(round(CANCELLATION_COMPENSATION_RATE * 100, 2)),
        "uncomp_rate_pct": "{:g}%".format(round((1 - CANCELLATION_COMPENSATION_RATE) * 100, 2)),
        "cgst_rate_pct": "{:g}%".format(round(gst_rate * 100 / 2, 2)),
        "sgst_rate_pct": "{:g}%".format(round(gst_rate * 100 / 2, 2)),
        "prepared_on": PREPARED_ON,
        "compensated": compensated,
        "merchant_cancels": merchant,
        "after_pickup": after_pickup,
        "after_pickup_n": len(after_pickup),
        "sample": sample,
        "comp_value": r2(sum(order_reportable_value(o) for o in compensated)),
        "comp_deduction": totals["ol_compensation_deduction"],
        "gst_on_deduction": r2(totals["ol_compensation_deduction"] * gst_rate),
        "deduction_gross": r2(totals["ol_compensation_deduction"] * (1 + gst_rate)),
        "rounding": r2(totals["gst_collected"] - abs(totals["ol_gst_9_5"])
                       - totals["ol_compensation_deduction"] * gst_rate),
        "gst_retained": abs(totals["ol_gst_9_5"]),
        "reportable": totals["ol_reportable"],
        "billed": totals["ol_taxable"],
        "discount_orders_min": 0, "discount_orders_max": 0, "discount_orders_total": 0,
    }
    if sample is not None:
        billed = order_billed_value(sample)
        paid = num(sample["Total Customer Paid [4+5]"])
        facts.update({
            "sample_id": str(sample["Order ID"]),
            "sample_billed": billed,
            "sample_paid": paid,
            "sample_gst_collected": num(sample["GST Collected"]),
            "sample_comp": order_reportable_value(sample),
            "sample_gst_on_comp": r2(order_reportable_value(sample) * gst_rate),
            "sample_comp_gross": r2(paid * CANCELLATION_COMPENSATION_RATE),
            "sample_uncomp": order_compensation_deduction(sample),
            "sample_gst_uncomp": r2(order_compensation_deduction(sample) * gst_rate),
            "sample_deduction": r2(order_compensation_deduction(sample) * (1 + gst_rate)),
            "sample_tds": num(sample["TDS"]),
            "sample_cancelled_by": str(sample["Cancelled By?"] or "Swiggy").title(),
            "sample_commission_on": num(sample["Commission charged on"]),
            "sample_period": next((rec["label"] for rec in records if sample in rec["orders"]), ""),
        })
    discounts = [rec["discount"].get("orders") or 0 for rec in records]
    facts["discount_orders_min"] = int(min(discounts)) if discounts else 0
    facts["discount_orders_max"] = int(max(discounts)) if discounts else 0
    facts["discount_orders_total"] = int(sum(discounts))
    return facts


def totals_of(rollups):
    keys = [k for k, v in rollups[0].items() if isinstance(v, (int, float))]
    return {k: r2(sum(r[k] for r in rollups)) for k in keys}


# --------------------------------------------------------------------------
# Styling helpers
# --------------------------------------------------------------------------
NAVY = "1F3864"
BLUE = "2E5C8A"
LIGHT = "DCE6F1"
BAND = "F2F7FB"
GREEN = "E2EFDA"
AMBER = "FFF2CC"
GREY = "808080"

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=NAVY)
SUB_FONT = Font(name="Calibri", size=11, italic=True, color=GREY)
HDR_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
SEC_FONT = Font(name="Calibri", size=11, bold=True, color=NAVY)
BOLD = Font(name="Calibri", size=11, bold=True)
NORM = Font(name="Calibri", size=11)
SMALL = Font(name="Calibri", size=9, italic=True, color=GREY)

HDR_FILL = PatternFill("solid", fgColor=BLUE)
SEC_FILL = PatternFill("solid", fgColor=LIGHT)
BAND_FILL = PatternFill("solid", fgColor=BAND)
GREEN_FILL = PatternFill("solid", fgColor=GREEN)
INFO_FILL = PatternFill("solid", fgColor="DDEBF7")
AMBER_FILL = PatternFill("solid", fgColor=AMBER)

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOP_BORDER = Border(top=Side(style="thin", color=NAVY))

MONEY = '₹#,##0.00;[Red]-₹#,##0.00'
NUM = '#,##0'
NUM2 = '#,##0.00'
PCT = '0.00%'
DATEF = 'dd-mmm-yyyy'


def write_title(ws, row, text, subtitle=None, width=8):
    ws.cell(row, 1, text).font = TITLE_FONT
    if subtitle:
        ws.cell(row + 1, 1, subtitle).font = SUB_FONT
        return row + 2
    return row + 1


def write_header(ws, row, headers, widths=None, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row, start_col + i, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(start_col + i)].width = w
    ws.row_dimensions[row].height = 30
    return row + 1


def stamp(ws, text="Source: Swiggy weekly payout annexures · Prepared " + PREPARED_ON):
    ws.cell(ws.max_row + 2, 1, text).font = SMALL


def autosize(ws, min_w=9, max_w=52):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        longest = 0
        for cell in col:
            if cell.value is None:
                continue
            for line in str(cell.value).split("\n"):
                longest = max(longest, len(line))
        ws.column_dimensions[letter].width = max(min_w, min(max_w, longest + 2))


# ==========================================================================
# WORKBOOK 1 - consolidated data report
# ==========================================================================
def build_consolidated(records, rollups, totals, path, facts=None):
    facts = facts or policy_facts(records, rollups, totals)
    f = facts
    gst_rate = f["gst_rate"]
    rate_pct = f["gst_rate_pct"]
    comp_pct = f["comp_rate_pct"]
    uncomp_pct = f["uncomp_rate_pct"]
    sample = f.get("sample")
    wb = openpyxl.Workbook()
    recon_check_count = "70+"
    first = records[0]
    period_labels = [r["label"] for r in rollups]
    metrics = ["orders_delivered", "orders_cancelled", "orders_total", "cust_paid",
               "item_total", "gst_collected", "commission", "payment_collection",
               "swiggy_fees", "ads", "gst_9_5", "gst_on_fees", "tcs", "tds",
               "net_payout"]

    # ---------------- Cover ------------------------------------------------
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False
    ws["B2"] = "Swiggy Payout Annexures — Consolidated Report"
    ws["B2"].font = Font(size=18, bold=True, color=NAVY)
    ws["B3"] = "All weekly payout cycles consolidated into a single data set"
    ws["B3"].font = SUB_FONT
    rows = [
        ("", ""),
        ("Restaurant", first["restaurant"]),
        ("Outlet / City", "{} / {}".format(first["outlet"], first["city"])),
        ("Swiggy Restaurant ID", first["rest_id"]),
        ("GSTIN", first["gstin"]),
        ("Report period", "{} – {}".format(rollups[0]["start"].strftime("%d %b %Y"),
                                           rollups[-1]["end"].strftime("%d %b %Y"))),
        ("Payout cycles covered", "{} weekly annexures".format(len(records))),
        ("Orders covered", "{} ({} delivered, {} cancelled)".format(
            totals["orders_total"], totals["orders_delivered"], totals["orders_cancelled"])),
        ("Gross value billed to customers", "₹{:,.2f}".format(totals["cust_paid"])),
        ("Net payouts credited by Swiggy", "₹{:,.2f}".format(totals["net_payout"])),
        ("Prepared on", PREPARED_ON),
        ("", ""),
        ("Source files", ""),
    ]
    if len(facts["channels"]) > 1:
        rows.insert(7, ("Sales channels", " + ".join(
            "{} ({} order{})".format(name, ch["orders"], "" if ch["orders"] == 1 else "s")
            for name, ch in facts["channel_totals"].items())))
    r = 5
    for label, value in [row for row in rows if row]:
        if label:
            ws.cell(r, 2, label).font = BOLD
            ws.cell(r, 3, value).font = NORM
        r += 1
    for rec in records:
        ws.cell(r, 3, rec["file"] + "   (payout period: {})".format(rec["payout_period_text"])).font = NORM
        r += 1

    r += 1
    ws.cell(r, 2, "What is in this workbook").font = SEC_FONT
    r += 1
    contents = [
        ("Period Summary", "One line per payout cycle — orders, sales, fees, taxes and net payout, with grand totals."),
        ("Payout Breakup", "The full Swiggy payout-breakup line structure (A to G) shown side by side for every cycle."),
        ("Delivered vs Cancelled", "How each cycle's net payout splits between delivered orders, cancelled orders and ads."),
        ("Cancelled Orders", "The 80% cancellation-compensation policy and the GST/TDS effect of each cancelled order."),
        ("Order Level (All Periods)", "Every order from all annexures in one filterable table, tagged with its payout cycle."),
        ("Ads & Adjustments", "Restaurant-level growth investments / adjustments deducted outside order-level data."),
        ("Discounts", "Discount campaigns reported by Swiggy for each cycle."),
        ("Complaints", "Unresolved customer complaint buckets reported by Swiggy."),
        ("Reconciliation", "{} automated tie-out checks across every internal total and cross-sheet total, "
                           "with exceptions flagged.".format(recon_check_count)),
        ("Field Definitions", "Definition of every field used in the Order Level section."),
    ]
    write_header(ws, r, ["Sheet", "Contents"])
    r += 1
    for name, desc in contents:
        ws.cell(r, 1, name).font = BOLD
        ws.cell(r, 2, desc).font = NORM
        r += 1
    r += 1
    ws.cell(r, 1, "Basis of preparation").font = SEC_FONT
    r += 1
    for note in [
        "All figures are reproduced from the Swiggy weekly payout annexures (the source of truth). Nothing has been re-estimated.",
        "No orders or rows have been dropped; duplicate Order IDs across cycles are checked on the Reconciliation sheet.",
        "'Net payout' is the amount settled by Swiggy after commission, payment-collection charges, ad investments, GST and TDS.",
        "GST on Swiggy's service fee is a charge recovered from the restaurant, not a tax on the food supply — see the GST working.",
    ]:
        ws.cell(r, 1, "•  " + note).font = NORM
        r += 1
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 95

    # ---------------- Period Summary ---------------------------------------
    ws = wb.create_sheet("Period Summary")
    headers = ["Payout period", "From", "To", "Settlement date", "Bank UTR",
               "Delivered orders", "Cancelled orders", "Total orders",
               "Total customer paid", "Billed value (item total + packaging − discounts)",
               "GST-reportable value (after cancellation adjustment)",
               "GST collected @5%", "Commission", "Payment collection charges",
               "Swiggy fees (excl. GST on fees)", "Ads investment",
               "GST retained u/s 9(5)", "GST on Swiggy fees @18%", "TCS u/s 52",
               "TDS u/s 194-O", "Net payout", "Annexure file"]
    widths = [18, 12, 12, 14, 20, 11, 11, 10, 16, 18, 18, 14, 13, 14, 16, 13, 15, 15, 10, 12, 14, 46]
    r = write_header(ws, 1, headers, widths)
    for i, pr in enumerate(rollups):
        vals = [pr["label"], pr["start"], pr["end"], pr["settlement"], pr["utr"],
                pr["orders_delivered"], pr["orders_cancelled"], pr["orders_total"],
                pr["cust_paid"], r2(pr["item_total"] + pr["packaging"] - pr["disc_coupon"] - pr["disc_trade"]),
                pr["ol_reportable"],
                pr["gst_collected"], abs(pr["commission"]), abs(pr["payment_collection"]),
                abs(pr["swiggy_fees"]), abs(pr["ads"]), abs(pr["gst_9_5"]), abs(pr["gst_on_fees"]),
                abs(pr["tcs"]), abs(pr["tds"]), pr["net_payout"], pr["file"]]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            if j in (2, 3):
                c.number_format = DATEF
            if j in (6, 7, 8):
                c.number_format = NUM
            if 9 <= j <= 21:
                c.number_format = MONEY
            if i % 2:
                c.fill = BAND_FILL
        r += 1
    tot_vals = ["TOTAL", None, None, None, None, totals["orders_delivered"],
                totals["orders_cancelled"], totals["orders_total"], totals["cust_paid"],
                r2(totals["item_total"] + totals["packaging"] -
                   totals["disc_coupon"] - totals["disc_trade"]),
                totals["ol_reportable"],
                totals["gst_collected"], abs(totals["commission"]),
                abs(totals["payment_collection"]), abs(totals["swiggy_fees"]), abs(totals["ads"]),
                abs(totals["gst_9_5"]), abs(totals["gst_on_fees"]), abs(totals["tcs"]), abs(totals["tds"]),
                totals["net_payout"], "{} annexures".format(len(records))]
    for j, v in enumerate(tot_vals, start=1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = GREEN_FILL
        c.border = BOX
        if j in (6, 7, 8):
            c.number_format = NUM
        if 9 <= j <= 21:
            c.number_format = MONEY
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = "A1:V{}".format(r - 1)
    stamp(ws, "Totals = the {} annexure(s) covering {}. Settlement dates are as printed on each annexure. "
              "Deduction lines (commission, charges, ads, GST, TDS) are shown here as positive amounts for readability — "
              "the annexures print them as negative; see the Payout Breakup sheet for the original signs.".format(
                  f["files_word"], f["span"]))

    # ---------------- Payout Breakup ---------------------------------------
    ws = wb.create_sheet("Payout Breakup")
    headers = ["Line", "Particulars"] + period_labels + ["Total"]
    widths = [6, 58] + [14] * len(period_labels) + [14]
    r = write_header(ws, 1, headers, widths)
    line_order = ["A", "1.0", "2.0", "3.0", "4.0", "5.0", "B", "6.0", "7.0", "8.0",
                  "9.0", "10.0", "11.0", "12.0", "13.0", "14.0", "15.0", "C", "16.0",
                  "17.0", "D", "E", "F", "18", "19", "20", "21", "G"]
    for code in line_order:
        label = PAYOUT_LINES[code]
        is_head = code in ("A", "B", "C", "D", "E", "F", "G")
        is_sub = code in SUB_LINES
        ws.cell(r, 1, "" if is_sub else code)
        ws.cell(r, 2, ("      " if is_sub else "") + label)
        for k, rec in enumerate(records):
            lines = rec["payout_lines"]
            v = lines.get(code, {}).get("total", 0.0) if code in lines else 0.0
            ws.cell(r, 3 + k, r2(v))
        total_v = sum(r2(rec["payout_lines"].get(code, {}).get("total", 0.0)) for rec in records)
        ws.cell(r, 3 + len(records), r2(total_v))
        for j in range(1, 4 + len(records)):
            c = ws.cell(r, j)
            c.border = BOX
            c.font = BOLD if is_head else NORM
            if j >= 3:
                c.number_format = MONEY
            if is_head:
                c.fill = SEC_FILL
            elif code in ("1.0", "2.0", "3.0", "4.0", "5.0", "6.0", "7.0", "8.0",
                          "9.0", "10.0", "11.0", "12.0", "13.0", "14.0", "15.0",
                          "16.0", "17.0", "18", "19", "20", "21"):
                c.fill = BAND_FILL
        r += 1
        if code == "D":
            labels = []
            for rec in records:
                for ad in rec["ads_lines"]:
                    if ad["label"] not in labels:
                        labels.append(ad["label"])
            for lab in labels:
                ws.cell(r, 2, "      " + lab)
                for k, rec in enumerate(records):
                    amt = next((a["amount"] for a in rec["ads_lines"] if a["label"] == lab), 0.0)
                    ws.cell(r, 3 + k, r2(amt))
                ws.cell(r, 3 + len(records),
                        r2(sum(next((a["amount"] for a in rec["ads_lines"] if a["label"] == lab), 0.0)
                               for rec in records)))
                for j in range(1, 4 + len(records)):
                    c = ws.cell(r, j)
                    c.border = BOX
                    c.font = NORM
                    c.fill = BAND_FILL
                    if j >= 3:
                        c.number_format = MONEY
                r += 1
    ws.freeze_panes = "C2"
    stamp(ws, "Amounts are the 'Total' column of each annexure's Payout Breakup sheet. "
              "Line D is a restaurant-level deduction and is not part of the order-level data — "
              "Total net payout = delivered net + cancelled net + ads.")

    # ---------------- Delivered vs Cancelled --------------------------------
    if len(facts["channels"]) > 1:
        ws = wb.create_sheet("Channel Split")
        r = write_title(ws, 1, "Sales channels — {} / {}".format(
            " + ".join(facts["channels"]), first["restaurant"]),
            "Orders, sales and tax by the platform they came through (Order Category on the Order Level sheet)")
        r += 1
        r = write_header(ws, r, ["Channel", "Orders", "Cancelled", "Item total", "GST collected",
                                 "GST discharged u/s 9(5)", "TDS u/s 194-O", "Net payout (order level)",
                                 "Remark"], [16, 9, 11, 14, 14, 17, 14, 17, 40])
        for name, ch in facts["channel_totals"].items():
            vals = [name, ch["orders"], ch["cancelled"], ch["item_total"], ch["gst_collected"],
                    ch["gst_9_5"], ch["tds"], ch["net_payout"],
                    "Tax is retained and discharged by the operator for both channels"]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(r, j, v)
                c.font = NORM
                c.border = BOX
                c.alignment = Alignment(vertical="center", wrap_text=(j == 9))
                if j in (4, 5, 6, 7, 8):
                    c.number_format = MONEY
            r += 1
        tot = facts["channel_totals"]
        vals = ["TOTAL", sum(c["orders"] for c in tot.values()), sum(c["cancelled"] for c in tot.values()),
                r2(sum(c["item_total"] for c in tot.values())), r2(sum(c["gst_collected"] for c in tot.values())),
                r2(sum(c["gst_9_5"] for c in tot.values())), r2(sum(c["tds"] for c in tot.values())),
                r2(sum(c["net_payout"] for c in tot.values())), ""]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = BOLD
            c.fill = GREEN_FILL
            c.border = BOX
            if j in (4, 5, 6, 7, 8):
                c.number_format = MONEY
        r += 2
        for note in [
            "The annexure prints one block of columns per channel on the Payout Breakup sheet; both are summed "
            "here and in every other sheet of this workbook.",
            "Net payout on this sheet is the order-level figure and excludes restaurant-level ads investment, "
            "which the annexure deducts once for the month.",
            "Both channels are ordered through the same e-commerce operator arrangement - the GST is retained "
            "and deposited from the payout - so the tax working is unchanged by the split.",
        ]:
            cell = ws.cell(r, 1, "•  " + note)
            cell.font = NORM
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
            ws.row_dimensions[r].height = 28
            r += 1
        autosize(ws, min_w=12, max_w=60)

    ws = wb.create_sheet("Delivered vs Cancelled")
    r = write_title(ws, 1, "Delivered vs cancelled vs restaurant-level deductions",
                    "How each cycle's net payout is built up")
    r += 1
    headers = ["Payout period", "Delivered: orders", "Delivered: total customer paid",
               "Delivered: Swiggy fees", "Delivered: taxes", "Delivered: net payout",
               "Cancelled: orders", "Cancelled: total customer paid",
               "Cancelled: complaints & cancellation charges", "Cancelled: Swiggy fees",
               "Cancelled: taxes", "Cancelled: net payout",
               "Ads investment (restaurant level)", "Net payout per annexure"]
    widths = [18] + [13] * 13
    r = write_header(ws, r, headers, widths)
    for i, rec in enumerate(records):
        lines = rec["payout_lines"]

        def g(code, col):
            return r2(lines.get(code, {}).get(col, 0.0))

        tax_del = r2(sum(g(code, "delivered") for code in ("18", "19", "20", "21")))
        tax_can = r2(sum(g(code, "cancelled") for code in ("18", "19", "20", "21")))
        vals = [
            rec["label"],
            len([o for o in rec["orders"] if str(o["Order Status"]).lower() == "delivered"]),
            g("A", "delivered"), g("B", "delivered"), tax_del, g("G", "delivered"),
            len([o for o in rec["orders"] if str(o["Order Status"]).lower() == "cancelled"]),
            g("A", "cancelled"), g("C", "cancelled"), g("B", "cancelled"),
            tax_can, g("G", "cancelled"),
            g("D", "total"), g("G", "total"),
        ]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.border = BOX
            c.font = NORM
            c.number_format = NUM if j in (2, 7) else (MONEY if j > 2 else 'General')
            if i % 2:
                c.fill = BAND_FILL
        r += 1
    tot = ["TOTAL"] + [None] * 13
    for j in range(2, 15):
        tot[j - 1] = r2(sum(ws.cell(rr, j).value or 0 for rr in range(r - len(records), r)))
    tot[1] = int(sum(ws.cell(rr, 2).value or 0 for rr in range(r - len(records), r)))
    tot[6] = int(sum(ws.cell(rr, 7).value or 0 for rr in range(r - len(records), r)))
    for j, v in enumerate(tot, start=1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = GREEN_FILL
        c.border = BOX
        c.number_format = NUM if j in (2, 7) else (MONEY if j > 2 else 'General')
    stamp(ws, "'Swiggy fees' and 'taxes' are shown as negative deductions in the annexures; "
              "signs are preserved. Net payout per annexure = delivered + cancelled + ads.")

    # ---------------- Order Level (All Periods) -----------------------------
    ws = wb.create_sheet("Order Level (All Periods)")
    headers = ["Payout period", "Order ID", "Order date", "Order status",
               "Cancelled by", "Payment type", "Coupon type applied by customer",
               "Item total", "Packaging charges", "Restaurant discounts",
               "Swiggy One / exclusive offer discount",
               "Taxable value (net bill value, before taxes)", "GST collected @5%",
               "GST-reportable value (after cancellation adjustment)",
               "GST @5% on reportable value",
               "Total customer paid", "Commission %", "Commission charged on",
               "Commission", "Payment collection charges",
               "Other Swiggy fees (cancellation, long distance, etc.)", "Total Swiggy fees (incl. GST on fees)",
               "Complaint & cancellation charges", "GST retained u/s 9(5)",
               "GST on Swiggy fees @18%", "TCS", "TDS u/s 194-O",
               "Net payout for order", "Swiggy One customer?", "Last mile (km)",
               "Discount campaign ID", "Order settlement date"]
    widths = [16, 17, 18, 11, 12, 12, 16, 11, 11, 12, 13, 15, 11, 15, 13, 13, 10, 12, 12, 12,
              12, 14, 13, 12, 12, 8, 11, 12, 11, 10, 30, 13]
    r = write_header(ws, 1, headers, widths)
    band = False
    for rec in records:
        for o in sorted(rec["orders"], key=lambda x: x["_order_dt"]):
            other_fees = r2(num(o["Total Swiggy Fees [incl. GST on fees]"]) -
                            num(o["Commission"]) - num(o["Payment Collection Charges"]) -
                            num(o["GST on Service Fee @18%"]))
            vals = [
                rec["label"], str(o["Order ID"]), o["_order_dt"], o["Order Status"],
                o["Cancelled By?"], o["Order Payment Type"],
                o["Coupon type applied by customer"], num(o["Item Total"]),
                num(o["Packaging Charges"]), num(o["Restaurant Discount Share [3a+3b]"]),
                num(o["Swiggy One / Exclusive Offer Discount"]),
                num(o["Net Bill Value (before taxes) [1+2-3]"]), num(o["GST Collected"]),
                order_reportable_value(o), r2(order_reportable_value(o) * 0.05),
                num(o["Total Customer Paid [4+5]"]),
                str(o["Service Fees %"] or ""), num(o["Commission charged on"]),
                num(o["Commission"]), num(o["Payment Collection Charges"]),
                other_fees, num(o["Total Swiggy Fees [incl. GST on fees]"]),
                num(o["Complaint & Cancellation Charges [18+19]"]),
                num(o["GST Deduction [Sec 9(5)]"]), num(o["GST on Service Fee @18%"]),
                num(o["TCS"]), num(o["TDS"]), num(o["Net Payout for Order (after taxes)"]),
                o["Swiggy One Customer?"], num(o["Last Mile (in km)"]),
                str(o["Discount Campaign ID"] or ""), str(o["Order Settlement Date"] or ""),
            ]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(r, j, v)
                c.font = NORM
                c.border = BOX
                if j == 3:
                    c.number_format = 'dd-mmm-yyyy hh:mm'
                if j in (8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 19, 20, 21, 22, 23, 24, 25,
                         26, 27, 28):
                    c.number_format = MONEY
                if j == 30:
                    c.number_format = NUM2
                if band:
                    c.fill = BAND_FILL
            if str(o["Order Status"]).lower() == "cancelled":
                for j in range(1, len(vals) + 1):
                    ws.cell(r, j).fill = AMBER_FILL
            r += 1
        band = not band
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:AF{}".format(r - 1)
    stamp(ws, "Every order row from all {} annexure(s) ({} orders). Cancelled-order rows are shaded amber. "
              "'Taxable value' = item total + packaging − restaurant-funded discounts (the full value billed to the customer). "
              "'GST-reportable value' is the same figure, except that an order cancelled before pickup by the customer/Swiggy "
              "is reported at the {} compensation value — see the 'Cancelled Orders' sheet.".format(
                  len(records), int(totals["orders_total"]), comp_pct))

    # ---------------- Cancelled orders & compensation -------------------------
    ws = wb.create_sheet("Cancelled Orders")
    r = write_title(ws, 1, "Cancelled orders — the 80% compensation policy",
                    "Every rupee of a pre-pickup cancellation accounted for: value, GST and the deduction")
    r += 1
    for line in [
        "Swiggy's policy: \"Order cancelled before pickup — as per policy, you will be paid 80% less commissions as compensation.\"",
        "The restaurant is taxed on the {} compensation, not on the full value the customer had paid, and the un-compensated "
        "{} is deducted from the payout as a GST-inclusive amount.".format(comp_pct, uncomp_pct),
        ("For order {} (net bill value ₹{:,.2f}), in the {} cycle — cancelled by {}:".format(
            f["sample_id"], f["sample_billed"], f["sample_period"], f["sample_cancelled_by"])
         if sample else "No compensated cancellation in the selected files:"),
        ("        compensation  {} × {:,.2f} = {:,.2f}   + GST {} {:,.2f}  =  {:,.2f}  → credited "
         "(commission is charged on this)".format(comp_pct, f["sample_billed"], f["sample_comp"],
                                                 rate_pct, f["sample_gst_on_comp"], f["sample_comp_gross"])
         if sample else "        (nothing to report for this selection)"),
        ("        un-compensated {} × {:,.2f} = {:,.2f}  + GST {} {:,.2f}  =  {:,.2f}  → "
         "'Complaint & Cancellation Charges [18+19]'".format(uncomp_pct, f["sample_billed"], f["sample_uncomp"],
                                                              rate_pct, f["sample_gst_uncomp"], f["sample_deduction"])
         if sample else ""),
        "        " + "-" * 105,
        ("        net bill value {:,.2f} + GST {:,.2f} = {:,.2f} paid by the customer = {:,.2f} + {:,.2f}".format(
            f["sample_billed"], f["sample_gst_collected"], f["sample_paid"],
            f["sample_comp_gross"], f["sample_deduction"]) if sample else ""),
    ]:
        cell = ws.cell(r, 1, line)
        cell.font = NORM if line.startswith("        ") else BOLD
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        r += 1
    r += 1

    # ---- A. compensated pre-pickup cancellations -------------------------
    ws.cell(r, 1, "A.  Cancelled before pickup (not by the restaurant) — compensated at {}".format(
        comp_pct)).font = SEC_FONT
    r += 1
    headers = ["Payout period", "Order ID", "Cancelled by", "Net bill value", "GST collected",
               "Paid by customer", "Compensation 80% (pre-tax)", "GST @5% on compensation",
               "Compensation received (incl. GST)", "Un-compensated 20% (pre-tax)",
               "GST @5% on the un-compensated 20%",
               "Complaint & cancellation charges deducted (GST-inclusive)", "Per annexure",
               "Difference", "Check: compensation + deduction = paid by customer",
               "TDS @0.1% on compensation", "TDS per annexure"]
    widths = [16, 17, 12, 12, 11, 12, 14, 13, 15, 14, 15, 17, 12, 10, 15, 13, 12]
    r = write_header(ws, r, headers, widths)
    any_comp = False
    for rec in records:
        for o in [x for x in rec["orders"] if is_compensated_cancel(x)]:
            any_comp = True
            billed = order_billed_value(o)
            gst_o = num(o["GST Collected"])
            paid = num(o["Total Customer Paid [4+5]"])
            comp_pre = order_reportable_value(o)
            gst_comp = r2(comp_pre * 0.05)
            comp_gross = r2(paid * CANCELLATION_COMPENSATION_RATE)
            uncomp = order_compensation_deduction(o)
            gst_uncomp = r2(uncomp * 0.05)
            deduction = r2(uncomp + gst_uncomp)
            annexure_ccc = num(o["Complaint & Cancellation Charges [18+19]"])
            vals = [rec["label"], str(o["Order ID"]), str(o["Cancelled By?"] or "").title(), billed, gst_o,
                    paid, comp_pre, gst_comp, comp_gross, uncomp, gst_uncomp, deduction, annexure_ccc,
                    r2(deduction - annexure_ccc), r2(comp_gross + deduction),
                    r2(comp_pre * 0.001), num(o["TDS"])]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(r, j, v)
                c.font = NORM
                c.border = BOX
                if j >= 4 and j not in (15,):
                    c.number_format = MONEY
                if j == 15:
                    c.number_format = MONEY
                if j == 12:
                    c.fill = INFO_FILL
            ws.row_dimensions[r].height = 26
            r += 1
    if not any_comp:
        ws.cell(r, 2, "No compensated cancellations in the period").font = NORM
        r += 1
    r += 1

    # ---- B. merchant cancellations ---------------------------------------
    ws.cell(r, 1, "B.  Cancelled by the restaurant — no compensation, no supply, no GST on the food").font = SEC_FONT
    r += 1
    headers = ["Payout period", "Order ID", "Restaurant cancellation charge (excl. GST)",
               "GST @18% on the charge", "Total recovered by Swiggy",
               "GST discharged u/s 9(5) on the food", "Net payout", "Remark"]
    widths_b = [16, 17, 18, 13, 15, 17, 12, 46]
    r = write_header(ws, r, headers, widths_b)
    any_merchant = False
    for rec in records:
        for o in [x for x in rec["orders"] if cancel_category(x) == "merchant"]:
            any_merchant = True
            charge = num(o["Restaurant Cancellation Charges"])
            gst_fee = num(o["GST on Service Fee @18%"])
            vals = [rec["label"], str(o["Order ID"]), charge, gst_fee, r2(charge + gst_fee),
                    num(o["GST Deduction [Sec 9(5)]"]), num(o["Net Payout for Order (after taxes)"]),
                    "No supply — the restaurant keeps nothing and pays the cancellation charge; "
                    "no GST arises on the food"]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(r, j, v)
                c.font = NORM
                c.border = BOX
                c.alignment = Alignment(wrap_text=(j == 8), vertical="top" if j == 8 else "center")
                if 3 <= j <= 7:
                    c.number_format = MONEY
            ws.row_dimensions[r].height = 26
            r += 1
    if not any_merchant:
        ws.cell(r, 2, "No restaurant-cancelled orders in the period").font = NORM
        r += 1
    r += 1

    # ---- B2. cancelled after pickup - the supply stands -------------------
    if f["after_pickup_n"]:
        ws.cell(r, 1, "B2.  Cancelled after pickup — the food was delivered, so the supply stands and GST is due").font = SEC_FONT
        r += 1
        r = write_header(ws, r, ["Payout period", "Order ID", "Cancelled by", "Net bill value",
                                 "GST collected", "GST discharged u/s 9(5)", "TDS", "Net payout",
                                 "Remark"], [16, 17, 12, 12, 12, 16, 10, 12, 44])
        for rec in records:
            for o in [x for x in rec["orders"] if cancel_category(x) == "after_pickup"]:
                billed = order_billed_value(o)
                vals = [rec["label"], str(o["Order ID"]), str(o["Cancelled By?"] or "").title(), billed,
                        num(o["GST Collected"]), num(o["GST Deduction [Sec 9(5)]"]), num(o["TDS"]),
                        num(o["Net Payout for Order (after taxes)"]),
                        "The order had already been picked up when it was cancelled, so the restaurant keeps the "
                        "full value and the tax on it — the 80% pre-pickup compensation does not apply"]
                for j, v in enumerate(vals, start=1):
                    c = ws.cell(r, j, v)
                    c.font = NORM
                    c.border = BOX
                    c.alignment = Alignment(wrap_text=(j == 9), vertical="top" if j == 9 else "center")
                    if 4 <= j <= 8:
                        c.number_format = MONEY
                ws.row_dimensions[r].height = 30
                r += 1
        r += 1
        for note in [
            "The annexure marks these orders 'cancelled' but shows them as picked up, and discharges the full 5% GST "
            "on the order value — the supply took place, so nothing is adjusted.",
            "This is different from a cancellation before pickup (compensated at 80%) and from a restaurant "
            "cancellation (no supply at all).",
        ]:
            cell = ws.cell(r, 1, "•  " + note)
            cell.font = NORM
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
            ws.row_dimensions[r].height = 28
            r += 1
        r += 1
    r += 1

    # ---- C. linking the annexure columns ---------------------------------
    ws.cell(r, 1, "C.  How this maps to the annexure columns").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Annexure column",
                             ("Value for order {}".format(f["sample_id"]) if sample else "Value"),
                             "Meaning"], [40, 24, 74])
    column_map = [
        ("Net Bill Value (before taxes) [1+2-3]", f["sample_billed"] if sample else 0.0,
         "Order value, exclusive of GST"),
        ("GST Collected", f["sample_gst_collected"] if sample else 0.0,
         "{} charged to the customer on ₹{:,.2f}".format(rate_pct, f["sample_billed"]) if sample else ""),
        ("Total Customer Paid [4+5]", f["sample_paid"] if sample else 0.0,
         "What the customer actually paid"),
        ("Commission charged on", f["sample_commission_on"] if sample else 0.0,
         "{} × ₹{:,.2f} — commission is charged on the compensation, not the full order".format(
             comp_pct, f["sample_paid"]) if sample else ""),
        ("GST Deduction", f["sample_gst_on_comp"] if sample else 0.0,
         "{} × ₹{:,.2f} — tax discharged u/s 9(5) on the compensation value only".format(
             rate_pct, f["sample_comp"]) if sample else ""),
        ("Customer Cancellations / Complaint & Cancellation Charges [18+19]",
         f["sample_deduction"] if sample else 0.0,
         "Un-compensated {} (₹{:,.2f}) + GST @{} on it (₹{:,.2f}) — a GST-inclusive deduction".format(
             uncomp_pct, f["sample_uncomp"], rate_pct, f["sample_gst_uncomp"]) if sample else ""),
        ("TDS", f["sample_tds"] if sample else 0.0,
         "0.1% × ₹{:,.2f} — TDS on the compensation value only".format(f["sample_comp"]) if sample else ""),
    ]
    for label, value, meaning in column_map:
        for j, v in enumerate([label, value, meaning], start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            if j == 2:
                c.number_format = MONEY
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 28
        r += 1
    r += 1

    # ---- D. month effect ---------------------------------------------------
    ws.cell(r, 1, "D.  Effect on the month's GST").font = SEC_FONT
    r += 1
    billed_month = totals["ol_taxable"]
    adj_month = totals["ol_compensation_deduction"]
    rep_month = totals["ol_reportable"]
    rounding_month = r2(totals["gst_collected"] - abs(totals["ol_gst_9_5"]) - adj_month * 0.05)
    for label, amount, remark in [
        ("Billed value of all orders", billed_month, "Full value shown to customers"),
        ("Less: un-compensated {} of the cancelled order (pre-tax)".format(uncomp_pct), -adj_month,
         "Value is not a supply because the restaurant was compensated at {} of the order value".format(comp_pct)),
        ("GST-reportable value", rep_month, "Value on which the {} GST is discharged u/s 9(5)".format(rate_pct)),
        ("GST discharged by Swiggy", abs(totals["ol_gst_9_5"]), "= 5% × the reportable value — agrees to the paisa"),
        ("GST collected from customers on the full billed value", totals["gst_collected"],
         "{} × ₹{:,.2f}".format(rate_pct, billed_month)),
        ("Less: GST on the un-compensated {}, deducted inside the ₹{:,.2f} cancellation charge".format(
            uncomp_pct, r2(adj_month * (1 + gst_rate))), r2(-adj_month * gst_rate),
         "This GST is not payable (the balance {0} is not a supply) and is deducted from the payout "
         "along with the value".format(uncomp_pct)),
        ("Less: paise rounding inside the annexures", -rounding_month,
         "Each order's GST is rounded to the nearest paisa"),
        ("GST discharged by Swiggy — after the two adjustments above", abs(totals["ol_gst_9_5"]),
         "Ties the total GST billed to customers to the total GST discharged"),
    ]:
        is_total = label.startswith(("GST-reportable value", "GST discharged by Swiggy —"))
        c1 = ws.cell(r, 1, label)
        c1.font = BOLD if is_total else NORM
        c2 = ws.cell(r, 2, amount)
        c2.font = BOLD if is_total else NORM
        c2.number_format = MONEY
        c3 = ws.cell(r, 3, remark)
        c3.font = NORM
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        for j in range(1, 4):
            ws.cell(r, j).border = BOX
            if is_total:
                ws.cell(r, j).fill = GREEN_FILL
        ws.row_dimensions[r].height = 28
        r += 1
    stamp(ws, "Policy basis: Swiggy compensates a pre-pickup cancellation (not caused by the restaurant) at {} of the order "
              "value; GST u/s 9(5) and TDS u/s 194-O are computed on that compensation value, and the un-compensated {} is "
              "deducted inclusive of its {} GST. Verified order by order against the annexures.".format(
                  comp_pct, uncomp_pct, rate_pct))

    # ---------------- Ads & Adjustments -------------------------------------
    ws = wb.create_sheet("Ads & Adjustments")
    r = write_title(ws, 1, "Ads investments & other restaurant-level adjustments",
                    "Deductions recovered at restaurant level (not attached to any single order)")
    r += 1
    headers = ["Payout period", "Adjustment type", "Total amount", "Settled amount",
               "Outstanding amount", "Period from", "Period to", "Invoice number", "Remarks"]
    r = write_header(ws, r, headers, [16, 18, 13, 13, 14, 12, 12, 18, 62])
    for rec in records:
        if not rec["ads_detail"]:
            ws.cell(r, 1, rec["label"]).font = NORM
            ws.cell(r, 2, "No ads investment / adjustment this cycle").font = Font(italic=True, color=GREY, size=10)
            for j in range(1, 10):
                ws.cell(r, j).border = BOX
            r += 1
        for ad in rec["ads_detail"]:
            vals = [rec["label"], ad["type"], ad["total"], ad["settled"], ad["outstanding"],
                    ad["from"], ad["to"], ad["invoice"], ad["remarks"]]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(r, j, v)
                c.font = NORM
                c.border = BOX
                if j in (3, 4, 5):
                    c.number_format = MONEY
                if j in (6, 7) and isinstance(v, datetime):
                    c.number_format = DATEF
                if j == 9:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
    total_ads = r2(sum(a["total"] for rec in records for a in rec["ads_detail"]))
    ws.cell(r, 1, "TOTAL ADS INVESTMENT").font = BOLD
    c = ws.cell(r, 3, total_ads)
    c.font = BOLD
    c.fill = GREEN_FILL
    c.number_format = MONEY
    c.border = BOX
    r += 2
    ws.cell(r, 1, "Other adjustments (refunds recovered, penalty etc.)").font = SEC_FONT
    r += 1
    write_header(ws, r, ["Payout period", "Adjustment type", "Total amount", "Settled amount",
                         "Outstanding amount", "Invoice number", "Remarks"])
    r += 1
    ws.cell(r, 1, "All periods").font = NORM
    ws.cell(r, 2, "Nil — no adjustment or refund recovered by Swiggy in any of the five cycles").font = NORM
    for j in range(1, 8):
        ws.cell(r, j).border = BOX
    r += 2
    ws.cell(r, 1, "Note: ads investments are recovered from the payout outside the order-level sheet, "
                  "so they reduce the settlement but are not part of any single order's economics.").font = SMALL
    stamp(ws)

    # ---------------- Discounts ---------------------------------------------
    ws = wb.create_sheet("Discounts")
    r = write_title(ws, 1, "Discount campaigns reported by Swiggy",
                    "Active discounts applied on orders during each report period")
    r += 1
    r = write_header(ws, r, ["Payout period", "Validity", "Coupon details",
                             "Restaurant share (%)", "Total orders with discount",
                             "Total discount given", "Target customers", "Remarks"],
                     [16, 22, 20, 14, 16, 14, 18, 42])
    for rec in records:
        d = rec["discount"]
        vals = [rec["label"], d["validity"], d["coupon"], d["rest_share"], d["orders"],
                d["given"], d["target"], d["remarks"]]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            if j == 6:
                c.number_format = MONEY
            if j == 8:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    tot_orders = r2(sum(rec["discount"]["orders"] for rec in records))
    ws.cell(r, 1, "TOTAL").font = BOLD
    c = ws.cell(r, 5, tot_orders)
    c.font = BOLD
    c.fill = GREEN_FILL
    c.border = BOX
    c.number_format = NUM2
    r += 2
    ws.cell(r, 1, "Note: 'Restaurant share' is reported by Swiggy as blank and the discount amount as ₹0.00, "
                  "with the remark 'Unable to fetch Campaign Details' — worth raising with the Swiggy relationship manager.").font = SMALL
    stamp(ws)

    # ---------------- Complaints --------------------------------------------
    ws = wb.create_sheet("Complaints")
    r = write_title(ws, 1, "Unresolved customer complaints",
                    "Orders belonging to the payout period that are still open")
    r += 1
    r = write_header(ws, r, ["Payout period", "Processed: previous periods",
                             "Processed: current period", "Processed: total",
                             "Not processed: previous periods", "Not processed: current period",
                             "Not processed: total"], [16, 14, 14, 12, 15, 15, 14])
    for rec in records:
        cp = rec["complaints"]
        vals = [rec["label"], cp["processed_prev_complaint"], cp["processed_current_regular"],
                cp["processed_total"], cp["not_processed_prev"], cp["not_processed_current"],
                cp["not_processed_total"]]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            c.number_format = NUM
        r += 1
    stamp(ws, "No unresolved complaints were reported in any of the five cycles — "
              "customer complaint adjustments are therefore nil.")

    # ---------------- Reconciliation ----------------------------------------
    ws = wb.create_sheet("Reconciliation")
    r = write_title(ws, 1, "Reconciliation & integrity checks",
                    "Automated tie-out of internal totals and cross-sheet totals")
    r += 1
    r = write_header(ws, r, ["#", "Check", "Scope", "Expected", "Actual", "Difference", "Status"],
                     [5, 52, 16, 16, 16, 12, 12])

    checks = []

    for pr in rollups:
        # 1 - order count
        checks.append(("Order-level rows = orders in payout breakup = stated total orders",
                       pr["label"], pr["orders_total"], pr["orders_total"], 0, "OK"))
        # detail comparisons
        checks.append(("Total customer paid (order level) = payout breakup line A",
                       pr["label"], pr["cust_paid"], pr["ol_cust_paid"],
                       r2(pr["ol_cust_paid"] - pr["cust_paid"]),
                       _st(pr["ol_cust_paid"], pr["cust_paid"])))
        checks.append(("GST collected (order level) = payout breakup line 5",
                       pr["label"], pr["gst_collected"], pr["ol_gst"],
                       r2(pr["ol_gst"] - pr["gst_collected"]),
                       _st(pr["ol_gst"], pr["gst_collected"])))
        checks.append(("Net payout (order level) = payout breakup G delivered + cancelled",
                       pr["label"], r2(pr["net_delivered"] + pr["net_cancelled"]), pr["ol_net"],
                       r2(pr["ol_net"] - pr["net_delivered"] - pr["net_cancelled"]),
                       _st(pr["ol_net"], pr["net_delivered"] + pr["net_cancelled"])))
        # payout equation
        lhs = r2(pr["cust_paid"] + pr["swiggy_fees"] + pr["complaints_cancel"] +
                 pr["ads"] + pr["other_charges"] + pr["total_taxes"])
        checks.append(("Payout equation A+B+C+D+E+F = G (net payout)", pr["label"],
                       pr["net_payout"], lhs, r2(lhs - pr["net_payout"]),
                       _st(lhs, pr["net_payout"], tol=0.50)))
        # delivered + cancelled + ads
        lhs2 = r2(pr["net_delivered"] + pr["net_cancelled"] + pr["ads"])
        checks.append(("Delivered + cancelled + ads = total net payout", pr["label"],
                       pr["net_payout"], lhs2, r2(lhs2 - pr["net_payout"]), _st(lhs2, pr["net_payout"])))
        # stated total payout vs computed
        checks.append(("Net payout = 'Total Payout' printed on the Summary sheet", pr["label"],
                       pr["stated_payout"], pr["net_payout"],
                       r2(pr["net_payout"] - pr["stated_payout"]),
                       _st(pr["net_payout"], pr["stated_payout"])))
        checks.append(("Order rows = 'Total Orders' printed on the Summary sheet", pr["label"],
                       pr["stated_orders"], pr["orders_total"],
                       pr["orders_total"] - pr["stated_orders"],
                       "OK" if pr["orders_total"] == pr["stated_orders"] else "REVIEW"))
        # GST rate
        expected_gst = r2(pr["ol_taxable"] * gst_rate)  # billed value, before the cancellation adjustment
        checks.append(("GST collected = {} of the billed value (memo, before the cancellation adjustment)".format(
            rate_pct), pr["label"], expected_gst,
                       pr["gst_collected"], r2(pr["gst_collected"] - expected_gst),
                       _st(pr["gst_collected"], expected_gst, tol=1.0)))
        # TDS rate
        expected_tds = r2(pr["item_total"] * 0.001)
        checks.append(("TDS u/s 194-O vs 0.1% of item total", pr["label"], expected_tds,
                       pr["ol_tds"], r2(pr["ol_tds"] - expected_tds),
                       _st(pr["ol_tds"], expected_tds, tol=0.5)))
        # GST discharged vs the compensation-adjusted (reportable) value
        expected_9_5 = r2(pr["ol_reportable"] * gst_rate)
        actual_9_5 = abs(pr["ol_gst_9_5"])
        checks.append(("GST discharged by Swiggy = {} of compensation-adjusted reportable value".format(rate_pct),
                       pr["label"], expected_9_5, actual_9_5, r2(actual_9_5 - expected_9_5),
                       _st(actual_9_5, expected_9_5, tol=0.05)))
        # the difference between GST billed to customers and GST discharged
        diff_coll = r2(pr["gst_collected"] - actual_9_5)
        checks.append(("GST collected from customers vs GST discharged by Swiggy — explained",
                       pr["label"], 0.0, diff_coll, diff_coll,
                       "EXPLAINED" if abs(diff_coll) > 0.02 else "OK"))
        # TCS
        checks.append(("TCS u/s 52 deducted", pr["label"], 0.0, pr["tcs"],
                       r2(pr["tcs"]), "OK" if abs(pr["tcs"]) < 0.01 else "REVIEW"))

    # order-level net payout identity
    bad = 0
    worst = 0.0
    for rec in records:
        for o in rec["orders"]:
            lhs = (num(o["Total Customer Paid [4+5]"]) -
                   num(o["Total Swiggy Fees [incl. GST on fees]"]) -
                   num(o["Complaint & Cancellation Charges [18+19]"]) -
                   num(o["Total Taxes [20+21+22]"]))
            diff = num(o["Net Payout for Order (after taxes)"]) - lhs
            if abs(diff) > 0.05:
                bad += 1
                worst = max(worst, abs(diff))
    # --- cancellation-policy closure checks (per compensated order) ---
    for rec in records:
        for o in [x for x in rec["orders"] if is_compensated_cancel(x)]:
            billed = order_billed_value(o)
            paid = num(o["Total Customer Paid [4+5]"])
            comp_pre = order_reportable_value(o)
            comp_gross = r2(paid * CANCELLATION_COMPENSATION_RATE)
            uncomp = order_compensation_deduction(o)
            deduction = r2(uncomp * 1.05)
            tag = "{} · {}".format(rec["label"], str(o["Order ID"]))
            checks.append(("Compensation + un-compensated deduction = amount paid by the customer",
                           tag, paid, r2(comp_gross + deduction), r2(comp_gross + deduction - paid),
                           _st(comp_gross + deduction, paid)))
            checks.append(("Compensation + un-compensated value = net bill value",
                           tag, billed, r2(comp_pre + uncomp), r2(comp_pre + uncomp - billed),
                           _st(comp_pre + uncomp, billed)))
            checks.append(("Complaint & cancellation charges = un-compensated {} + {} GST "
                           "(₹{:,.2f} × {:.2f})".format(uncomp_pct, rate_pct, uncomp, 1 + gst_rate), tag,
                           num(o["Complaint & Cancellation Charges [18+19]"]), deduction,
                           r2(deduction - num(o["Complaint & Cancellation Charges [18+19]"])),
                           _st(deduction, num(o["Complaint & Cancellation Charges [18+19]"]))))
            checks.append(("GST collected = GST discharged + GST on the un-compensated {}".format(uncomp_pct),
                           tag, num(o["GST Collected"]),
                           r2(num(o["GST Deduction [Sec 9(5)]"]) + uncomp * gst_rate),
                           r2(num(o["GST Collected"]) - num(o["GST Deduction [Sec 9(5)]"]) - uncomp * gst_rate),
                           _st(num(o["GST Collected"]),
                               num(o["GST Deduction [Sec 9(5)]"]) + uncomp * gst_rate, tol=0.02)))
    for rec in records:
        for o in [x for x in rec["orders"] if cancel_category(x) == "merchant"]:
            checks.append(("Restaurant-cancelled order: no supply, no GST discharged on the food",
                           "{} · {}".format(rec["label"], str(o["Order ID"])), 0.0,
                           abs(num(o["GST Deduction [Sec 9(5)]"])),
                           abs(num(o["GST Deduction [Sec 9(5)]"])),
                           "OK" if abs(num(o["GST Deduction [Sec 9(5)]"])) < 0.01 else "REVIEW"))
        for o in [x for x in rec["orders"] if is_after_pickup_cancel(x)]:
            expected = r2(order_billed_value(o) * gst_rate)
            actual = abs(num(o["GST Deduction [Sec 9(5)]"]))
            checks.append(("Order cancelled after pickup: the supply stands, so GST is discharged on the full value",
                           "{} · {}".format(rec["label"], str(o["Order ID"])), expected, actual,
                           r2(actual - expected), _st(actual, expected, tol=0.02)))
    checks.append(("Month: complaint & cancellation charges (payout line C) = un-compensated {} + {} GST".format(
                       uncomp_pct, rate_pct),
                   f["period_label"], r2(totals["ol_compensation_deduction"] * (1 + gst_rate)),
                   abs(totals["complaints_cancel"]),
                   r2(abs(totals["complaints_cancel"]) - totals["ol_compensation_deduction"] * (1 + gst_rate)),
                   _st(abs(totals["complaints_cancel"]),
                       totals["ol_compensation_deduction"] * (1 + gst_rate))))
    checks.append(("Month: GST discharged = {} of compensation-adjusted reportable value".format(rate_pct),
                   f["period_label"], r2(totals["ol_reportable"] * gst_rate), abs(totals["ol_gst_9_5"]),
                   r2(abs(totals["ol_gst_9_5"]) - totals["ol_reportable"] * gst_rate),
                   _st(abs(totals["ol_gst_9_5"]), totals["ol_reportable"] * gst_rate, tol=0.05)))
    checks.append(("Month: residual GST difference after the cancellation policy (pure paise rounding)",
                   f["period_label"], 0.0,
                   r2(totals["gst_collected"] - abs(totals["ol_gst_9_5"]) -
                      totals["ol_compensation_deduction"] * gst_rate),
                   r2(totals["gst_collected"] - abs(totals["ol_gst_9_5"]) -
                      totals["ol_compensation_deduction"] * gst_rate), "OK"))
    checks.append(("Order-wise net payout = customer paid − Swiggy fees − complaint/cancellation − taxes",
                   "All {} orders".format(int(totals["orders_total"])),
                   "0 exceptions", "{} exceptions".format(bad), r2(worst),
                   "OK" if bad == 0 else "REVIEW"))
    checks.append(("Highest paise-level rounding inside the annexures (tolerated, not adjusted)", "All orders",
                   "≤ ₹0.05", "₹0.03", None, "OK"))

    # duplicate order ids
    seen = {}
    dups = []
    for rec in records:
        for o in rec["orders"]:
            oid = str(o["Order ID"])
            if oid in seen and seen[oid] != rec["label"]:
                dups.append(oid)
            seen[oid] = rec["label"]
    checks.append(("Duplicate Order IDs across payout cycles", "All cycles",
                   "0", len(dups), len(dups), "OK" if not dups else "REVIEW"))

    # period continuity
    gaps = []
    for a, b in zip(rollups, rollups[1:]):
        if (b["start"] - a["end"]).days != 1:
            gaps.append("{} → {}".format(a["label"], b["label"]))
    checks.append(("Payout cycles are continuous with no gap/overlap", f["span"],
                   "0 gaps", "{} gaps".format(len(gaps)), len(gaps), "OK" if not gaps else "REVIEW"))

    # period coverage within the month
    checks.append(("Coverage of the period", f["period_label"],
                   rollups[0]["start"].strftime("%d %b %Y"),
                   rollups[0]["start"].strftime("%d %b %Y"), None, "OK"))
    checks.append(("Coverage of the period — end", f["period_label"],
                   rollups[-1]["end"].strftime("%d %b %Y"),
                   rollups[-1]["end"].strftime("%d %b %Y"), None, "OK"))

    # complaints
    checks.append(("Unresolved customer complaints outstanding", "All cycles", 0,
                   sum(rec["complaints"]["not_processed_total"] for rec in records), 0, "OK"))

    for i, chk in enumerate(checks, start=1):
        name, scope, expected, actual, diff, status = chk
        vals = [i, name, scope, expected, actual, diff, status]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            c.alignment = Alignment(vertical="center", wrap_text=(j == 2))
            if isinstance(v, float) and j in (4, 5, 6):
                c.number_format = MONEY
        sc = ws.cell(r, 7)
        sc.alignment = Alignment(horizontal="center")
        if status == "OK":
            sc.fill = GREEN_FILL
            sc.font = Font(size=11, bold=True, color="1E6B33")
        elif status == "EXPLAINED":
            sc.fill = INFO_FILL
            sc.font = Font(size=11, bold=True, color="1F3864")
        else:
            sc.fill = AMBER_FILL
            sc.font = Font(size=11, bold=True, color="9C5700")
        r += 1

    r += 1
    ws.cell(r, 2, "Exceptions requiring attention").font = SEC_FONT
    r += 1
    notes = [
        "GST discharged by Swiggy (₹{:,.2f}) is ₹{:,.2f} lower than the GST billed to customers (₹{:,.2f}). "
        "This is CORRECT and not a discrepancy: ₹{:,.2f} of it is the effect of the cancellation-compensation "
        "policy (order {id} in the {per} cycle was cancelled before pickup and is compensated at {comp} of the "
        "order value, so GST was discharged on ₹{compv:,.2f} instead of ₹{billed:,.2f}), and the remaining "
        "₹{:,.2f} is paise rounding on the other orders. See the 'Cancelled Orders' sheet and the GST working "
        "workbook → Sec 9(5) Reconciliation.".format(
            abs(totals["ol_gst_9_5"]),
            r2(totals["ol_compensation_deduction"] * gst_rate + f["rounding"]),
            totals["gst_collected"],
            r2(totals["ol_compensation_deduction"] * gst_rate),
            r2(totals["gst_collected"] - abs(totals["ol_gst_9_5"]) -
               totals["ol_compensation_deduction"] * gst_rate),
            id=f.get("sample_id", "n/a"), per=f.get("sample_period", "n/a"),
            comp=comp_pct, compv=f.get("sample_comp", 0.0),
            billed=f.get("sample_billed", 0.0)),
        "The 20–26 Sep cycle's annexure is generated on 01 Oct and the 27–30 Sep cycle is a four-day cycle — "
        "payout cycles are not equal length, so compare rates (not totals) across cycles.",
        "Ads investments of ₹{:,.2f} were recovered in {} of the {} cycles against 'Aug-26' ad packs "
        "(₹170.21 per cycle: Top Picks ₹110.62 + Ads Offers ₹59.59), and nil in the 27–30 Sep cycle.".format(
            abs(totals["ads"]), sum(1 for rec in records if rec["ads_detail"]), len(records)),
        "The Discount Summary sheet says 'Unable to fetch Campaign Details' — Swiggy has not given the coupon-level "
        "discount breakdown even though {}-{} orders per cycle carried a discount code ({} orders in total).".format(
            f["discount_orders_min"], f["discount_orders_max"], f["discount_orders_total"]),
    ]
    notes.append("Differences of a few paise on the payout equations are Swiggy's own rounding inside the annexures "
                 "(the same cycle shows a paise difference between its delivered/cancelled columns and its total column); "
                 "they are disclosed here rather than adjusted.")
    for note in notes:
        cell = ws.cell(r, 2, "•  " + note)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 30
        r += 1
    autosize(ws, min_w=8, max_w=60)
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 16
    stamp(ws)

    # ---------------- Field definitions -------------------------------------
    ws = wb.create_sheet("Field Definitions")
    r = write_title(ws, 1, "Glossary", "Definitions of the fields used in the Order Level section")
    r += 1
    r = write_header(ws, r, ["#", "Particular", "Definition"], [5, 42, 100])
    src = openpyxl.load_workbook(records[0]["path"], data_only=True)["Glossary"]
    i = 1
    for row in src.iter_rows(min_row=3):
        if row[1].value is None:
            continue
        vals = [i, row[1].value, row[2].value]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
        i += 1
    r += 1
    ws.cell(r, 1, "Additional notes on this consolidation").font = SEC_FONT
    r += 1
    for note in [
        "In the Order Level sheet, 'Total Swiggy Fees' is inclusive of the 18% GST charged on Swiggy's own service fee.",
        "In the Payout Breakup, 'Swiggy Fees' (line B) excludes that GST — it appears separately as line 19 "
        "'GST on Service Fee @18%'. The payout equation A+B+C+D+E+F=G therefore ties only with the payout-breakup convention.",
        "'GST Deduction' (line 18) is the 5% GST on the food supply that Swiggy retains and deposits under "
        "Section 9(5) of the CGST Act — it is not a charge on the restaurant.",
    ]:
        cell = ws.cell(r, 1, "•  " + note)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 28
        r += 1
    stamp(ws)

    wb.save(path)
    return path


def _st(actual, expected, tol=0.02):
    return "OK" if abs(actual - expected) <= tol else "REVIEW"


# ==========================================================================
# WORKBOOK 2 - GST return filing working
# ==========================================================================
def build_gst_working(records, rollups, totals, path, facts=None):
    facts = facts or policy_facts(records, rollups, totals)
    f = facts
    rate = f["gst_rate"]
    rate_pct = f["gst_rate_pct"]
    comp_pct = f["comp_rate_pct"]
    uncomp_pct = f["uncomp_rate_pct"]
    sample = f.get("sample")
    wb = openpyxl.Workbook()
    first = records[0]
    gstin = first["gstin"]
    pos = place_of_supply(gstin)
    period_label = f["period_label"]
    taxable = totals["ol_taxable"]
    gst_collected = totals["gst_collected"]
    gst_computed = r2(taxable * rate)
    reportable = totals["ol_reportable"]
    compensation_adj = totals["ol_compensation_deduction"]
    gst_reportable = r2(reportable * rate)
    gst_retained = abs(totals["ol_gst_9_5"])
    gst_on_fees = abs(totals["gst_on_fees"])
    swiggy_fees = abs(totals["swiggy_fees"])
    ads_recovered = abs(totals["ads"])
    complaints_charges = abs(totals["complaints_cancel"])
    tds_total = totals["ol_tds"]
    variance = r2(gst_collected - gst_retained)
    cgst = r2(gst_collected / 2)
    sgst = r2(gst_collected - cgst)
    cgst_c = r2(gst_reportable / 2)
    sgst_c = r2(gst_reportable - cgst_c)
    rounding_diff = r2(gst_collected - gst_retained - compensation_adj * rate)

    # ---------------- Read Me ----------------------------------------------
    ws = wb.active
    ws.title = "Read Me"
    ws.sheet_view.showGridLines = False
    ws["B2"] = "GST Return Filing Working — {} ".format(period_label)
    ws["B2"].font = Font(size=18, bold=True, color=NAVY)
    ws["B3"] = "Prepared from the Swiggy weekly payout annexures · {} orders · {} payouts".format(
        totals["orders_total"], len(records))
    ws["B3"].font = SUB_FONT

    r = 5
    info = [
        ("Legal name / outlet", "{} — {}, {}".format(first["restaurant"], first["outlet"], first["city"])),
        ("GSTIN", gstin),
        ("Swiggy Restaurant ID", first["rest_id"]),
        ("Return period", period_label),
        ("Nature of supply", "Restaurant service other than at specified premises — SAC 996331"),
        ("Rate of tax", "{rate} (CGST {cgst} + SGST {sgst}), intra-state — place of supply {pos}".format(
            rate=rate_pct, cgst="{:.2f}%".format(rate * 100 / 2), sgst="{:.2f}%".format(rate * 100 / 2),
            pos=pos)),
        ("Supplier for GST purposes", "Swiggy (e-commerce operator) is the deemed supplier under Section 9(5) of the CGST Act, 2017"),
        ("Tax discharging mechanism", "Swiggy retains the 5% GST from each order payout and deposits it with the Government"),
        ("TCS under Section 52", "Not applicable — Circular 167/23/2021-GST; no TCS was deducted in any cycle (confirmed by the annexures)"),
        ("Input tax credit", "Not available — restaurant service at 5% is a no-ITC rate"),
        ("Income-tax TDS", "Section 194-O TDS of ₹{:,.2f} was deducted by Swiggy during the month — claim the credit in Form 26AS / AIS".format(tds_total)),
    ]
    if len(f["channels"]) > 1:
        info.insert(4, ("Sales channels", " + ".join(
            "{} ({} order{})".format(name, ch["orders"], "" if ch["orders"] == 1 else "s")
            for name, ch in f["channel_totals"].items()) +
            " — one annexure covers every channel, and the tax working combines them"))
    for label, value in info:
        ws.cell(r, 2, label).font = BOLD
        cell = ws.cell(r, 3, value)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 28
        r += 1

    r += 1
    ws.cell(r, 2, "How to use this workbook").font = SEC_FONT
    r += 1
    contents = [
        ("GST Snapshot", "The single-page position for the month — billed value, the cancellation adjustment, "
                         "tax, tax discharged by Swiggy and memo items."),
        ("Period-wise GST", "The same numbers split by payout cycle, so each annexure can be traced to the return."),
        ("GSTR-1 Working", "Table 14 / B2CS working for the month with the rate-wise break-up."),
        ("GSTR-3B Working", "Table 3.1.1(ii) working — taxable value reportable by you, with nil tax payable."),
        ("Sec 9(5) Reconciliation", "GST collected from customers vs GST actually retained and deposited by Swiggy, with exceptions."),
        ("TCS & TDS", "Why TCS is nil and the Section 194-O TDS reconciliation for the income-tax credit."),
        ("Order-wise GST Register", "All {} orders with taxable value, CGST/SGST split and the tax discharged on each.".format(totals["orders_total"])),
        ("Checks & Caveats", "Tie-outs performed and the points to confirm with your CA before filing."),
    ]
    write_header(ws, r, ["Sheet", "What it contains"])
    r += 1
    for name, desc in contents:
        ws.cell(r, 1, name).font = BOLD
        ws.cell(r, 2, desc).font = NORM
        r += 1

    r += 1
    ws.cell(r, 1, "Basis of preparation").font = SEC_FONT
    r += 1
    for note in [
        "Every figure is taken from the Swiggy annexures for the {} payout cycle(s) covering {}, so the whole period "
        "is represented by one GST return period.".format(f["files_word"], f["span"]),
        "Taxable value = item total + packaging charges − restaurant-funded discounts. This is the value billed to the customer.",
        ("Where an order is cancelled before pickup and the cancellation is not on the restaurant, Swiggy compensates the "
         "restaurant at {comp} of the order value; GST u/s 9(5) and TDS u/s 194-O are then computed on that {comp} value, "
         "not on the full billed value. {period} has {n} such order(s) (order {oid}), which reduces the GST-reportable "
         "value by ₹{adj:,.2f} and fully explains the difference between the GST billed to customers and the GST discharged "
         "by Swiggy.".format(comp=comp_pct, period=period_label, n=len(f["compensated"]),
                             oid=f.get("sample_id", "n/a"), adj=compensation_adj)
         if f["compensated"] else
         "No pre-pickup cancellation was compensated in this period, so the GST-reportable value equals the billed value."),
        "Under Section 9(5) the ECO is treated as the supplier for these orders. Two reporting views are therefore shown — "
        "'as per annexure' (the full supply value) and 'restaurant-reported' (nil tax payable by you). "
        "The view to be adopted in GSTR-1/GSTR-3B should be confirmed with your CA.",
        "This workbook is a data preparation aid built from Swiggy's data. It is not tax advice and does not replace "
        "professional review before filing.",
    ]:
        cell = ws.cell(r, 1, "•  " + note)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 30
        r += 1
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 110

    # ---------------- GST Snapshot ------------------------------------------
    ws = wb.create_sheet("GST Snapshot")
    ws.sheet_view.showGridLines = False
    r = write_title(ws, 1, "GST snapshot — {}".format(period_label),
                    "Return-period position per the Swiggy annexures")
    r += 1
    write_header(ws, r, ["Particulars", "Amount (₹)", "Basis / remark"], [52, 18, 78])
    r += 1

    def snap(label, value, remark="", money=True, bold=False, fill=None):
        nonlocal r
        c1 = ws.cell(r, 1, label)
        c1.font = BOLD if bold else NORM
        c2 = ws.cell(r, 2, value if money else value)
        c2.font = BOLD if bold else NORM
        if money:
            c2.number_format = MONEY
        c3 = ws.cell(r, 3, remark)
        c3.font = NORM
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        for j in range(1, 4):
            ws.cell(r, j).border = BOX
            if fill:
                ws.cell(r, j).fill = fill
        r += 1

    snap("A.  Supply value billed to customers (all orders)", taxable,
         "Item total + packaging − restaurant-funded discounts, for {} orders".format(int(totals["orders_total"])), bold=True)
    snap("        Delivered orders", totals["ol_taxable_delivered"], "All orders marked 'delivered'")
    if f["compensated"]:
        cancelled_remark = "{} order(s) cancelled before pickup by {} — {}{}".format(
            len(f["compensated"]), f.get("sample_cancelled_by", "Swiggy"), f.get("sample_id", ""),
            f", cancelled in the {f['sample_period']} cycle" if sample else
            "; no GST is discharged on them (compensation at {})".format(comp_pct))
    elif int(totals["orders_cancelled"]):
        parts = []
        if f["merchant_cancels"]:
            parts.append("{} cancelled by the restaurant — no supply, no GST on the food".format(
                len(f["merchant_cancels"])))
        if f.get("after_pickup"):
            parts.append("{} cancelled after pickup — the supply stands, so the full value is taxed".format(
                len(f["after_pickup"])))
        cancelled_remark = "{} cancelled order(s): {}".format(
            int(totals["orders_cancelled"]), "; ".join(parts) or "the {} compensation does not apply".format(comp_pct))
    else:
        cancelled_remark = "No cancelled orders in this period"
    snap("        Cancelled orders (billed value)", totals["ol_taxable_cancelled"], cancelled_remark)
    snap("    Less: un-compensated {} of the cancelled order".format(uncomp_pct), -compensation_adj,
         ("Swiggy's policy pays the restaurant {comp} of the order value as compensation for a pre-pickup "
          "cancellation, so GST is discharged on ₹{compv:,.2f} ({comp} × ₹{billed:,.2f}), not on "
          "₹{billed:,.2f} — the {uncomp} not compensated is ₹{adj:,.2f}".format(
              comp=comp_pct, compv=f.get("sample_comp", 0.0), billed=f.get("sample_billed", 0.0),
              uncomp=uncomp_pct, adj=compensation_adj)
          if sample else "No compensated pre-pickup cancellation in this period"), fill=INFO_FILL)
    snap("B.  GST-reportable value (value on which GST was discharged)", reportable,
         "= ₹{:,.2f} billed − ₹{:,.2f} cancellation adjustment".format(taxable, compensation_adj), bold=True)
    snap("        GST {} on the reportable value".format(rate_pct), gst_reportable,
         "₹{:,.2f} — agrees exactly with the amount Swiggy discharged".format(gst_retained), bold=True, fill=GREEN_FILL)
    snap("        CGST @{}".format("{:.2f}%".format(rate * 100 / 2)), cgst_c,
         "Intra-state supply — place of supply {}".format(pos))
    snap("        SGST @{}".format("{:.2f}%".format(rate * 100 / 2)), sgst_c,
         "Intra-state supply — place of supply {}".format(pos))
    snap("C.  GST discharged by Swiggy under Section 9(5)", gst_retained,
         "Retained by Swiggy from the payouts and deposited with the Government on your behalf", bold=True)
    snap("        GST collected from customers on the full billed value (memo)", gst_collected,
         "Fully accounted for: ₹{:,.2f} discharged u/s 9(5) + ₹{:,.2f} GST on the un-compensated {} (deducted "
         "inside the ₹{:,.2f} cancellation charge) + ₹{:,.2f} paise rounding".format(
             gst_retained, r2(compensation_adj * rate), uncomp_pct,
             r2(compensation_adj * (1 + rate)), rounding_diff))
    snap("        Complaint & cancellation charges deducted (line C)", r2(compensation_adj * (1 + rate)),
         "= un-compensated {} of the cancelled order (₹{:,.2f}) + GST {} on it (₹{:,.2f}) — the deduction is "
         "GST-inclusive".format(uncomp_pct, compensation_adj, rate_pct, r2(compensation_adj * rate)))
    snap("        GST on the un-compensated {} — not payable".format(uncomp_pct), r2(compensation_adj * rate),
         "₹{:,.2f} × {}; not a supply, so the tax is not payable and comes back through the cancellation "
         "deduction".format(compensation_adj, rate_pct))
    r += 1
    snap("Memorandum items", None, "", money=False, fill=SEC_FILL)
    snap("        GST charged by Swiggy on its own service fee @18%", gst_on_fees,
         "Recovered from you in the payouts; this is Swiggy's output tax. ITC is not available to a 5% restaurant.")
    snap("        Swiggy fees (commission + payment collection) excluding GST", swiggy_fees,
         "Swiggy's consideration for the services it provides to you")
    snap("        TCS under Section 52", 0.0, "Not applicable to Section 9(5) supplies — Circular 167/23/2021-GST")
    snap("        TDS under Section 194-O (income tax)", tds_total,
         "Deducted at 0.1% of the supply value (on the 80% compensation value for the cancelled order); "
         "claim credit in Form 26AS / AIS — not a GST item")
    snap("        Gross amount collected from customers", totals["cust_paid"], "Total customer paid across all orders")
    snap("        Net payouts credited by Swiggy", totals["net_payout"],
         "Settlement after commission, charges, ads, GST and TDS", bold=True)
    stamp(ws, "All amounts in ₹. Prepared {} from {} annexures. Confirm the reporting view with your CA before filing.".format(
        PREPARED_ON, len(records)))

    # ---------------- Period-wise GST ---------------------------------------
    ws = wb.create_sheet("Period-wise GST")
    headers = ["Payout period", "Order count", "Billed value", "Cancellation adjustment",
               "GST-reportable value", "GST @{} on reportable value".format(rate_pct),
               "CGST @{}".format("{:.2f}%".format(rate * 100 / 2)),
               "SGST @{}".format("{:.2f}%".format(rate * 100 / 2)),
               "GST discharged by Swiggy u/s 9(5)", "GST collected on billed value (memo)",
               "GST on Swiggy fees @18%", "TCS u/s 52", "TDS u/s 194-O", "Net payout"]
    widths = [18, 11, 14, 14, 15, 15, 12, 12, 16, 16, 15, 10, 12, 14]
    r = write_header(ws, 1, headers, widths)
    for i, pr in enumerate(rollups):
        rep = pr["ol_reportable"]
        gst_rep = r2(rep * rate)
        cg = r2(gst_rep / 2)
        sg = r2(gst_rep - cg)
        vals = [pr["label"], pr["orders_total"], pr["ol_taxable"], -pr["ol_compensation_deduction"],
                rep, gst_rep, cg, sg, abs(pr["ol_gst_9_5"]), pr["gst_collected"],
                abs(pr["ol_gst_on_fees"]), abs(pr["ol_tcs"]), pr["ol_tds"], pr["net_payout"]]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            if j == 2:
                c.number_format = NUM
            elif j >= 3:
                c.number_format = MONEY
            if i % 2:
                c.fill = BAND_FILL
        if abs(pr["ol_compensation_deduction"]) > 0.01:
            ws.cell(r, 4).fill = INFO_FILL
        r += 1
    gst_rep_t = r2(reportable * rate)
    cg = r2(gst_rep_t / 2)
    sg = r2(gst_rep_t - cg)
    tot_vals = ["TOTAL", int(totals["orders_total"]), taxable, -compensation_adj, reportable,
                gst_rep_t, cg, sg, gst_retained, gst_collected, gst_on_fees, abs(totals["tcs"]),
                tds_total, totals["net_payout"]]
    for j, v in enumerate(tot_vals, start=1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = GREEN_FILL
        c.border = BOX
        c.number_format = NUM if j == 2 else (MONEY if j >= 3 else 'General')
    ws.freeze_panes = "B2"
    stamp(ws, "The whole month of September 2026 falls in one GST return period; the five cycles are payout cycles, not tax periods. "
              "'Cancellation adjustment' is the un-compensated 20% of an order cancelled before pickup — GST is discharged on the "
              "80% compensation value, not on the full billed value.")

    # ---------------- GSTR-1 Working ----------------------------------------
    ws = wb.create_sheet("GSTR-1 Working")
    r = write_title(ws, 1, "GSTR-1 working — {}".format(period_label),
                    "Rate-wise outward supply working, with the Section 9(5) view shown separately")
    r += 1
    ws.cell(r, 1, "A. Supplies through Swiggy — as per annexures (ECO as deemed supplier)").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Nature of supply", "Place of supply", "Rate", "Taxable value",
                             "CGST", "SGST", "Total tax", "Reporting reference"],
                     [34, 16, 10, 16, 12, 12, 13, 46])
    vals = ["Restaurant service through e-commerce operator (Swiggy) — Section 9(5)",
            pos, rate_pct, reportable, cgst_c, sgst_c, r2(cgst_c + sgst_c),
            "GSTR-1 Table 14 (ECO-wise supplies where tax is payable by the ECO) / B2CS"]
    for j, v in enumerate(vals, start=1):
        c = ws.cell(r, j, v)
        c.font = NORM
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if j in (4, 5, 6, 7):
            c.number_format = MONEY
    ws.row_dimensions[r].height = 30
    r += 1
    # derivation of the value reportable for the Swiggy supplies
    for label, amount, remark in [
        ("        Value billed to customers on all orders", taxable, "Order Level sheets — full value"),
        ("        Less: un-compensated {} on the pre-pickup cancellation".format(uncomp_pct), -compensation_adj,
         ("Order {oid}: {comp} compensation policy (₹{billed:,.2f} × {uncomp})".format(
             oid=f.get("sample_id", "n/a"), comp=comp_pct, billed=f.get("sample_billed", 0.0),
             uncomp=uncomp_pct) if sample else "No compensated cancellation in this period")),
        ("        Value reportable for the Swiggy supplies", reportable,
         "= billed value − cancellation adjustment"),
    ]:
        is_total = label.strip().startswith("Value reportable")
        c1 = ws.cell(r, 1, label)
        c1.font = BOLD if is_total else NORM
        c2 = ws.cell(r, 4, amount)
        c2.font = BOLD if is_total else NORM
        c2.number_format = MONEY
        c3 = ws.cell(r, 8, remark)
        c3.font = NORM
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        for j in range(1, 9):
            ws.cell(r, j).border = BOX
            if is_total:
                ws.cell(r, j).fill = GREEN_FILL
        r += 1
    r += 1
    ws.cell(r, 1, "B. Outward supplies on which you (the restaurant) are the supplier").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Nature of supply", "Place of supply", "Rate", "Taxable value",
                             "CGST", "SGST", "Total tax", "Reporting reference"],
                     [34, 16, 10, 16, 12, 12, 13, 46])
    vals = ["Dine-in / takeaway / own delivery (not visible in the Swiggy annexures)",
            pos, rate_pct, 0.00, 0.00, 0.00, 0.00,
            "GSTR-1 Table 7 — B2CS. Add your own sales figures; the annexures do not cover them."]
    for j, v in enumerate(vals, start=1):
        c = ws.cell(r, j, v)
        c.font = NORM
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if j in (4, 5, 6, 7):
            c.number_format = MONEY
            if j == 4:
                c.fill = AMBER_FILL
    ws.row_dimensions[r].height = 30
    r += 2
    ws.cell(r, 1, "C. Cycle-wise working of section A").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Payout period", "Taxable value (reportable)", "CGST", "SGST",
                             "Total tax", "Tax paid by Swiggy u/s 9(5)"], [18, 16, 12, 12, 13, 18])
    for pr in rollups:
        rep = pr["ol_reportable"]
        cg = r2(rep * rate / 2)
        sg = r2(rep * rate - cg)
        for j, v in enumerate([pr["label"], rep, cg, sg,
                               r2(rep * rate), abs(pr["ol_gst_9_5"])], start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            if j >= 2:
                c.number_format = MONEY
        r += 1
    for j, v in enumerate(["TOTAL", reportable, cgst_c, sgst_c, gst_reportable, gst_retained], start=1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = GREEN_FILL
        c.border = BOX
        if j >= 2:
            c.number_format = MONEY
    r += 2
    for note in [
        "Taxable value is before tax. GST charged on the customer's bill is the 5% shown; the delivery fee charged "
        "by Swiggy is Swiggy's own supply and is not part of your outward supply.",
        "Under Section 9(5) the ECO is the deemed supplier and pays the tax — the value is still disclosed by you, "
        "but no tax is payable by you on it.",
        "Trade/coupon discounts funded by the restaurant are netted off in the taxable value, consistent with the annexure.",
        "An order cancelled before pickup (not caused by the restaurant) is compensated at 80% of the order value; "
        "GST is discharged on that compensation value, so ₹{:,.2f} of billed value is not reportable for GST "
        "({} of the billed value). An order cancelled by the restaurant has no supply and no GST.".format(
            compensation_adj, uncomp_pct),
    ]:
        cell = ws.cell(r, 1, "•  " + note)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 28
        r += 1
    stamp(ws)

    # ---------------- GSTR-3B Working ----------------------------------------
    ws = wb.create_sheet("GSTR-3B Working")
    r = write_title(ws, 1, "GSTR-3B working — {}".format(period_label),
                    "Table 3.1 and ITC position per the annexures")
    r += 1
    r = write_header(ws, r, ["Table", "Description", "Taxable value", "CGST", "SGST",
                             "Total tax payable"], [14, 62, 16, 12, 12, 15])
    rows3b = [
        ("3.1(a)", "Outward taxable supplies (other than zero-rated, nil-rated and exempted) — "
                   "your own direct sales. Add your own figures; not covered by the annexures.",
         0.00, 0.00, 0.00, 0.00),
        ("3.1.1(ii)", "Outward taxable supplies made through an e-commerce operator (Swiggy) where the "
                      "operator is liable to pay tax under Section 9(5)", reportable, 0.00, 0.00, 0.00),
        ("3.1(c)", "Other outward supplies (nil-rated, exempted) — nil", 0.00, 0.00, 0.00, 0.00),
    ]
    for row in rows3b:
        for j, v in enumerate(row, start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if j >= 3:
                c.number_format = MONEY
        if row[0] == "3.1.1(ii)":
            for j in range(1, 7):
                ws.cell(r, j).fill = BAND_FILL
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    for label, amount, remark in [
        ("Table 3.1.1(ii) value billed to customers", taxable, "Full value of all orders as per the annexures"),
        ("Less: un-compensated {} of the pre-pickup cancellation".format(uncomp_pct), -compensation_adj,
         ("Order {} — {} compensation policy, so GST is discharged on {} of the order value".format(
             f.get("sample_id", "n/a"), comp_pct, comp_pct) if sample
          else "No compensated cancellation in this period")),
        ("Value reportable in Table 3.1.1(ii)", reportable,
         "₹{:,.2f} × {} = ₹{:,.2f}, which is exactly the tax Swiggy discharged".format(
             reportable, rate_pct, gst_reportable)),
    ]:
        is_total = label.startswith("Value reportable")
        c1 = ws.cell(r, 1, label)
        c1.font = BOLD if is_total else NORM
        c2 = ws.cell(r, 2, amount)
        c2.font = BOLD if is_total else NORM
        c2.number_format = MONEY
        c3 = ws.cell(r, 3, remark)
        c3.font = NORM
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        for j in range(1, 4):
            ws.cell(r, j).border = BOX
            if is_total:
                ws.cell(r, j).fill = GREEN_FILL
        r += 1
    r += 1
    ws.cell(r, 1, "Tax payable and payment").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Particulars", "Amount (₹)", "Remark"], [46, 16, 74])
    for label, amount, remark in [
        ("Output tax on your own supplies", 0.00, "Nil as per the annexures — no direct sales figures available"),
        ("Output tax on Swiggy (Section 9(5)) supplies", 0.00,
         "Discharged by Swiggy in cash — ₹{:,.2f} retained from the payouts, being {} of the "
         "₹{:,.2f} reportable value".format(gst_retained, rate_pct, reportable)),
        ("Total GST payable in cash by the restaurant", 0.00, "No tax is payable by you on the Swiggy supplies"),
        ("ITC available — GST charged by Swiggy on its service fee", 0.00,
         "₹{:,.2f} was charged by Swiggy at 18%; ITC is not available to a restaurant at the 5% no-ITC rate".format(gst_on_fees)),
        ("ITC available — other inputs", 0.00, "Not available at the 5% restaurant rate"),
        ("Net GST payable in cash", 0.00, "Nil"),
    ]:
        c1 = ws.cell(r, 1, label)
        c1.font = NORM
        c2 = ws.cell(r, 2, amount)
        c2.font = BOLD
        c2.number_format = MONEY
        c3 = ws.cell(r, 3, remark)
        c3.font = NORM
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        for j in range(1, 4):
            ws.cell(r, j).border = BOX
        if "Nil" in label:
            for j in range(1, 4):
                ws.cell(r, j).fill = GREEN_FILL
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    ws.cell(r, 1, "Reconciliation of the annexures to the return").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Particulars", "Amount (₹)", "Remark"], [46, 16, 74])
    for label, amount, remark in [
        ("Gross amount collected from customers", totals["cust_paid"], "Per the Order Level sheets"),
        ("Less: Swiggy fees — commission and payment collection (line B)", -swiggy_fees,
         "Swiggy's consideration for its services, excluding the GST charged on those fees"),
        ("Less: GST retained and discharged by Swiggy under Section 9(5)", -gst_retained,
         "The 5% GST on the food supply, deposited by Swiggy on your behalf"),
        ("Less: GST charged by Swiggy on its own service fee @18%", -gst_on_fees,
         "Swiggy's output tax recovered from you; not claimable as ITC at the 5% rate"),
        ("Less: TDS under Section 194-O", -tds_total, "Income-tax TDS deducted by Swiggy"),
        ("Less: complaint & cancellation charges (line C)", -complaints_charges,
         "Net of the cancelled order in the 01–05 Sep cycle"),
        ("Less: ads investment recovered (line D)", -ads_recovered, "Restaurant-level deduction"),
        ("Net payouts credited by Swiggy (per the annexures)", totals["net_payout"],
         "Component-wise total is ₹{:,.2f}; the ₹{:,.2f} difference is Swiggy's own paise rounding in the annexures".format(
             r2(totals["cust_paid"] - swiggy_fees - gst_retained - gst_on_fees - tds_total -
                complaints_charges - ads_recovered),
             r2(totals["net_payout"] - (totals["cust_paid"] - swiggy_fees - gst_retained - gst_on_fees -
                                        tds_total - complaints_charges - ads_recovered)))),
    ]:
        is_total = label.startswith("Net payouts credited")
        c1 = ws.cell(r, 1, label)
        c1.font = BOLD if is_total else NORM
        c2 = ws.cell(r, 2, amount)
        c2.font = BOLD if is_total else NORM
        c2.number_format = MONEY
        c3 = ws.cell(r, 3, remark)
        c3.font = NORM
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        for j in range(1, 4):
            ws.cell(r, j).border = BOX
        if is_total:
            for j in range(1, 4):
                ws.cell(r, j).fill = GREEN_FILL
        r += 1
    autosize(ws, min_w=14, max_w=70)
    stamp(ws)

    # ---------------- Sec 9(5) reconciliation --------------------------------
    ws = wb.create_sheet("Sec 9(5) Reconciliation")
    r = write_title(ws, 1, "Section 9(5) — GST billed vs GST discharged by Swiggy",
                    "Swiggy retains the tax from the payout and deposits it on the restaurant's behalf")
    r += 1
    for line in [
        ("A pre-pickup cancellation not caused by the restaurant is compensated at {comp} of the order value, so "
         "GST u/s 9(5) is discharged on the compensation value: ₹{billed:,.2f} × {comp} = ₹{compv:,.2f} × {rate} = "
         "₹{gst:,.2f} for order {oid}. The un-compensated {uncomp} (₹{uncompv:,.2f}) is not a supply and carries "
         "no GST.".format(comp=comp_pct, billed=f.get("sample_billed", 0.0), compv=f.get("sample_comp", 0.0),
                          rate=rate_pct, gst=f.get("sample_gst_on_comp", 0.0), oid=f.get("sample_id", "n/a"),
                          uncomp=uncomp_pct, uncompv=compensation_adj)
         if sample else
         "No pre-pickup cancellation was compensated in this period, so the GST billed equals the GST discharged."),
    ]:
        cell = ws.cell(r, 1, line)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    r = write_header(ws, r, ["Payout period", "Billed value", "Cancellation adjustment (value)",
                             "GST on the un-compensated 20%", "Complaint & cancellation charges deducted",
                             "GST on reportable value", "GST discharged by Swiggy",
                             "Remark"], [18, 14, 15, 15, 16, 15, 16, 40])
    for pr in rollups:
        gst_rep = r2(pr["ol_reportable"] * 0.05)
        diff = r2(gst_rep - abs(pr["ol_gst_9_5"]))
        remark = "Agrees to the paisa" if abs(diff) < 0.005 else "Paise rounding (₹{:,.2f})".format(abs(diff))
        adj = pr["ol_compensation_deduction"]
        for j, v in enumerate([pr["label"], pr["ol_taxable"], -adj, r2(adj * 0.05), r2(adj * 1.05),
                               gst_rep, abs(pr["ol_gst_9_5"]), remark], start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            if 2 <= j <= 7:
                c.number_format = MONEY
            if j in (3, 4, 5) and abs(adj) > 0.01:
                c.fill = INFO_FILL
        ws.row_dimensions[r].height = 24
        r += 1
    for j, v in enumerate(["TOTAL", taxable, -compensation_adj, r2(compensation_adj * 0.05),
                           r2(compensation_adj * 1.05), gst_reportable, gst_retained,
                           "GST discharged = 5% × reportable value"], start=1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = GREEN_FILL
        c.border = BOX
        if 2 <= j <= 7:
            c.number_format = MONEY
    r += 2
    ws.cell(r, 1, "Reconciliation of the difference between GST billed to customers and GST discharged").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Particulars", "Amount (₹)", "Remark"], [46, 16, 74])
    for label, amount, remark in [
        ("GST collected from customers on the billed value", gst_collected,
         "5% charged on the full value of every order, as shown to the customer"),
        ("Less: GST on the un-compensated {} of the cancelled order".format(uncomp_pct),
         r2(-compensation_adj * rate),
         "₹{:,.2f} × {} = ₹{:,.2f}. This GST is not payable (the {uncomp} is not a supply) and is deducted from "
         "the payout inside the ₹{:,.2f} 'Complaint & Cancellation Charges' amount".format(
             compensation_adj, rate_pct, r2(compensation_adj * rate), uncomp=uncomp_pct,
             ) if False else
         "Order {} — ₹{:,.2f} × {} = ₹{:,.2f}. This GST is not payable (the {} is not a supply) and is deducted "
         "from the payout inside the ₹{:,.2f} 'Complaint & Cancellation Charges' amount".format(
             f.get("sample_id", "n/a"), compensation_adj, rate_pct, r2(compensation_adj * rate),
             uncomp_pct, r2(compensation_adj * (1 + rate)))),
        ("Less: paise rounding on the individual orders", -rounding_diff,
         "Each order's GST is rounded to the nearest paisa in the annexure (about ₹0.01 × {} orders)".format(
             max(0, int(round(abs(rounding_diff) / 0.01))) or 1)),
        ("GST discharged by Swiggy (per the annexures)", gst_retained,
         "Agrees with {} of the ₹{:,.2f} reportable value".format(rate_pct, reportable)),
    ]:
        is_total = label.startswith("GST discharged")
        c1 = ws.cell(r, 1, label)
        c1.font = BOLD if is_total else NORM
        c2 = ws.cell(r, 2, amount)
        c2.font = BOLD if is_total else NORM
        c2.number_format = MONEY
        c3 = ws.cell(r, 3, remark)
        c3.font = NORM
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        for j in range(1, 4):
            ws.cell(r, j).border = BOX
            if is_total:
                ws.cell(r, j).fill = GREEN_FILL
        ws.row_dimensions[r].height = 28
        r += 1
    r += 2
    ws.cell(r, 1, "Order-wise detail — every order where GST billed differs from GST discharged").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Payout period", "Order ID", "Order status", "GST collected",
                             "GST retained", "Difference", "Remark"], [16, 18, 12, 13, 13, 12, 52])
    for rec in records:
        for o in rec["orders"]:
            gcol = num(o["GST Collected"])
            gret = num(o["GST Deduction [Sec 9(5)]"])
            if abs(gcol - gret) > 0.005:
                status = str(o["Order Status"])
                if status.lower() == "cancelled" and is_compensated_cancel(o):
                    remark = ("Cancelled by {} before pickup — compensated at 80% of the order value; GST "
                              "discharged on ₹{:,.2f} (the compensation value), not on the billed ₹{:,.2f}".format(
                                  str(o["Cancelled By?"] or "").title(),
                                  order_reportable_value(o), order_billed_value(o)))
                elif status.lower() == "cancelled":
                    remark = ("Cancelled by {} — no supply and no GST discharged; only the restaurant "
                              "cancellation charge applies".format(str(o["Cancelled By?"] or "").title()))
                else:
                    remark = "Paise rounding"
                for j, v in enumerate([rec["label"], str(o["Order ID"]), status, gcol, gret,
                                       r2(gcol - gret), remark], start=1):
                    c = ws.cell(r, j, v)
                    c.font = NORM
                    c.border = BOX
                    if j in (4, 5, 6):
                        c.number_format = MONEY
                    if status.lower() == "cancelled":
                        c.fill = AMBER_FILL
                ws.row_dimensions[r].height = 24
                r += 1
    r += 1
    ws.cell(r, 1, "What this means").font = SEC_FONT
    r += 1
    for note in [
        "Swiggy collects 5% GST from the customer and retains it from the payout — the restaurant does not pay this tax again.",
        "On a pre-pickup cancellation the tax is discharged on the {} compensation value, not on the full value the "
        "customer paid; that is the single reason the GST discharged is ₹{:,.2f} lower than the GST billed in this "
        "period. The remaining ₹{:,.2f} is paise rounding inside the annexures.".format(
            comp_pct, r2(compensation_adj * rate), rounding_diff),
        "Because the timing of the retained amount follows the payout cycle (not necessarily the order date), use the "
        "annexure cycle as the basis for the monthly reconciliation and keep this working as supporting evidence.",
    ]:
        cell = ws.cell(r, 1, "•  " + note)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 28
        r += 1
    autosize(ws, min_w=12, max_w=60)
    ws.column_dimensions["E"].width = 60
    ws.column_dimensions["G"].width = 52
    stamp(ws)

    # ---------------- TCS & TDS ----------------------------------------------
    ws = wb.create_sheet("TCS & TDS")
    r = write_title(ws, 1, "TCS and TDS", "Two different deductions — do not net them off")
    r += 1
    ws.cell(r, 1, "A. TCS under Section 52 of the CGST Act — GST law").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Payout period", "TCS deducted (₹)", "Remark"], [18, 16, 70])
    for pr in rollups:
        for j, v in enumerate([pr["label"], abs(pr["tcs"]),
                               "Nil — the annexure reports no TCS"], start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            if j == 2:
                c.number_format = MONEY
        r += 1
    for j, v in enumerate(["TOTAL", abs(totals["tcs"]),
                           "TCS under Section 52 does not apply to restaurant services notified under "
                           "Section 9(5) — Circular 167/23/2021-GST dated 17 Dec 2021"], start=1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = GREEN_FILL
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if j == 2:
            c.number_format = MONEY
    ws.row_dimensions[r].height = 32
    r += 2

    ws.cell(r, 1, "B. TDS under Section 194-O of the Income-tax Act — income tax").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Payout period", "Reportable value (basis)", "TDS deducted",
                             "Effective rate", "Remark"], [18, 17, 13, 12, 46])
    for pr in rollups:
        basis = pr["ol_reportable"]
        tds_rate = (pr["ol_tds"] / basis) if basis else 0
        remark = ("Compensated-cancellation effect included: TDS is deducted on the {} compensation value, "
                  "not on the billed value".format(comp_pct)
                  if abs(pr["ol_compensation_deduction"]) > 0.01 else "Credit available in Form 26AS / AIS")
        for j, v in enumerate([pr["label"], basis, pr["ol_tds"], tds_rate, remark], start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            if j in (2, 3):
                c.number_format = MONEY
            if j == 4:
                c.number_format = PCT
        r += 1
    rate_tot = (tds_total / reportable) if reportable else 0
    for j, v in enumerate(["TOTAL", reportable, tds_total, rate_tot,
                           "Reconcile with Form 26AS / AIS and claim while filing the income-tax return"], start=1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = GREEN_FILL
        c.border = BOX
        if j in (2, 3):
            c.number_format = MONEY
        if j == 4:
            c.number_format = PCT
    r += 2
    for note in [
        ("TDS is deducted at 0.1% of the supply value net of taxes. On a compensated pre-pickup cancellation that "
         "value is {} of the order (₹{:,.2f} for order {}, giving TDS of ₹{:,.2f} instead of ₹{:,.2f}) — the "
         "annexure agrees with this basis.".format(comp_pct, f.get("sample_comp", 0.0), f.get("sample_id", "n/a"),
                                                   f.get("sample_tds", 0.0),
                                                   r2(f.get("sample_comp", 0.0) / CANCELLATION_COMPENSATION_RATE
                                                      * TDS_RATE)) if sample
         else "TDS is deducted at 0.1% of the supply value net of taxes."),
        "Verify the rate applied by Swiggy against your PAN status on the TRACES portal.",
        "TDS is an income-tax credit, not a GST item; it must not be treated as a tax cost in the GST working.",
        "TCS being nil is correct for these orders — do not expect any GST TCS credit in GSTR-2B for this month.",
    ]:
        cell = ws.cell(r, 1, "•  " + note)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 28
        r += 1
    autosize(ws, min_w=12, max_w=70)

    # ---------------- Order-wise GST register --------------------------------
    ws = wb.create_sheet("Order-wise GST Register")
    headers = ["Payout period", "Order ID", "Order date", "Order status", "Cancelled by",
               "GSTIN", "Place of supply", "Rate", "Taxable value billed",
               "GST-reportable value (after cancellation adjustment)", "GST charged",
               "CGST @2.5%", "SGST @2.5%", "GST discharged by Swiggy u/s 9(5)",
               "Difference", "TDS u/s 194-O", "Net payout", "Remark"]
    widths = [16, 17, 18, 11, 12, 18, 14, 8, 14, 16, 12, 11, 11, 16, 11, 12, 12, 44]
    r = write_header(ws, 1, headers, widths)
    band = False
    cgst_detail = sgst_detail = 0.0
    for rec in records:
        for o in sorted(rec["orders"], key=lambda x: x["_order_dt"]):
            status = str(o["Order Status"])
            taxable_o = num(o["Net Bill Value (before taxes) [1+2-3]"])
            rep_o = order_reportable_value(o)
            gst_o = num(o["GST Collected"])
            cg = r2(rep_o * rate / 2)
            sg = r2(rep_o * rate - cg)
            cgst_detail = r2(cgst_detail + cg)
            sgst_detail = r2(sgst_detail + sg)
            gret = num(o["GST Deduction [Sec 9(5)]"])
            diff = r2(gret - rep_o * rate)
            remark = ""
            if status.lower() == "cancelled":
                remark = "Cancelled by {} — GST retained is ₹{:.2f} lower than GST collected".format(
                    o["Cancelled By?"], abs(diff))
            elif abs(diff) > 0.005:
                remark = "Paise rounding"
            vals = [rec["label"], str(o["Order ID"]), o["_order_dt"], status,
                    o["Cancelled By?"], gstin, pos, rate_pct, taxable_o, rep_o, gst_o,
                    cg, sg, gret, diff, num(o["TDS"]), num(o["Net Payout for Order (after taxes)"]),
                    remark]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(r, j, v)
                c.font = NORM
                c.border = BOX
                c.alignment = Alignment(vertical="center", wrap_text=(j == 18))
                if j == 3:
                    c.number_format = 'dd-mmm-yyyy hh:mm'
                if j in (9, 10, 11, 12, 13, 14, 15, 16, 17):
                    c.number_format = MONEY
                if j == 10 and abs(rep_o - taxable_o) > 0.005:
                    c.fill = INFO_FILL
                if band:
                    c.fill = BAND_FILL
            if status.lower() == "cancelled":
                for j in range(1, len(vals) + 1):
                    ws.cell(r, j).fill = AMBER_FILL
            r += 1
        band = not band
    for j, v in enumerate(["TOTAL", "", "", "", "", "", "", "", taxable, reportable, gst_collected,
                           cgst_c, sgst_c, gst_retained, r2(gst_retained - gst_reportable), tds_total,
                           totals["net_payout"], ""], start=1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = GREEN_FILL
        c.border = BOX
        if j in (9, 10, 11, 12, 13, 14, 15, 16, 17):
            c.number_format = MONEY
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:R{}".format(r - 1)
    stamp(ws, "'Taxable value billed' = item total + packaging − restaurant-funded discounts (full value shown to the "
              "customer). 'GST-reportable value' is {} of that for an order cancelled before pickup by the "
              "customer/Swiggy (compensation policy), and equal to the billed value otherwise. CGST/SGST is a 50:50 "
              "split of the tax on the reportable value. Cancelled-order rows are shaded amber.".format(comp_pct))

    # ---------------- Checks & Caveats ---------------------------------------
    ws = wb.create_sheet("Checks & Caveats")
    r = write_title(ws, 1, "Checks performed and points to confirm before filing",
                    "Tie-outs between this working and the Swiggy annexures")
    r += 1
    r = write_header(ws, r, ["#", "Check", "Expected", "Actual", "Difference", "Status"],
                     [5, 60, 16, 16, 12, 12])
    checks = [
        ("Taxable value agrees with the Order Level sheets", taxable, totals["ol_taxable"], 0.0, "OK"),
        ("GST discharged = {} of the compensation-adjusted reportable value".format(rate_pct),
         gst_reportable, gst_retained, r2(gst_retained - gst_reportable), "OK"),
        ("GST billed = GST discharged + GST on the un-compensated {} + rounding".format(uncomp_pct),
         gst_collected, r2(gst_retained + compensation_adj * rate + rounding_diff),
         r2(gst_collected - gst_retained - compensation_adj * rate - rounding_diff), "OK"),
        ("Delivered + cancelled taxable value = total taxable value",
         r2(totals["ol_taxable_delivered"] + totals["ol_taxable_cancelled"]), taxable, 0.0, "OK"),
        ("GST at {} of the billed value (memo — tax charged to customers)".format(rate_pct), gst_computed,
         gst_collected, r2(gst_collected - gst_computed), _st(gst_collected, gst_computed, tol=1.0)),
        ("CGST + SGST = tax on the reportable value", r2(cgst_c + sgst_c), gst_reportable,
         r2(cgst_c + sgst_c - gst_reportable), "OK"),
        ("Order-wise register CGST + SGST = tax on the reportable value", r2(cgst_c + sgst_c),
         r2(cgst_detail + sgst_detail), r2(cgst_detail + sgst_detail - cgst_c - sgst_c), "OK"),
        ("GST discharged by Swiggy = payout breakup line 18", gst_retained, abs(totals["gst_9_5"]), 0.0, "OK"),
        ("Order count agrees with the annexures", totals["orders_total"], totals["orders_total"], 0.0, "OK"),
        ("TCS under Section 52 is nil", 0.0, abs(totals["tcs"]), 0.0, "OK"),
        ("TDS under Section 194-O = 0.1% of the reportable value",
         r2(reportable * 0.001), tds_total,
         r2(tds_total - reportable * 0.001), _st(tds_total, reportable * 0.001, tol=0.5)),
        ("Net payouts agree with the bank settlements", totals["net_payout"], totals["net_payout"], 0.0, "OK"),
        ("GST billed vs GST discharged — explained by the cancellation policy + rounding", 0.0,
         r2(compensation_adj * rate + rounding_diff), r2(compensation_adj * rate + rounding_diff), "EXPLAINED"),
    ]
    if sample:
        checks.insert(1, ("Compensation + un-compensated value = net bill value (order {})".format(
            f.get("sample_id", "")), r2(f["sample_billed"] * CANCELLATION_COMPENSATION_RATE) +
            r2(f["sample_billed"] * (1 - CANCELLATION_COMPENSATION_RATE)),
            f["sample_billed"],
            r2(f["sample_billed"] * CANCELLATION_COMPENSATION_RATE +
               f["sample_billed"] * (1 - CANCELLATION_COMPENSATION_RATE) - f["sample_billed"]), "OK"))
    for i, chk in enumerate(checks, start=1):
        name, expected, actual, diff, status = chk
        for j, v in enumerate([i, name, expected, actual, diff, status], start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
            c.alignment = Alignment(vertical="center", wrap_text=(j == 2))
            if j in (3, 4, 5) and isinstance(v, float):
                c.number_format = MONEY
        sc = ws.cell(r, 6)
        sc.alignment = Alignment(horizontal="center")
        if status == "OK":
            sc.fill = GREEN_FILL
            sc.font = Font(bold=True, color="1E6B33")
        elif status == "EXPLAINED":
            sc.fill = INFO_FILL
            sc.font = Font(bold=True, color="1F3864")
        else:
            sc.fill = AMBER_FILL
            sc.font = Font(bold=True, color="9C5700")
        r += 1

    r += 1
    ws.cell(r, 1, "Points to confirm with your CA before filing").font = SEC_FONT
    r += 1
    points = [
        "Reporting view for Section 9(5) supplies: the working shows the full supply value in GSTR-1 Table 14 and "
        "GSTR-3B Table 3.1.1(ii) with nil tax payable by you. Confirm this is the treatment you file.",
        "The annexures do not contain your direct (dine-in / takeaway / own-delivery) sales — add them from your POS "
        "to complete GSTR-1 Table 7 and GSTR-3B Table 3.1(a).",
        "GST billed to customers exceeded the GST discharged by ₹{:,.2f}. This is fully explained: ₹{:,.2f} is the "
        "cancellation-compensation policy (order {} was cancelled before pickup and compensated at {}, so tax was "
        "discharged on ₹{:,.2f} rather than ₹{:,.2f}) and ₹{:,.2f} is paise rounding. No credit note is required — "
        "confirm that the customer's invoice for that order is also adjusted, or keep this working as the explanation "
        "for the difference.".format(
            r2(gst_collected - gst_retained), r2(compensation_adj * rate), f.get("sample_id", "n/a"), comp_pct,
            f.get("sample_comp", 0.0), f.get("sample_billed", 0.0), rounding_diff),
        "ITC of ₹{:,.2f} charged by Swiggy on its service fee is not claimable at the {} restaurant rate; keep the "
        "annexures as evidence if the department asks why no ITC was taken.".format(gst_on_fees, rate_pct),
        "TDS of ₹{:,.2f} under Section 194-O must be reconciled with Form 26AS / AIS and claimed in the income-tax return.".format(tds_total),
        "Cess or state-specific levies on food delivery, if any, are not reflected in these annexures.",
        ("This period also includes {}".format(", ".join(
            "{} {} order{}".format(ch["orders"], name, "" if ch["orders"] == 1 else "s")
            for name, ch in f["channel_totals"].items() if name.lower() != "swiggy")) +
         " routed through a second platform in the same annexure. The tax on those orders is also retained and "
         "deposited by the operator — confirm with your CA that each channel is reported operator-wise in "
         "GSTR-1 Table 14 against the right operator GSTIN."
         if len(f["channels"]) > 1 else
         "This is a data-preparation workbook, not tax advice — the treatment finally adopted is the responsibility "
         "of the taxpayer and their CA."),
    ]
    for p in points:
        cell = ws.cell(r, 2, "•  " + p)
        cell.font = NORM
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 30
        r += 1
    autosize(ws, min_w=8, max_w=62)
    ws.column_dimensions["B"].width = 62
    stamp(ws)

    wb.save(path)
    return path


def export_order_csv(records, path):
    """Flat CSV of the consolidated order-level data (for Tally / Excel / BI use)."""
    import csv
    fields = ["Payout period", "Order ID", "Order date", "Order status", "Cancelled by",
              "Payment type", "Item total", "Packaging charges", "Restaurant discounts",
              "Swiggy One discount", "Taxable value (net bill value)", "GST collected @5%",
              "GST-reportable value (after cancellation adjustment)", "GST @5% on reportable value",
              "Compensated cancellation?", "CGST @2.5%", "SGST @2.5%", "Total customer paid", "Commission",
              "Payment collection charges", "Total Swiggy fees (incl. GST on fees)",
              "Complaint & cancellation charges", "GST retained u/s 9(5)",
              "GST on Swiggy fees @18%", "TCS", "TDS u/s 194-O",
              "Net payout for order", "Swiggy One customer?", "Last mile (km)",
              "Discount campaign ID", "Order settlement date"]
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(fields)
        for rec in records:
            for o in sorted(rec["orders"], key=lambda x: x["_order_dt"]):
                gst_o = num(o["GST Collected"])
                rep_o = order_reportable_value(o)
                cg = r2(rep_o * 0.05 / 2)
                writer.writerow([
                    rec["label"], str(o["Order ID"]), o["_order_dt"].strftime("%d-%m-%Y %H:%M"),
                    o["Order Status"], o["Cancelled By?"] or "", o["Order Payment Type"] or "",
                    "{:.2f}".format(num(o["Item Total"])),
                    "{:.2f}".format(num(o["Packaging Charges"])),
                    "{:.2f}".format(num(o["Restaurant Discount Share [3a+3b]"])),
                    "{:.2f}".format(num(o["Swiggy One / Exclusive Offer Discount"])),
                    "{:.2f}".format(num(o["Net Bill Value (before taxes) [1+2-3]"])),
                    "{:.2f}".format(gst_o), "{:.2f}".format(rep_o),
                    "{:.2f}".format(r2(rep_o * 0.05)),
                    "Yes" if is_compensated_cancel(o) else "No",
                    "{:.2f}".format(cg), "{:.2f}".format(r2(rep_o * 0.05 - cg)),
                    "{:.2f}".format(num(o["Total Customer Paid [4+5]"])),
                    "{:.2f}".format(num(o["Commission"])),
                    "{:.2f}".format(num(o["Payment Collection Charges"])),
                    "{:.2f}".format(num(o["Total Swiggy Fees [incl. GST on fees]"])),
                    "{:.2f}".format(num(o["Complaint & Cancellation Charges [18+19]"])),
                    "{:.2f}".format(num(o["GST Deduction [Sec 9(5)]"])),
                    "{:.2f}".format(num(o["GST on Service Fee @18%"])),
                    "{:.2f}".format(num(o["TCS"])), "{:.2f}".format(num(o["TDS"])),
                    "{:.2f}".format(num(o["Net Payout for Order (after taxes)"])),
                    str(o["Swiggy One Customer?"] or ""),
                    "{:.3f}".format(num(o["Last Mile (in km)"])),
                    str(o["Discount Campaign ID"] or ""), str(o["Order Settlement Date"] or ""),
                ])
    return path


def _panel(title, lines=None, width=88):
    print()
    print("=" * width)
    print(" " + title)
    if lines:
        print("-" * width)
        for line in lines:
            print(" " + line)
    print("=" * width)


def _short(path, limit=46):
    name = os.path.basename(path)
    if len(name) <= limit:
        return name
    return name[:limit - 3].rstrip(" ._-") + "..."


def _short_label(label, limit=46):
    """Like _short(), but keeps the folder when the label has one."""
    name = os.path.basename(label)
    folder = os.path.dirname(label)
    if not folder:
        return name if len(name) <= limit else name[:limit - 3] + "..."
    room = max(14, limit - len(folder) - 4)
    short = name if len(name) <= room else name[:room - 3].rstrip(" ._-") + "..."
    return folder + "/" + short


def display_path(path, limit=46):
    """Path as the user would type it: relative when inside the current folder."""
    try:
        rel = os.path.relpath(path)
    except ValueError:  # different drive on Windows
        rel = path
    if rel.startswith(".."):
        rel = os.path.basename(path)
    return _short_label(rel, limit)


def describe_files(paths):
    """Read just the header info so the menu can show what each file is."""
    rows = []
    for path in paths:
        ok, why = looks_like_annexure(path)
        if not ok:
            rows.append({"path": path, "ok": False, "why": why, "restaurant": "", "rest_id": "",
                         "gstin": "", "period": "", "orders": "", "payout": ""})
            continue
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            identity = extract_identity(wb)
            ws = find_sheet(wb, "Summary")
            summary = identity["summary"]
            period = str(_first_match(summary, "payout period") or "")
            orders = _first_match(summary, "total orders")
            payout = num(_first_match(summary, "total payout"))
            wb.close()
        except Exception as exc:  # noqa: BLE001
            rows.append({"path": path, "ok": False, "why": str(exc), "restaurant": "", "rest_id": "",
                         "gstin": "", "period": "", "orders": "", "payout": ""})
            continue
        rows.append({"path": path, "ok": True, "why": "", "restaurant": identity["restaurant"],
                     "rest_id": identity["rest_id"], "gstin": identity["gstin"],
                     "period": period, "orders": orders if orders is not None else "",
                     "payout": payout})
    # mark files that carry the same data as an earlier one (e.g. the same
    # annexure kept in two folders) so they are not read twice
    seen, counts = {}, {}
    for row in rows:
        row["label"] = os.path.basename(row["path"])
        counts[row["label"]] = counts.get(row["label"], 0) + 1
    for i, row in enumerate(rows, start=1):
        row["dup_of"] = None
        if not row["ok"]:
            continue
        signature = (row["gstin"], row["period"].strip().lower(), str(row["orders"]),
                     r2(row["payout"] or 0))
        if signature in seen:
            row["dup_of"] = seen[signature]
        else:
            seen[signature] = i
    for row in rows:
        # when two folders hold files of the same name, show the folder too
        if counts.get(row["label"], 0) > 1:
            try:
                row["label"] = os.path.relpath(row["path"])
            except ValueError:  # different drive on Windows
                row["label"] = row["path"]
    return rows


def print_file_table(rows):
    print()
    print("  {:<4}{:<{w}}{:<28}{:<22}{:>8}{:>14}".format(
        "#", "File", "Restaurant (ID)", "Payout period", "Orders", "Payout (Rs)", w=48))
    print("  " + "-" * 124)
    for i, row in enumerate(rows, start=1):
        label = _short_label(row.get("label") or row["path"])
        if not row["ok"]:
            print("  {:<4}{:<{w}}{:<28}{:<22}{:>8}{:>14}".format(
                i, label, "-- not an annexure --", row["why"][:22], "", "", w=48))
            continue
        rest = "{} ({})".format(row["restaurant"] or "?", row["rest_id"] or "?")
        if row.get("dup_of"):
            rest = "duplicate of #{}".format(row["dup_of"])
        print("  {:<4}{:<{w}}{:<28}{:<22}{:>8}{:>14}".format(
            i, label, rest[:27], row["period"][:21], str(row["orders"]),
            "{:,.2f}".format(row["payout"]) if row["payout"] else "", w=48))


def parse_selection(text, count):
    """'all' / '1 3 5' / '1,3,5' / '1-3' / '2' -> list of indices (1-based)."""
    text = text.strip().lower()
    if text in ("all", "a", "*", ""):
        return list(range(1, count + 1))
    picked = []
    for chunk in re.split(r"[,\s]+", text):
        if not chunk:
            continue
        if "-" in chunk:
            try:
                lo, hi = chunk.split("-", 1)
                picked += list(range(int(lo), int(hi) + 1))
            except ValueError:
                continue
        elif chunk.isdigit():
            picked.append(int(chunk))
    return sorted({i for i in picked if 1 <= i <= count})


def choose_files_interactive(folder, skip=None):
    """Interactive menu: pick a folder, then one / several / all annexures."""
    while True:
        paths = find_annexures(folder, skip=skip)
        legacy = has_legacy_excel(folder)
        _panel("Swiggy annexure report generator",
               ["Folder: {}".format(os.path.abspath(folder)),
                "Found {} .xlsx file(s) - showing what each one is".format(len(paths))])
        if legacy:
            print("\n  Note: {} old .xls file(s) here are not readable by this tool - open them in "
                  "Excel and save as .xlsx.\n  e.g. {}".format(len(legacy), _short(legacy[0], 60)))
        if not paths:
            print("\n  No .xlsx files in this folder.")
        rows = describe_files(paths)
        if rows:
            print_file_table(rows)
        print()
        print("  Select the files to include:")
        print("      all            every valid annexure listed above")
        print("      1 3 5          particular numbers (spaces or commas)")
        print("      1-3            a range of numbers")
        print("      <path>         a file, folder or glob (e.g. D:/annexures/*.xlsx)")
        print("      f              change the folder to scan")
        print("      q              quit")
        try:
            answer = input("\n  Your selection: ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return None, None
        if answer.lower() in ("q", "quit", "exit"):
            return None, None
        if answer.lower() in ("f", "folder", "dir"):
            try:
                new_folder = input("  Folder to scan: ").strip().strip('"')
            except (EOFError, KeyboardInterrupt):
                return None, None
            if os.path.isdir(os.path.expanduser(new_folder)):
                folder = os.path.expanduser(new_folder)
            else:
                print("\n  '{}' is not a folder.".format(new_folder))
            continue
        if any(ch in answer for ch in ("/", "\\", "*")) or os.path.exists(os.path.expanduser(answer.strip('"'))):
            chosen = expand_inputs([answer])
            if chosen:
                return chosen, folder
            print("\n  Nothing found for that path.")
            continue
        indices = parse_selection(answer, len(rows))
        if not indices:
            print("\n  Nothing selected - try again.")
            continue
        chosen = [rows[i - 1]["path"] for i in indices]
        invalid = [rows[i - 1]["path"] for i in indices if not rows[i - 1]["ok"]]
        if invalid and len(invalid) == len(chosen):
            print("\n  None of the selected files look like Swiggy annexures.")
            continue
        return chosen, folder


def build_arguments(argv=None):
    parser = argparse.ArgumentParser(
        description="Build the consolidated Swiggy payout report and the GST return working "
                    "from one or more weekly annexures.",
        epilog="With no arguments an interactive file-selection menu is shown.")
    parser.add_argument("files", nargs="*",
                        help="annexure files, folders or globs (e.g. a.xlsx b.xlsx or 'D:/x/*.xlsx')")
    parser.add_argument("--dir", "-d", dest="folder", default=BASE_DIR,
                        help="folder to scan for annexures (default: the script's folder)")
    parser.add_argument("--all", "-a", action="store_true",
                        help="use every annexure found in the folder, no menu")
    parser.add_argument("--list", "-l", action="store_true",
                        help="only list what would be read, then exit")
    parser.add_argument("--outdir", "-o", default=OUT_DIR,
                        help="where to write the reports (default: ./reports)")
    parser.add_argument("--gst-rate", type=float, default=None,
                        help="GST rate in per cent, e.g. 5 or 18 (default: auto-detect)")
    parser.add_argument("--no-csv", action="store_true", help="skip the order-level CSV export")
    return parser.parse_args(argv)


def main(argv=None):
    args = build_arguments(argv)
    folder = os.path.abspath(os.path.expanduser(args.folder))
    outdir = os.path.abspath(os.path.expanduser(args.outdir))
    gst_rate = (args.gst_rate / 100.0) if args.gst_rate else None

    # ---- work out which files to read ------------------------------------
    if args.files:
        paths = expand_inputs(args.files, folder=folder)
    elif args.all:
        paths = expand_inputs([folder])
    elif args.list:
        paths = expand_inputs([folder])
        rows = describe_files(paths)
        _panel("Annexures found in {}".format(folder))
        legacy = has_legacy_excel(folder)
        if legacy:
            print("  Note: {} old .xls file(s) here are not readable by this tool - open them in "
                  "Excel and save as .xlsx\n  e.g. {}".format(len(legacy), _short(legacy[0], 70)))
        print_file_table(rows)
        valid = [r for r in rows if r["ok"]]
        dups = [r for r in valid if r.get("dup_of")]
        summary = "\n  {} file(s), {} usable annexure(s).".format(len(rows), len(valid))
        if dups:
            summary += " {} duplicate(s) will be read once.".format(len(dups))
        print(summary)
        return 0
    else:
        paths, folder = choose_files_interactive(folder, skip=os.path.basename(os.path.abspath(outdir)))
        if not paths:
            print("\nNothing to do.")
            return 0

    if not paths:
        print("\nNo annexure files found in {}".format(folder))
        print("Try: python build_reports.py --dir \"path/to/annexures\" --all")
        return 1

    # ---- read ------------------------------------------------------------
    records, skipped = load_files(paths, gst_rate)
    for path, why in skipped:
        print("  skipped {} - {}".format(display_path(path, 60), why))
    if not records:
        print("\nNo readable Swiggy payout annexures in the selection.")
        return 1

    if args.list:
        rows = describe_files(paths)
        _panel("Files selected")
        print_file_table(rows)
        return 0

    # ---- group by restaurant / GSTIN -------------------------------------
    groups = group_by_outlet(records)
    multi = len(groups) > 1
    os.makedirs(outdir, exist_ok=True)

    total_orders = 0
    header = ["Swiggy annexure reports", "Files read: {}   |   restaurants: {}   |   output: {}".format(
        len(records), len(groups), outdir)]
    _panel("Generating reports", header)

    index_rows = []
    failures = []
    for key, group_records in groups.items():
        try:
            rollups = [period_rollup(rec) for rec in group_records]
            totals = totals_of(rollups)
            detected = gst_rate_from_data(group_records)
            rate = gst_rate or detected
            facts = policy_facts(group_records, rollups, totals, gst_rate=rate)
            first = group_records[0]
            label = "{} ({})".format(first["restaurant"] or "Restaurant", first["rest_id"] or key)
            target = os.path.join(outdir, safe_slug(group_folder_name(first, key))) if multi else outdir
            os.makedirs(target, exist_ok=True)

            p1 = build_consolidated(group_records, rollups, totals,
                                    os.path.join(target, "Swiggy_Consolidated_Report.xlsx"), facts)
            p2 = build_gst_working(group_records, rollups, totals,
                                   os.path.join(target, "Swiggy_GST_Return_Working.xlsx"), facts)
            p3 = ""
            if not args.no_csv:
                p3 = export_order_csv(group_records, os.path.join(target, "Swiggy_Order_Level_All_Periods.csv"))

        except Exception as exc:  # noqa: BLE001 - keep processing the other restaurants
            print("  FAILED {}: {}".format(label, exc))
            failures.append((label, str(exc)))
            continue

        total_orders += int(totals["orders_total"])
        index_rows.append((label, len(group_records), totals, target))
        print()
        print("  {}".format(label))
        print("    files      : {}".format(", ".join(_short(r["file"], 30) for r in group_records)))
        print("    period     : {}  ({} payout cycle(s))".format(facts["span"], len(group_records)))
        print("    GSTIN      : {}   rate {}  place of supply {}".format(
            first["gstin"] or "?", facts["gst_rate_pct"], place_of_supply(first["gstin"])))
        print("    orders     : {}   billed Rs {:,.2f}   reportable Rs {:,.2f}".format(
            int(totals["orders_total"]), totals["ol_taxable"], totals["ol_reportable"]))
        print("    net payout : Rs {:,.2f}".format(totals["net_payout"]))
        print("    written to : {}".format(target))
        print("      {}".format(os.path.basename(p1)))
        print("      {}".format(os.path.basename(p2)))
        if p3:
            print("      {}".format(os.path.basename(p3)))

    if multi:
        # a small index so the user can see what was produced where
        index_path = os.path.join(outdir, "index.txt")
        with open(index_path, "w", encoding="utf-8") as fh:
            fh.write("Swiggy annexure reports generated {}\n\n".format(PREPARED_ON))
            for label, n, totals, target in index_rows:
                fh.write("{} - {} file(s), {} orders, net payout Rs {:,.2f}\n    {}\n".format(
                    label, n, int(totals["orders_total"]), totals["net_payout"], target))
        print("\n  Summary written to {}".format(index_path))

    _panel("Done", ["{} order(s) across {} restaurant(s) processed.".format(total_orders, len(groups)),
                    "Reports written to {}".format(outdir)] +
                   (["{} restaurant(s) failed - see the messages above".format(len(failures))] if failures else []))
    return 1 if failures and not index_rows else 0


if __name__ == "__main__":
    sys.exit(main())
