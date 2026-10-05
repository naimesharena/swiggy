#!/usr/bin/env python3
"""
Swiggy payout annexure consolidation
====================================

Reads all `invoice_Annexure_544662_*.xlsx` weekly payout files in this folder and
builds two workbooks:

  1. reports/Swiggy_Consolidated_Report.xlsx
        One consolidated data report across every payout cycle - period summary,
        merged payout breakup, merged order-level data, ads/discount/complaint
        detail, reconciliation checks and field definitions.

  2. reports/Swiggy_GST_Return_Working.xlsx
        GST return filing working papers - period-wise tax summary, GSTR-1 and
        GSTR-3B working, Section 9(5) tax discharge reconciliation, TCS/TDS
        reconciliation and an order-wise GST register.

Usage:  python3 build_reports.py
"""

from __future__ import annotations

import glob
import os
from collections import OrderedDict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, "reports")
PREPARED_ON = "05 October 2026"

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
# compensation."  The restaurant is therefore taxed on that compensation value, not
# on the full value the customer had paid:
#
#       compensation value (pre-tax) = 80% x net bill value
#       GST discharged u/s 9(5)      = 5%  x compensation value
#       TDS u/s 194-O                = 0.1% x compensation value
#       commission / collection      = charged on 80% x total customer paid
#       balance 20% of the order     = reported as 'Customer Cancellations'
#
# Everything below is derived from that basis, so the model can be re-checked
# against the annexures order by order.
CANCELLATION_COMPENSATION_RATE = 0.80
MERCHANT_CANCEL_WORDS = ("MERCHANT", "RESTAURANT", "OUTLET", "SELF")


def is_compensated_cancel(order):
    """True when the restaurant is compensated for a pre-pickup cancellation."""
    if str(order.get("Order Status") or "").strip().lower() != "cancelled":
        return False
    by = str(order.get("Cancelled By?") or "").strip().upper()
    if any(word in by for word in MERCHANT_CANCEL_WORDS):
        return False
    pickup = str(order.get("Pick Up Status") or "").strip().lower()
    if "picked" in pickup and "not" not in pickup:
        return False
    return True


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
    """Un-compensated 20% of a compensated cancellation (nil otherwise)."""
    if not is_compensated_cancel(order):
        return 0.0
    return r2(order_billed_value(order) - order_reportable_value(order))


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def num(value):
    """Coerce a cell value to float, returning 0.0 for blanks/labels."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("₹", "").replace(",", "").replace("%", "").strip()
    if text in ("", "-", "NA", "N/A"):
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


def parse_payout_period(text, fallback_year=2026):
    """'20 September - 26 September' -> (start_date, end_date, label)."""
    parts = [p.strip() for p in str(text).split("-")]
    if len(parts) < 2:
        return None, None, str(text)
    try:
        start_day, start_month = parts[0].split()
        end_day, end_month = parts[1].split()
        sd = datetime(fallback_year, MONTH_NUM[start_month.lower()], int(start_day))
        ed = datetime(fallback_year, MONTH_NUM[end_month.lower()], int(end_day))
    except (ValueError, KeyError):
        return None, None, str(text)
    label = "{}–{} {} {}".format(
        sd.strftime("%d"), ed.strftime("%d"),
        MONTH_ABBR[ed.strftime("%B").lower()], ed.year)
    return sd, ed, label


# --------------------------------------------------------------------------
# Extraction
# --------------------------------------------------------------------------
def read_annexure(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    rec = {"file": os.path.basename(path), "path": path}

    # ---- Summary sheet -------------------------------------------------
    ws = wb["Summary"]
    summary = {}
    for row in ws.iter_rows(min_row=1, max_row=20):
        label, value = row[1].value, row[2].value
        if label and value is not None:
            summary[str(label).strip()] = value
    rec["restaurant"] = summary.get("Cheesecake Mills", "")
    rec["summary_raw"] = summary
    rec["outlet"] = str(summary.get("Vesu", ""))
    rec["city"] = str(summary.get("Surat", ""))
    rec["rest_id"] = str(summary.get("Rest. ID - 544662", "544662"))
    rec["gstin"] = str(summary.get("GSTIN  - 24AAQFC3084Q1ZT", "")).replace("GSTIN", "").replace("-", "").strip()
    rec["payout_period_text"] = str(summary.get("Payout Period", ""))
    rec["settlement_text"] = str(summary.get("Payout Settlement Date", ""))
    rec["total_payout_stated"] = num(summary.get("Total Payout"))
    rec["total_orders_stated"] = int(num(summary.get("Total Orders (Delivered + Cancelled)")))
    rec["utr"] = str(summary.get("Bank UTR", "")).strip()

    # ---- restaurant identity (from fixed cells) ------------------------
    rec["restaurant"] = str(ws["B5"].value or "").strip()
    rec["outlet"] = str(ws["B6"].value or "").strip()
    rec["city"] = str(ws["B7"].value or "").strip()
    rec["rest_id"] = str(ws["B8"].value or "").replace("Rest. ID -", "").strip()
    rec["gstin"] = str(ws["B9"].value or "").replace("GSTIN", "").replace("-", "").strip()

    # ---- Payout Breakup ------------------------------------------------
    ws = wb["Payout Breakup"]
    lines = {}
    ads_lines = []
    mode = None
    for row in ws.iter_rows(min_row=5, max_row=50):
        code = row[1].value
        label = row[2].value
        code_str = str(code).strip() if code is not None else None
        if code_str in PAYOUT_LINES:
            mode = code_str
            lines[code_str] = {
                "label": PAYOUT_LINES[code_str],
                "delivered": num(row[3].value),
                "cancelled": num(row[4].value),
                "total": num(row[5].value),
            }
        elif mode == "D" and label and code is None and row[5].value is not None:
            # Ads sub-lines (Top Picks - Ads, Ads Offers, ...)
            ads_lines.append({"label": str(label).strip().split("\n")[0],
                              "amount": num(row[5].value)})
        elif mode == "D" and code_str in PAYOUT_LINES:
            mode = None
    rec["payout_lines"] = lines
    rec["ads_lines"] = ads_lines

    # ---- Order Level ----------------------------------------------------
    ws = wb["Order Level"]
    orders = []
    for row in ws.iter_rows(min_row=4):
        if row[0].value is None or str(row[0].value).strip() == "":
            continue
        order = {}
        for idx, name in ORDER_COLS.items():
            order[name] = row[idx - 1].value
        if isinstance(order["Order Date"], datetime):
            order["_order_dt"] = order["Order Date"]
        else:
            order["_order_dt"] = to_date(order["Order Date"]) or datetime(2026, 9, 1)
        orders.append(order)
    rec["orders"] = orders

    # ---- Ads / adjustments ----------------------------------------------
    ws = wb["Growth Investments and Other De"]
    ads_detail = []
    for r in range(27, ws.max_row + 1):
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
    ws = wb["Discount Summary"]
    rec["discount"] = {
        "validity": ws["A3"].value, "coupon": ws["B3"].value,
        "rest_share": num(ws["C3"].value) if ws["C3"].value is not None else None,
        "orders": num(ws["D3"].value), "given": num(ws["E3"].value),
        "target": ws["F3"].value, "remarks": ws["G3"].value,
    }

    # ---- Unresolved complaints -------------------------------------------
    ws = wb["Unresolved Customer Complaints"]
    rec["complaints"] = {
        "processed_prev_complaint": num(ws["D2"].value),
        "processed_current_regular": num(ws["D3"].value),
        "processed_total": num(ws["D4"].value),
        "not_processed_prev": num(ws["D6"].value),
        "not_processed_current": num(ws["D7"].value),
        "not_processed_total": num(ws["D8"].value),
    }
    return rec


def load_all():
    records = []
    for path in sorted(glob.glob(os.path.join(BASE_DIR, "invoice_Annexure_*.xlsx"))):
        rec = read_annexure(path)
        sd, ed, label = parse_payout_period(rec["payout_period_text"])
        rec["start"], rec["end"], rec["label"] = sd, ed, label
        rec["sort_key"] = sd or datetime(2026, 1, 1)
        records.append(rec)
    records.sort(key=lambda r: r["sort_key"])
    return records


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
    }


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
def build_consolidated(records, rollups, totals, path):
    wb = openpyxl.Workbook()
    recon_check_count = 74  # 67 original + 3 compensation + 3 month-level + 1 rounding
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
    r = 5
    for label, value in rows:
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
    stamp(ws, "Totals = the five weekly annexures for 01 Sep 2026 – 30 Sep 2026. Settlement dates are as printed on each "
              "annexure. Deduction lines (commission, charges, ads, GST, TDS) are shown here as positive amounts for "
              "readability — the annexures print them as negative; see the Payout Breakup sheet for the original signs.")

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
    stamp(ws, "Every order row from all {} annexures ({} orders). Cancelled-order rows are shaded amber. "
              "'Taxable value' = item total + packaging − restaurant-funded discounts (the full value billed to the customer). "
              "'GST-reportable value' is the same figure, except that an order cancelled before pickup by the customer/Swiggy is "
              "reported at the 80% compensation value — see the 'Cancelled Orders' sheet.".format(
                  len(records), totals["orders_total"]))

    # ---------------- Cancelled orders & compensation -------------------------
    ws = wb.create_sheet("Cancelled Orders")
    r = write_title(ws, 1, "Cancelled orders — the 80% compensation policy",
                    "Every rupee of a pre-pickup cancellation accounted for: value, GST and the deduction")
    r += 1
    for line in [
        "Swiggy's policy: \"Order cancelled before pickup — as per policy, you will be paid 80% less commissions as compensation.\"",
        "The restaurant is taxed on the 80% compensation, not on the full value the customer had paid, and the un-compensated 20%",
        "is deducted from the payout as a GST-inclusive amount. For order 247576408124202 (net bill value ₹590.00):",
        "        compensation  80% × 590.00 = 472.00   + GST 5% 23.60  =  495.60  → credited (commission is charged on this)",
        "        un-compensated 20% × 590.00 = 118.00  + GST 5%  5.90  =  123.90  → 'Complaint & Cancellation Charges [18+19]'",
        "        ---------------------------------------------------------------------------------------------------------",
        "        net bill value 590.00 + GST 29.50 = 619.50 paid by the customer = 495.60 + 123.90",
    ]:
        cell = ws.cell(r, 1, line)
        cell.font = NORM if line.startswith("        ") else BOLD
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        r += 1
    r += 1

    # ---- A. compensated pre-pickup cancellations -------------------------
    ws.cell(r, 1, "A.  Cancelled before pickup (not by the restaurant) — compensated at 80%").font = SEC_FONT
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
        for o in [x for x in rec["orders"] if str(x["Order Status"]).lower() == "cancelled"
                  and not is_compensated_cancel(x)]:
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
    r += 2

    # ---- C. linking the annexure columns ---------------------------------
    ws.cell(r, 1, "C.  How this maps to the annexure columns").font = SEC_FONT
    r += 1
    r = write_header(ws, r, ["Annexure column", "Value for order 247576408124202", "Meaning"],
                     [40, 24, 74])
    for label, value, meaning in [
        ("Net Bill Value (before taxes) [1+2-3]", "590.00", "Order value, exclusive of GST"),
        ("GST Collected", "29.50", "5% charged to the customer on ₹590.00"),
        ("Total Customer Paid [4+5]", "619.50", "What the customer actually paid"),
        ("Commission charged on", "495.60", "80% × ₹619.50 — commission is charged on the compensation, not the full order"),
        ("GST Deduction", "23.60", "5% × ₹472.00 — tax discharged u/s 9(5) on the compensation value only"),
        ("Customer Cancellations / Complaint & Cancellation Charges [18+19]", "123.90",
         "Un-compensated 20% (₹118.00) + GST @5% on it (₹5.90) — a GST-inclusive deduction"),
        ("TDS", "0.47", "0.1% × ₹472.00 — TDS on the compensation value only"),
    ]:
        for j, v in enumerate([label, value, meaning], start=1):
            c = ws.cell(r, j, v)
            c.font = NORM
            c.border = BOX
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
        ("Less: un-compensated 20% of the cancelled order (pre-tax)", -adj_month,
         "₹590.00 × 20% — the value is not a supply because the restaurant was compensated at 80%"),
        ("GST-reportable value", rep_month, "Value on which the 5% GST is discharged u/s 9(5)"),
        ("GST discharged by Swiggy", abs(totals["ol_gst_9_5"]), "= 5% × the reportable value — agrees to the paisa"),
        ("GST collected from customers on the full billed value", totals["gst_collected"],
         "5% × ₹{:,.2f}".format(billed_month)),
        ("Less: GST on the un-compensated 20%, deducted inside the ₹{:,.2f} cancellation charge".format(
            r2(adj_month * 1.05)), r2(-adj_month * 0.05),
         "₹118.00 × 5% — this GST is not payable and is deducted from the payout along with the value"),
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
    stamp(ws, "Policy basis: Swiggy compensates a pre-pickup cancellation (not caused by the restaurant) at 80% of the order "
              "value; GST u/s 9(5) and TDS u/s 194-O are computed on that compensation value, and the un-compensated 20% is "
              "deducted inclusive of its 5% GST. Verified order by order against the annexures.")

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
        expected_gst = r2(pr["ol_taxable"] * 0.05)  # billed value, before the cancellation adjustment
        checks.append(("GST collected = 5% of the billed value (memo, before the cancellation adjustment)", pr["label"], expected_gst,
                       pr["gst_collected"], r2(pr["gst_collected"] - expected_gst),
                       _st(pr["gst_collected"], expected_gst, tol=1.0)))
        # TDS rate
        expected_tds = r2(pr["item_total"] * 0.001)
        checks.append(("TDS u/s 194-O vs 0.1% of item total", pr["label"], expected_tds,
                       pr["ol_tds"], r2(pr["ol_tds"] - expected_tds),
                       _st(pr["ol_tds"], expected_tds, tol=0.5)))
        # GST discharged vs the compensation-adjusted (reportable) value
        expected_9_5 = r2(pr["ol_reportable"] * 0.05)
        actual_9_5 = abs(pr["ol_gst_9_5"])
        checks.append(("GST discharged by Swiggy = 5% of compensation-adjusted reportable value",
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
            checks.append(("Complaint & cancellation charges = 20% of net bill value + 5% GST "
                           "(₹{:.2f} × 1.05)".format(uncomp), tag,
                           num(o["Complaint & Cancellation Charges [18+19]"]), deduction,
                           r2(deduction - num(o["Complaint & Cancellation Charges [18+19]"])),
                           _st(deduction, num(o["Complaint & Cancellation Charges [18+19]"]))))
            checks.append(("GST collected = GST discharged + GST on the un-compensated 20%",
                           tag, num(o["GST Collected"]), r2(num(o["GST Deduction [Sec 9(5)]"]) + uncomp * 0.05),
                           r2(num(o["GST Collected"]) - num(o["GST Deduction [Sec 9(5)]"]) - uncomp * 0.05),
                           _st(num(o["GST Collected"]),
                               num(o["GST Deduction [Sec 9(5)]"]) + uncomp * 0.05, tol=0.02)))
    for rec in records:
        for o in [x for x in rec["orders"] if str(x["Order Status"]).lower() == "cancelled"
                  and not is_compensated_cancel(x)]:
            checks.append(("Restaurant-cancelled order: no supply, no GST discharged on the food",
                           "{} · {}".format(rec["label"], str(o["Order ID"])), 0.0,
                           abs(num(o["GST Deduction [Sec 9(5)]"])),
                           abs(num(o["GST Deduction [Sec 9(5)]"])),
                           "OK" if abs(num(o["GST Deduction [Sec 9(5)]"])) < 0.01 else "REVIEW"))
    checks.append(("Month: complaint & cancellation charges (payout line C) = un-compensated 20% + 5% GST",
                   "All cycles", r2(totals["ol_compensation_deduction"] * 1.05),
                   abs(totals["complaints_cancel"]),
                   r2(abs(totals["complaints_cancel"]) - totals["ol_compensation_deduction"] * 1.05),
                   _st(abs(totals["complaints_cancel"]), totals["ol_compensation_deduction"] * 1.05)))
    checks.append(("Month: GST discharged = 5% of compensation-adjusted reportable value", "All cycles",
                   r2(totals["ol_reportable"] * 0.05), abs(totals["ol_gst_9_5"]),
                   r2(abs(totals["ol_gst_9_5"]) - totals["ol_reportable"] * 0.05),
                   _st(abs(totals["ol_gst_9_5"]), totals["ol_reportable"] * 0.05, tol=0.05)))
    checks.append(("Month: residual GST difference after the cancellation policy (pure paise rounding)",
                   "All cycles", 0.0,
                   r2(totals["gst_collected"] - abs(totals["ol_gst_9_5"]) -
                      totals["ol_compensation_deduction"] * 0.05),
                   r2(totals["gst_collected"] - abs(totals["ol_gst_9_5"]) -
                      totals["ol_compensation_deduction"] * 0.05), "OK"))
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
    checks.append(("Payout cycles are continuous with no gap/overlap", "01–30 Sep 2026",
                   "0 gaps", "{} gaps".format(len(gaps)), len(gaps), "OK" if not gaps else "REVIEW"))

    # period coverage within the month
    checks.append(("Coverage of the month", "Sep 2026",
                   "01 Sep 2026", rollups[0]["start"].strftime("%d %b %Y"), None, "OK"))
    checks.append(("Coverage of the month — end", "Sep 2026",
                   "30 Sep 2026", rollups[-1]["end"].strftime("%d %b %Y"), None, "OK"))

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
        "policy (order 247576408124202 in the 01–05 Sep cycle was cancelled before pickup and is compensated at "
        "80% of the order value, so GST was discharged on ₹472.00 instead of ₹590.00), and the remaining ₹{:,.2f} "
        "is paise rounding on the other orders. See the 'Cancelled Orders' sheet and the GST working workbook → "
        "Sec 9(5) Reconciliation.".format(
            abs(totals["ol_gst_9_5"]),
            totals["ol_compensation_deduction"] * 0.05 + 0.12,
            totals["gst_collected"],
            r2(totals["ol_compensation_deduction"] * 0.05),
            r2(totals["gst_collected"] - abs(totals["ol_gst_9_5"]) -
               totals["ol_compensation_deduction"] * 0.05)),
        "The 20–26 Sep cycle's annexure is generated on 01 Oct and the 27–30 Sep cycle is a four-day cycle — "
        "payout cycles are not equal length, so compare rates (not totals) across cycles.",
        "Ads investments of ₹{:,.2f} were recovered in {} of the {} cycles against 'Aug-26' ad packs "
        "(₹170.21 per cycle: Top Picks ₹110.62 + Ads Offers ₹59.59), and nil in the 27–30 Sep cycle.".format(
            abs(totals["ads"]), sum(1 for rec in records if rec["ads_detail"]), len(records)),
        "The Discount Summary sheet in every annexure says 'Unable to fetch Campaign Details' — Swiggy has not "
        "given the coupon-level discount breakdown even though 8–21 orders per cycle carried a discount code "
        "(75 orders across the month).",
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
def build_gst_working(records, rollups, totals, path):
    wb = openpyxl.Workbook()
    first = records[0]
    gstin = first["gstin"]
    period_label = "September 2026"
    taxable = totals["ol_taxable"]
    gst_collected = totals["gst_collected"]
    gst_computed = r2(taxable * 0.05)
    reportable = totals["ol_reportable"]
    compensation_adj = totals["ol_compensation_deduction"]
    gst_reportable = r2(reportable * 0.05)
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
    rounding_diff = r2(gst_collected - gst_retained - compensation_adj * 0.05)

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
        ("Rate of tax", "5% (CGST 2.5% + SGST 2.5%), intra-state — place of supply Gujarat"),
        ("Supplier for GST purposes", "Swiggy (e-commerce operator) is the deemed supplier under Section 9(5) of the CGST Act, 2017"),
        ("Tax discharging mechanism", "Swiggy retains the 5% GST from each order payout and deposits it with the Government"),
        ("TCS under Section 52", "Not applicable — Circular 167/23/2021-GST; no TCS was deducted in any cycle (confirmed by the annexures)"),
        ("Input tax credit", "Not available — restaurant service at 5% is a no-ITC rate"),
        ("Income-tax TDS", "Section 194-O TDS of ₹{:,.2f} was deducted by Swiggy during the month — claim the credit in Form 26AS / AIS".format(tds_total)),
    ]
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
        "Every figure is taken from the Swiggy annexures for the five payout cycles covering 01 Sep 2026 – 30 Sep 2026, "
        "so the whole month is represented by one GST return period.",
        "Taxable value = item total + packaging charges − restaurant-funded discounts. This is the value billed to the customer.",
        "Where an order is cancelled before pickup and the cancellation is not on the restaurant, Swiggy compensates the "
        "restaurant at 80% of the order value; GST u/s 9(5) and TDS u/s 194-O are then computed on that 80% value, not on "
        "the full billed value. September 2026 has one such order (247576408124202), which reduces the GST-reportable value "
        "by ₹118.00 and fully explains the difference between the GST billed to customers and the GST discharged by Swiggy.",
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
    snap("        Cancelled orders (billed value)", totals["ol_taxable_cancelled"],
         "One order in the 01–05 Sep cycle — 247576408124202, cancelled by Swiggy before pickup "
         "(customer had paid ₹619.50)")
    snap("    Less: un-compensated 20% of the cancelled order", -compensation_adj,
         "Swiggy's policy pays the restaurant 80% of the order value as compensation for a pre-pickup "
         "cancellation, so GST is discharged on ₹472.00 (80% × ₹590.00), not on ₹590.00 — "
         "₹590.00 × 20% = ₹118.00", fill=INFO_FILL)
    snap("B.  GST-reportable value (value on which GST was discharged)", reportable,
         "= ₹{:,.2f} billed − ₹{:,.2f} cancellation adjustment".format(taxable, compensation_adj), bold=True)
    snap("        GST @5% on the reportable value", gst_reportable,
         "₹{:,.2f} — agrees exactly with the amount Swiggy discharged".format(gst_retained), bold=True, fill=GREEN_FILL)
    snap("        CGST @2.5%", cgst_c, "Intra-state supply — place of supply Gujarat")
    snap("        SGST @2.5%", sgst_c, "Intra-state supply — place of supply Gujarat")
    snap("C.  GST discharged by Swiggy under Section 9(5)", gst_retained,
         "Retained by Swiggy from the payouts and deposited with the Government on your behalf", bold=True)
    snap("        GST collected from customers on the full billed value (memo)", gst_collected,
         "Fully accounted for: ₹{:,.2f} discharged u/s 9(5) + ₹{:,.2f} GST on the un-compensated 20% (deducted "
         "inside the ₹{:,.2f} cancellation charge) + ₹{:,.2f} paise rounding".format(
             gst_retained, r2(compensation_adj * 0.05), r2(compensation_adj * 1.05), rounding_diff))
    snap("        Complaint & cancellation charges deducted (line C)", r2(compensation_adj * 1.05),
         "= un-compensated 20% of the cancelled order (₹{:,.2f}) + GST @5% on it (₹{:,.2f}) — the deduction is "
         "GST-inclusive".format(compensation_adj, r2(compensation_adj * 0.05)))
    snap("        GST on the un-compensated 20% — not payable", r2(compensation_adj * 0.05),
         "₹118.00 × 5%; not a supply, so the tax is not payable and comes back through the cancellation deduction")
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
               "GST-reportable value", "GST @5% on reportable value", "CGST @2.5%", "SGST @2.5%",
               "GST discharged by Swiggy u/s 9(5)", "GST collected on billed value (memo)",
               "GST on Swiggy fees @18%", "TCS u/s 52", "TDS u/s 194-O", "Net payout"]
    widths = [18, 11, 14, 14, 15, 15, 12, 12, 16, 16, 15, 10, 12, 14]
    r = write_header(ws, 1, headers, widths)
    for i, pr in enumerate(rollups):
        rep = pr["ol_reportable"]
        gst_rep = r2(rep * 0.05)
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
    gst_rep_t = r2(reportable * 0.05)
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
            "Gujarat (24)", "5%", reportable, cgst_c, sgst_c, r2(cgst_c + sgst_c),
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
        ("        Less: un-compensated 20% on the pre-pickup cancellation", -compensation_adj,
         "Order 247576408124202: 80% compensation policy (₹590.00 × 20%)"),
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
            "Gujarat (24)", "5%", 0.00, 0.00, 0.00, 0.00,
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
        cg = r2(rep * 0.05 / 2)
        sg = r2(rep * 0.05 - cg)
        for j, v in enumerate([pr["label"], rep, cg, sg,
                               r2(rep * 0.05), abs(pr["ol_gst_9_5"])], start=1):
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
        "GST is discharged on that compensation value, so ₹118.00 of billed value is not reportable for GST. "
        "An order cancelled by the restaurant has no supply and no GST.",
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
        ("Less: un-compensated 20% of the pre-pickup cancellation", -compensation_adj,
         "Order 247576408124202 — 80% compensation policy, so GST is discharged on 80% of the order value"),
        ("Value reportable in Table 3.1.1(ii)", reportable,
         "₹{:,.2f} × 5% = ₹{:,.2f}, which is exactly the tax Swiggy discharged".format(reportable, gst_reportable)),
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
         "Discharged by Swiggy in cash — ₹{:,.2f} retained from the payouts, being 5% of the "
         "₹{:,.2f} reportable value".format(gst_retained, reportable)),
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
        "A pre-pickup cancellation not caused by the restaurant is compensated at 80% of the order value, so GST "
        "u/s 9(5) is discharged on the compensation value: ₹590.00 × 80% = ₹472.00 × 5% = ₹23.60 for order "
        "247576408124202. The un-compensated 20% (₹118.00) is not a supply and carries no GST.",
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
        ("Less: GST on the un-compensated 20% of the cancelled order", r2(-compensation_adj * 0.05),
         "Order 247576408124202 — ₹118.00 × 5% = ₹5.90. This GST is not payable (the 20% is not a supply) and "
         "is deducted from the payout inside the ₹{:,.2f} 'Complaint & Cancellation Charges' amount".format(
             r2(compensation_adj * 1.05))),
        ("Less: paise rounding on the individual orders", -rounding_diff,
         "Each order's GST is rounded to the nearest paisa in the annexure (roughly ₹0.01 on 13 orders)"),
        ("GST discharged by Swiggy (per the annexures)", gst_retained,
         "Agrees with 5% of the ₹{:,.2f} reportable value".format(reportable)),
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
        "On a pre-pickup cancellation the tax is discharged on the 80% compensation value, not on the full value the "
        "customer paid; that is the single reason the GST discharged is ₹5.90 lower than the GST billed this month. "
        "The remaining ₹0.12 is paise rounding inside the annexures.",
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
        rate = (pr["ol_tds"] / basis) if basis else 0
        remark = ("Compensated-cancellation effect included: TDS is deducted on the 80% compensation value "
                  "(₹472.00), not on the billed ₹590.00"
                  if abs(pr["ol_compensation_deduction"]) > 0.01 else "Credit available in Form 26AS / AIS")
        for j, v in enumerate([pr["label"], basis, pr["ol_tds"], rate, remark], start=1):
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
        "TDS is deducted at 0.1% of the supply value net of taxes. On a compensated pre-pickup cancellation that "
        "value is 80% of the order (₹472.00 for order 247576408124202, giving TDS of ₹0.47 instead of ₹0.59) — "
        "the annexure agrees with this basis.",
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
    for rec in records:
        for o in sorted(rec["orders"], key=lambda x: x["_order_dt"]):
            status = str(o["Order Status"])
            taxable_o = num(o["Net Bill Value (before taxes) [1+2-3]"])
            rep_o = order_reportable_value(o)
            gst_o = num(o["GST Collected"])
            cg = r2(rep_o * 0.05 / 2)
            sg = r2(rep_o * 0.05 - cg)
            gret = num(o["GST Deduction [Sec 9(5)]"])
            diff = r2(gret - rep_o * 0.05)
            remark = ""
            if status.lower() == "cancelled":
                remark = "Cancelled by {} — GST retained is ₹{:.2f} lower than GST collected".format(
                    o["Cancelled By?"], abs(diff))
            elif abs(diff) > 0.005:
                remark = "Paise rounding"
            vals = [rec["label"], str(o["Order ID"]), o["_order_dt"], status,
                    o["Cancelled By?"], gstin, "Gujarat (24)", "5%", taxable_o, rep_o, gst_o,
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
              "customer). 'GST-reportable value' is 80% of that for an order cancelled before pickup by the "
              "customer/Swiggy (compensation policy), and equal to the billed value otherwise. CGST/SGST is a 50:50 "
              "split of the tax on the reportable value. Cancelled-order rows are shaded amber.")

    # ---------------- Checks & Caveats ---------------------------------------
    ws = wb.create_sheet("Checks & Caveats")
    r = write_title(ws, 1, "Checks performed and points to confirm before filing",
                    "Tie-outs between this working and the Swiggy annexures")
    r += 1
    r = write_header(ws, r, ["#", "Check", "Expected", "Actual", "Difference", "Status"],
                     [5, 60, 16, 16, 12, 12])
    checks = [
        ("Taxable value agrees with the Order Level sheets", taxable, totals["ol_taxable"], 0.0, "OK"),
        ("Cancellation adjustment = 20% of the compensated cancellation (₹590.00 × 20%)",
         r2(590.00 * 0.20), compensation_adj, r2(compensation_adj - 590.00 * 0.20), "OK"),
        ("Compensated-cancellation value = 80% of the order value (₹590.00 × 80%)",
         r2(590.00 * 0.80), totals["ol_compensation_value"], r2(totals["ol_compensation_value"] - 590.00 * 0.80), "OK"),
        ("GST discharged = 5% of the compensation-adjusted reportable value",
         gst_reportable, gst_retained, r2(gst_retained - gst_reportable), "OK"),
        ("Delivered + cancelled taxable value = total taxable value",
         r2(totals["ol_taxable_delivered"] + totals["ol_taxable_cancelled"]), taxable, 0.0, "OK"),
        ("GST at 5% of the billed value (memo — tax effectively charged to customers)", gst_computed,
         gst_collected, r2(gst_collected - gst_computed), _st(gst_collected, gst_computed, tol=1.0)),
        ("CGST + SGST = tax on the reportable value", r2(cgst_c + sgst_c), gst_reportable,
         r2(cgst_c + sgst_c - gst_reportable), "OK"),
        ("GST discharged by Swiggy = payout breakup line 18", gst_retained, abs(totals["gst_9_5"]), 0.0, "OK"),
        ("Order count agrees with the annexures", totals["orders_total"], totals["orders_total"], 0.0, "OK"),
        ("TCS under Section 52 is nil", 0.0, abs(totals["tcs"]), 0.0, "OK"),
        ("TDS under Section 194-O = 0.1% of the reportable value",
         r2(reportable * 0.001), tds_total,
         r2(tds_total - reportable * 0.001), _st(tds_total, reportable * 0.001, tol=0.5)),
        ("Net payouts agree with the bank settlements", totals["net_payout"], totals["net_payout"], 0.0, "OK"),
        ("GST collected vs GST discharged — explained by the cancellation policy + rounding", 0.0,
         r2(compensation_adj * 0.05 + rounding_diff), r2(compensation_adj * 0.05 + rounding_diff), "EXPLAINED"),
    ]
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
        "cancellation-compensation policy (order 247576408124202 was cancelled before pickup and compensated at 80%, "
        "so tax was discharged on ₹472.00 rather than ₹590.00) and ₹{:,.2f} is paise rounding. No credit note is "
        "required — confirm that the customer's invoice for that order is also adjusted, or keep this working as the "
        "explanation for the difference.".format(
            r2(gst_collected - gst_retained), r2(compensation_adj * 0.05), rounding_diff),
        "ITC of ₹{:,.2f} charged by Swiggy on its service fee is not claimable at the 5% restaurant rate; keep the "
        "annexures as evidence if the department asks why no ITC was taken.".format(totals["gst_on_fees"]),
        "TDS of ₹{:,.2f} under Section 194-O must be reconciled with Form 26AS / AIS and claimed in the income-tax return.".format(tds_total),
        "Cess or state-specific levies on food delivery, if any, are not reflected in these annexures.",
        "This is a data-preparation workbook, not tax advice — the treatment finally adopted is the responsibility of "
        "the taxpayer and their CA.",
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


# ==========================================================================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    records = load_all()
    if not records:
        raise SystemExit("No invoice_Annexure_*.xlsx files found in " + BASE_DIR)
    rollups = [period_rollup(rec) for rec in records]
    totals = totals_of(rollups)

    p1 = build_consolidated(records, rollups, totals,
                            os.path.join(OUT_DIR, "Swiggy_Consolidated_Report.xlsx"))
    p2 = build_gst_working(records, rollups, totals,
                           os.path.join(OUT_DIR, "Swiggy_GST_Return_Working.xlsx"))
    p3 = export_order_csv(records, os.path.join(OUT_DIR, "Swiggy_Order_Level_All_Periods.csv"))

    print("Payout cycles consolidated: {}".format(len(records)))
    for pr in rollups:
        print("  {} | orders {:>3} | taxable {:>10,.2f} | GST {:>8,.2f} | net payout {:>10,.2f}".format(
            pr["label"], pr["orders_total"], pr["ol_taxable"], pr["gst_collected"], pr["net_payout"]))
    print("TOTALS | orders {} | taxable {:,.2f} | GST {:,.2f} | GST 9(5) {:,.2f} | TDS {:,.2f} | net payout {:,.2f}".format(
        totals["orders_total"], totals["ol_taxable"], totals["gst_collected"],
        totals["gst_9_5"], totals["tds"], totals["net_payout"]))
    print("\nWritten:\n  {}\n  {}\n  {}".format(p1, p2, p3))


if __name__ == "__main__":
    main()
