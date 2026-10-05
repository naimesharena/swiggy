# Swiggy payout annexures — consolidation & GST return working

This folder holds the weekly Swiggy payout annexures for **Cheesecake Mills, Vesu, Surat**
(Restaurant ID 544662, GSTIN 24AAQFC3084Q1ZT) for **01–30 September 2026**, and the two
workbooks built from them.

## Deliverables (`reports/`)

| File | What it is |
|---|---|
| `Swiggy_Consolidated_Report.xlsx` | **Consolidated data report** — all 5 payout cycles merged into one data set |
| `Swiggy_GST_Return_Working.xlsx` | **GST return filing working** — period-wise tax working, GSTR-1/GSTR-3B support, Sec 9(5) and TCS/TDS reconciliation |
| `Swiggy_Order_Level_All_Periods.csv` | Flat CSV of all 95 orders (for Tally / Excel / BI) |

### 1. Consolidation report — sheets

- **Cover** — scope, source files, basis of preparation
- **Period Summary** — one line per payout cycle: orders, sales, fees, ads, taxes and net payout, with grand totals
- **Payout Breakup** — Swiggy's full payout-breakup structure (lines A→G with all sub-lines) side by side for every cycle
- **Delivered vs Cancelled** — how each cycle's net payout splits between delivered orders, cancelled orders and ads
- **Order Level (All Periods)** — every order in one filterable table, tagged with its payout cycle (cancelled rows in amber)
- **Ads & Adjustments** — growth investments / adjustments deducted at restaurant level
- **Discounts**, **Complaints** — campaign and complaint buckets reported by Swiggy
- **Reconciliation** — 67 automated tie-out checks with status and exceptions
- **Field Definitions** — glossary plus notes on the sign conventions

### 2. GST working — sheets

- **Read Me** — basis, legal position, and the two reporting views
- **GST Snapshot** — the month's position on one page
- **Period-wise GST** — the same numbers split by payout cycle
- **GSTR-1 Working** — rate-wise outward supply working (Table 14 / B2CS)
- **GSTR-3B Working** — Table 3.1.1(ii) working, ITC position, and an annexure-to-return reconciliation
- **Sec 9(5) Reconciliation** — GST collected from customers vs GST actually retained and deposited by Swiggy, with order-wise exceptions
- **TCS & TDS** — why TCS is nil and the Sec 194-O income-tax TDS reconciliation
- **Order-wise GST Register** — all 95 orders with taxable value, CGST/SGST split and tax discharged
- **Checks & Caveats** — tie-outs performed and points to confirm with your CA

## Headline figures (01–30 Sep 2026)

| | |
|---|---|
| Orders | 95 (92 delivered, 3 cancelled) across 5 payout cycles |
| Taxable value (item total + packaging − discounts) | ₹46,236.33 |
| GST @5% (CGST 2.5% + SGST 2.5%) | ₹2,311.94 |
| GST discharged by Swiggy u/s 9(5) | ₹2,305.92 |
| TCS u/s 52 | Nil (not applicable to Sec 9(5) supplies) |
| TDS u/s 194-O | ₹46.14 (0.1% of item total) |
| Gross collected from customers | ₹48,548.27 |
| **Net payouts credited by Swiggy** | **₹30,139.91** |

## Things worth acting on

1. **₹6.02 GST shortfall** — Swiggy collected ₹2,311.94 from customers but retained/discharged only
   ₹2,305.92. ₹5.90 of this is on cancelled order `247576408124202` (01–05 Sep cycle, GST ₹29.50 collected
   vs ₹23.60 retained); the rest is paise rounding. Raise it with Swiggy.
2. **Ads deductions** — ₹170.21 (Top Picks ₹110.62 + Ads Offers ₹59.59) was recovered in 4 of the 5 cycles,
   ₹680.84 for the month, against "Aug-26" packs.
3. **Discount details missing** — every annexure says *"Unable to fetch Campaign Details"*, even though
   75 orders carried a discount code. Worth chasing.
4. **Direct sales are not in the annexures** — add dine-in / takeaway / own-delivery figures before filing
   GSTR-1 Table 7 and GSTR-3B Table 3.1(a).
5. **Paise differences** — a few cycles show ₹0.01–0.17 differences between their own columns; these are inside
   Swiggy's annexures and are disclosed, not adjusted.

## Re-running

```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl
.venv/bin/python build_reports.py
```

`build_reports.py` reads every `invoice_Annexure_*.xlsx` in this folder (by column/label, not by fixed row),
so new weekly files can simply be dropped in and the reports regenerated. Outputs are written to `reports/`.

> The GST workbook is a data-preparation aid built from Swiggy's data. It is not tax advice —
> confirm the reporting treatment with your CA before filing.
