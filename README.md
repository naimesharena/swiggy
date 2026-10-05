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
- **Order Level (All Periods)** — every order in one filterable table, tagged with its payout cycle (cancelled rows in amber), with a `GST-reportable value` column
- **Cancelled Orders** — the 80% compensation policy worked order by order, and its effect on the month's GST
- **Ads & Adjustments** — growth investments / adjustments deducted at restaurant level
- **Discounts**, **Complaints** — campaign and complaint buckets reported by Swiggy
- **Reconciliation** — 81 automated tie-out checks with status and exceptions
- **Field Definitions** — glossary plus notes on the sign conventions

### 2. GST working — sheets

- **Read Me** — basis, legal position, and the two reporting views
- **GST Snapshot** — the month's position on one page
- **Period-wise GST** — the same numbers split by payout cycle
- **GSTR-1 Working** — rate-wise outward supply working on the reportable value (Table 14 / B2CS)
- **GSTR-3B Working** — Table 3.1.1(ii) working, ITC position, and an annexure-to-return reconciliation
- **Sec 9(5) Reconciliation** — GST billed vs GST actually discharged by Swiggy, with the cancellation adjustment and order-wise exceptions
- **TCS & TDS** — why TCS is nil and the Sec 194-O income-tax TDS reconciliation
- **Order-wise GST Register** — all 95 orders with taxable value, CGST/SGST split and tax discharged
- **Checks & Caveats** — 14 tie-outs performed and points to confirm with your CA

## Headline figures (01–30 Sep 2026)

| | |
|---|---|
| Orders | 95 (92 delivered, 3 cancelled) across 5 payout cycles |
| Taxable value (item total + packaging − discounts) | ₹46,236.33 |
| Billed value (all orders) | ₹46,236.33 |
| Less: un-compensated 20% of the pre-pickup cancellation | ₹118.00 |
| **GST-reportable value** | **₹46,118.33** |
| GST discharged by Swiggy u/s 9(5) @5% | ₹2,305.92 (CGST ₹1,152.96 + SGST ₹1,152.96) |
| GST billed to customers on the full value (memo) | ₹2,311.94 |
| TCS u/s 52 | Nil (not applicable to Sec 9(5) supplies) |
| TDS u/s 194-O | ₹46.14 (0.1% of item total) |
| Gross collected from customers | ₹48,548.27 |
| GST on the un-compensated 20% (deducted, not payable) | ₹5.90 |
| **Net payouts credited by Swiggy** | **₹30,139.91** |

## The 80% cancellation-compensation policy (explains the GST difference)

When an order is cancelled before pickup and the restaurant is not at fault, Swiggy compensates the
restaurant at **80% of the order value** — so the restaurant is taxed on that compensation, not on the
full value the customer paid. For order `247576408124202` (01–05 Sep cycle):

| Basis | Value | GST @5% | Total |
|---|---|---|---|
| Net bill value of the order | ₹590.00 | ₹29.50 | ₹619.50 paid by the customer |
| **Compensation side** — 80% | ₹472.00 | ₹23.60 | ₹495.60 credited |
| **Un-compensated side** — 20% | ₹118.00 | ₹5.90 | **₹123.90 deducted** |

Every rupee closes, three ways:

- **Value:** ₹472.00 + ₹118.00 = ₹590.00 = net bill value
- **Cash:** ₹495.60 + ₹123.90 = ₹619.50 = what the customer paid
- **GST:** ₹23.60 discharged u/s 9(5) + ₹5.90 not payable = ₹29.50 = GST collected

The ₹123.90 appears in the annexure as *"Customer Cancellations"* / *"Complaint & Cancellation Charges
[18+19]"* — it is the **GST-inclusive** un-compensated 20% (₹118.00 + 5% = ₹123.90). Each annexure column
then ties to the policy:

| Annexure column | Value | Derivation |
|---|---|---|
| GST Deduction (u/s 9(5)) | ₹23.60 | 5% × ₹472.00 (the compensation value) |
| Commission charged on | ₹495.60 | 80% × ₹619.50 |
| Customer Cancellations / Complaint & Cancellation [18+19] | ₹123.90 | ₹118.00 × 1.05 |
| TDS u/s 194-O | ₹0.47 | 0.1% × ₹472.00 |

Applied to the whole month: ₹46,236.33 billed − ₹118.00 (the un-compensated 20%) = **₹46,118.33** × 5% =
**₹2,305.92**, which is exactly the GST Swiggy discharged. The ₹6.02 difference between the GST billed to
customers and the GST discharged is therefore fully explained — **₹5.90 policy + ₹0.12 paise rounding** — and
nothing needs to be recovered from Swiggy.

Orders cancelled **by the restaurant** are the opposite case: no compensation and no supply, so no GST arises
on the food at all — only the cancellation charge, which is quoted *exclusive* of GST with 18% added on top
(₹82.50 + ₹14.86 and ₹252.25 + ₹45.40, both recovered through "Total Swiggy Fees").

Both workbooks and the CSV carry this logic: a `GST-reportable value` column (80% for a compensated
cancellation, full value otherwise), a dedicated **Cancelled Orders** sheet in the consolidation report
(section A compensated orders, section B restaurant-cancelled orders, section C annexure-column mapping,
section D month effect), and a compensation-adjusted Section 9(5) reconciliation in the GST working.

## Things worth acting on

1. **Ads deductions** — ₹170.21 (Top Picks ₹110.62 + Ads Offers ₹59.59) was recovered in 4 of the 5 cycles,
   ₹680.84 for the month, against "Aug-26" packs.
2. **Discount details missing** — every annexure says *"Unable to fetch Campaign Details"*, even though
   75 orders carried a discount code. Worth chasing.
3. **Direct sales are not in the annexures** — add dine-in / takeaway / own-delivery figures before filing
   GSTR-1 Table 7 and GSTR-3B Table 3.1(a).
4. **Check the customer's invoice for the cancelled order** — the customer was billed GST on ₹590.00 while only
   ₹472.00 of value was taxed. Keep this working as the explanation for the difference.
5. **Paise differences** — a few cycles show ₹0.01–0.17 differences between their own columns; these are inside
   Swiggy's annexures and are disclosed, not adjusted.

## Re-running

```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl
.venv/bin/python build_reports.py            # interactive file-selection menu
.venv/bin/python build_reports.py --all      # every annexure in this folder
```

Run it with no arguments and it shows what it found and asks which files to use:

```
  #   File                                            Restaurant (ID)             Payout period           Orders   Payout (Rs)
  ----------------------------------------------------------------------------------------------------------------------------
  1   invoice_Annexure_544662_01102026_1790849827...  Cheesecake Mills (544662)   20 September - 26 Sep   10        2,972.66
  2   invoice_Annexure_544662_03102026_1791048772...  Cheesecake Mills (544662)   27 September - 30 Sep   17        4,819.24
  3   invoice_Annexure_544662_09092026_1788952060...  Cheesecake Mills (544662)   01 September - 05 Sep   19        6,845.46
  4   invoice_Annexure_544662_16092026_1789552968...  Cheesecake Mills (544662)   06 September - 12 Sep   25        7,479.09
  5   invoice_Annexure_544662_23092026_1790175033...  Cheesecake Mills (544662)   13 September - 19 Sep   24        8,023.46

  Select the files to include:
      all            every valid annexure listed above
      1 3 5          particular numbers (spaces or commas)
      1-3            a range of numbers
      <path>         a file, folder or glob (e.g. D:/annexures/*.xlsx)
      f              change the folder to scan
      q              quit
```

Any file that is not a Swiggy annexure (a previous report, a template) is shown as
`-- not an annexure --` with the reason, and is simply left out.

### Command line

| Option | Effect |
|---|---|
| `files …` | annexure files, folders or globs (`a.xlsx`, `D:/annexures`, `D:/annexures/*.xlsx`) |
| `--dir FOLDER` / `-d` | folder to scan for annexures (default: the script's folder) |
| `--all` / `-a` | use every annexure found, no menu |
| `--list` / `-l` | show what would be read, then exit |
| `--outdir DIR` / `-o` | where to write the reports (default: `reports/`) |
| `--gst-rate N` | force the GST rate instead of auto-detecting it |
| `--no-csv` | skip the order-level CSV |

Examples:

```bash
python build_reports.py --list                                   # what is in this folder?
python build_reports.py --all                                    # all 5 cycles -> reports/
python build_reports.py --dir "D:/Swiggy Oct" --all --outdir out # another month
python build_reports.py invoice_Annexure_*_09092026_*.xlsx       # just one week
```

### Using it for other restaurants

The tool does not hard-code this restaurant. For every file it reads the identity from the **Summary** sheet
(restaurant name, outlet, Restaurant ID, GSTIN), takes the period from the file name
(`invoice_Annexure_<id>_<ddmmyyyy>_<…>.xlsx`) and the year from the Summary sheet, and works out the GST rate
from the annexure data (falling back to 5%). Nothing else is assumed.

- One restaurant selected → outputs go straight into `reports/` (as here).
- Several restaurants selected → each gets its own folder `reports/<Restaurant_Name>/`, plus a small
  `reports/index.txt` listing what was produced where.
- The GSTIN's state code drives the state name and place of supply on every sheet (Gujarat 24, Maharashtra 27 …),
  so an outlet in another state is labelled correctly without any code change.
- The Order Level sheet is read by **column heading**, not by column position, so a client file with an extra
  column inserted, a column renamed, or an optional column missing still reads correctly. If a column the
  reports genuinely need is absent, the file is reported as skipped with the column names listed instead of
  producing wrong figures.
- The same annexure kept in two folders (for example `creamy/` and the project root) is detected as a
  duplicate and read **once**, so totals are never double-counted. Duplicates are shown in the file list as
  `duplicate of #N`, and reported when the run starts.
- A file that cannot be read is reported (with the missing sheets) and skipped; the other restaurants are
  still processed.

Inputs are always read **by column and label, never by fixed row**, and the Order Level columns are matched
by heading, so Swiggy can change the annexure layout slightly (or a client can be on a newer template) without
breaking anything.

> **Keeping one copy of each annexure.** If the same weekly file is stored twice, put both copies in the
> selection and the tool will read one and say which file it treated as the duplicate. Removing one of the
> two copies keeps the folder tidy, but it is not necessary.

> Food delivery through Swiggy is a local (intra-state) supply, so the tax is split CGST/SGST on the restaurant's
> own state code. The annexures do not carry the customer's delivery state; if you ever need IGST treatment for
> an inter-state delivery, that split has to be made outside this working.

> The GST workbook is a data-preparation aid built from Swiggy's data. It is not tax advice —
> confirm the reporting treatment with your CA before filing.
