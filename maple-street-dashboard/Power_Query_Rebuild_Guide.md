# Rebuild Guide: AP Aging & Cash Flow Dashboard (Power Query + Power Pivot)

This walks you through recreating the "Maple Street Custom Furniture" sample project
yourself, in your own Excel, using the three raw CSV files provided
(`Vendor_Invoice_Tracker.csv`, `QBO_GL_Export.csv`, `Bank_Transactions_Export.csv`).

Doing this hands-on — not just looking at the finished workbook — is the point:
it means you can genuinely walk a client through your process in an interview or
proposal call, not just show them a finished file.

Reporting date used throughout: **June 30, 2026**.

---

## Part 1 — Get & clean the Vendor Invoice Tracker

1. **Excel > Data tab > Get Data > From File > From Text/CSV** → select
   `Vendor_Invoice_Tracker.csv` → click **Transform Data** (not "Load") to open
   Power Query Editor.
2. **Trim/Clean the Vendor column**: Select the `Vendor` column → Transform tab →
   **Format > Trim**, then **Format > Clean**. This removes stray leading/trailing
   spaces and non-printing characters, but won't fix casing/spelling variants yet.
3. **Standardize casing** as an intermediate step: right-click `Vendor` column →
   Transform > **Capitalize Each Word** (or lowercase everything first if you'd
   rather match on a fully-lowercased key).
4. **Normalize vendor name variants** — this is the real cleanup step. You have two
   good options:
   - **Option A (simple, good for a small vendor list):** Add a Custom Column with
     nested logic:
     ```
     = if Text.Contains([Vendor], "Coastal", Comparer.OrdinalIgnoreCase) then "Coastal Hardwood Supply"
       else if Text.Contains([Vendor], "Blue Ridge", Comparer.OrdinalIgnoreCase) then "Blue Ridge Hardware Co."
       else if Text.Contains([Vendor], "Piedmont", Comparer.OrdinalIgnoreCase) then "Piedmont Finishes & Stain"
       else if Text.Contains([Vendor], "Guilford", Comparer.OrdinalIgnoreCase) then "Guilford Packaging Solutions"
       else if Text.Contains([Vendor], "Apex", Comparer.OrdinalIgnoreCase) then "Apex Freight & Shipping"
       else if Text.Contains([Vendor], "Kernersville", Comparer.OrdinalIgnoreCase) then "Kernersville Utilities"
       else if Text.Contains([Vendor], "Triad", Comparer.OrdinalIgnoreCase) then "Triad Equipment Service"
       else if Text.Contains([Vendor], "Summit", Comparer.OrdinalIgnoreCase) then "Summit Business Insurance"
       else [Vendor]
     ```
   - **Option B (more scalable, what you'd actually do for a real client with 50+
     vendors):** Build a small two-column **Vendor Mapping** reference table
     (`Raw Name` → `Canonical Name`) as its own query, then **Merge Queries** (left
     outer join) on a lowercased/trimmed key. This is the more "real" Power Query
     skill to demonstrate — it's how you'd handle this at scale.
5. **Parse the date columns.** `Invoice Date`, `Due Date`, and `Payment Date` arrive
   in inconsistent formats (`01/24/2026`, `2026-01-24`, `11-Feb-26`). Power Query's
   default `Date.FromText` will fail on mixed formats in one column, so add a custom
   column that tries each format in turn:
   ```
   = try Date.FromText([Invoice Date], "en-US")
     otherwise try Date.FromText(Text.Replace([Invoice Date], "-", "/"), "en-US")
     otherwise null
   ```
   Repeat for `Due Date` and `Payment Date`. Then remove the original text columns
   and rename the parsed ones.
6. **Remove duplicate rows.** Select the `Vendor` (normalized), `Invoice Date`,
   `Due Date`, and `Amount` columns → right-click → **Remove Duplicates**. This is
   what catches the manual double-entry errors in the tracker.
7. **Close & Load To...** → Only Create Connection (you'll load it to the Data
   Model in Part 4).

---

## Part 2 — Clean the QuickBooks GL Export

1. Repeat the Get Data > From Text/CSV import for `QBO_GL_Export.csv`.
2. Trim/Clean the `Vendor` and `Memo` columns (the raw export has extra padding
   spaces around the memo text).
3. Apply the same vendor normalization logic from Part 1, step 4.
4. Fill blank `Account` values: select the `Account` column → right-click → **Fill
   Down** only works if blanks should inherit the row above, which isn't the case
   here — instead, add a custom column that maps normalized vendor → account
   (mirroring the `VENDOR_TO_ACCOUNT` logic), so every row gets a correct GL account
   regardless of whether the export left it blank.
5. Close & Load To Connection Only.

---

## Part 3 — Clean the Bank Transactions Export

1. Import `Bank_Transactions_Export.csv`.
2. Parse `Post Date` the same way as Part 1 step 5 (it mixes `2026-02-26` and
   `02/01/26` style formats).
3. Add a custom column `Transaction Type` = `if [Amount] < 0 then "Outflow" else "Inflow"`.
4. Add a `Month` column: `= Date.ToText([Post Date], "yyyy-MM")`.
5. Close & Load To Connection Only.

---

## Part 4 — Build the data model & the visuals

1. **Data tab > Queries & Connections** — right-click each of your three cleaned
   queries → **Load To... > Add to Data Model** (this is Power Pivot).
2. **Power Pivot tab > Manage Data Model** to open the Power Pivot window.
3. Add a couple of calculated measures using DAX, e.g.:
   ```
   Total Open AP := CALCULATE(SUM(Invoices[Amount]), Invoices[Paid] = "N")
   90+ Days Past Due := CALCULATE(SUM(Invoices[Amount]), Invoices[Aging Bucket] = "90+ Days")
   ```
4. Back in Excel, **Insert > PivotTable** (choose "Use this workbook's Data Model")
   to build the AP Aging summary and Top Vendors summary tables.
5. **Insert > PivotChart** from each PivotTable to get the aging bar chart and the
   vendor ranking chart. For the cash flow trend, build a PivotTable of the bank
   query grouped by Month and Transaction Type, then chart it.
6. Arrange the PivotCharts and a few KPI cells (just formulas referencing your
   PivotTable totals, styled with a light fill) on a dedicated "Dashboard" sheet,
   and hide the gridlines (**View tab > uncheck Gridlines**) for a cleaner look.

---

## What to say about this in a client conversation

If a client asks about your process, the short version is: *"I import from wherever
the data lives — QuickBooks exports, bank downloads, manual trackers — clean and
standardize it in Power Query so vendor names, dates, and duplicates stop causing
errors, then build the reporting layer in Power Pivot so it refreshes with one
click instead of being rebuilt by hand every month."* That's exactly what this
sample project demonstrates, and it's the same pattern behind your real
achievement of unifying AP data and cutting reporting time in your last role.
