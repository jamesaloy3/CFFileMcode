Here’s a clear, “operator + maintainer” README for your Power Query setup.

# Power Query Folder Import: Hotel IS/CF → Central Facts Table

## 1) What this does (high level)

* Scans **only the top-level files** in a folder (ignores subfolders).
* Filters to an **allow-list of hotel codes** (from an Excel table named `HotelCodes`).
* For each allowed workbook, reads:

  * **`2025 IS`** (Income Statement)
  * **`2025 CF`** (Cash Flow)
* Skips the first **4 rows** (headers are on row 5), **promotes headers**, and then:

  * **Unpivots all non-label columns** to a tidy shape.
  * **Parses column headers as real Excel dates** and keeps **only** the dates you specify in `PQ_Dates` (e.g., E6\:P6 on your control sheet).
  * Recognizes **Allocate to Reserve Fund** even if it’s labeled in **Col D** or identified by the **Col B account code `46000-110`**.
* Returns a **WIDE** matrix by default (Hotel | Metric | date columns), with a **LONG** facts table also available if you change the final `in`.

Why? So each hotel sheet (named with its 4–5 digit code) can pull values by `(Hotel, Metric, Date)` with simple formulas, and **one refresh updates all hotels**.

---

## 2) Prerequisites & naming

On your **control** sheet:

* Cell **B2**: folder path → define named range **`PQ_Folder`** (e.g., `C:\Hotels\Monthly\`).
* Range **E6\:P6**: the target report dates (must be real Excel dates) → define named range **`PQ_Dates`**.
* A 1-column **Excel Table** named **`HotelCodes`** with header `Hotel` and one code per row (the first 4–5 digits of each file name).

> Important: Workbooks must contain sheets named exactly **`2025 IS`** and **`2025 CF`** with date headers in row 5.

---

## 3) What the query returns

* **Primary output (default)**: `Hotels_Metrics_2025_Wide`
  Columns: `Hotel | Metric | 1/31/2025 | 2/28/2025 | …`
* **Also available**: `Hotels_Metrics_2025_Long`
  Columns: `Hotel | Metric | Date | Value`
  (To use: duplicate the query and change the last line to return the “Long” step.)

Use the **Long** table with `SUMIFS` on your hotel sheets, or the **Wide** table for matrix views.

---

## 4) How the code works (step-by-step)

### 4.1 Parameters & date list

* Pulls `PQ_Folder` and the entire `PQ_Dates` range.
* **Flattens** the `PQ_Dates` range (horizontal or vertical) and converts every cell to a **true `date`**.
  This ensures matching is based on real dates, not text.

### 4.2 Allow-list

* Reads table `HotelCodes` → list of allowed hotel codes (digits).
* Derives a **Hotel code** from each file name by taking the **leading digits (up to 5)**.
  Example: “`12345 Hotel A.xlsx`” → hotel code `12345`.

### 4.3 Folder scan (top-level only)

* Uses `Folder.Contents(PQ_Folder)` → **not recursive**, so subfolders are ignored.
* Keeps **file** rows (`Attributes[Directory] = false`) and only **.xlsx/.xlsm**.
* Adds the derived **Hotel** code and filters to rows where `Hotel` is in **`HotelCodes`** and has length ≥ 4.

### 4.4 Robust sheet reader

For each workbook (allowed file):

* Reads binary with **`Binary.Buffer(File.Contents(...))`** (helps speed/stability).
* Opens with `Excel.Workbook`, looks for the sheet **by name** (`2025 IS`, `2025 CF`).
* **Skips 4 rows** and **promotes headers**.

> If a file is missing a sheet or has an unexpected structure, the code returns an **empty table** for that file instead of erroring.

### 4.5 Normalize 2025 IS

Purpose: reshape raw IS into rows of `Metric | Date | Value`.

Key logic:

* **Rename** the first 1–4 columns to `ColA..ColD` (only if present). These act as “label columns”.
* **Unpivot all other columns** into two columns: `DateHeader` and `Value`.
  This allows dates to be anywhere on the right, with gaps or extra columns.
* **Metric detection (null-safe):**

  * If **Col B** (account code) = `46000-110` **OR** **Col D** label contains “Allocate to Reserve Fund” → `Metric = "Allocate to Reserve Fund"`.
  * Else if **Col C or D** contain “USALI EBITDA” → `Metric = "USALI EBITDA"`.
  * Else if **Col C or D** contain “Gross Operating Profit” → `Metric = "Gross Operating Profit (Loss)"`.
  * Else → use **Col B** as the metric label (typical IS row labels).
* **Filter to the 3 IS metrics** you want.
* **Parse headers to true dates** (tolerant of text/serials/datetimes).
* **Keep rows whose `DateHeader` exists in your `PQ_Dates` date list** (safe boolean filter).
* **Coerce `Value` to number.**
* Return **`Metric | Date | Value`**.

### 4.6 Normalize 2025 CF

Very similar to IS, with one difference:

* The CF **label** source is **Col C**, and the **kept metrics** are:

  * `Debt Service`
  * `Cash Generated (Used) After Debt Service`.

All other steps (unpivot, date parse, target date filter, numeric coercion) are the same.

### 4.7 Combine & shape

* **ProcessFile** returns IS + CF rows for one hotel and adds the `Hotel` code column.
* All processed files are **combined** into one **Long** table `Hotels_Metrics_2025_Long`.
* A **Wide** table is built by pivoting the `Date` column to one column per date, preserving your metric order.

---

## 5) Error-handling & hardening (why it’s stable)

* **Null-safe** everywhere: we use `try`, `Record.FieldOrDefault`, and test types (e.g., `Value.Is([DateHeader], type date)`).
* If a sheet or workbook is missing/empty/misaligned:

  * The normalizers return an **empty table** with the correct schema.
  * The folder loop still completes; you still get all available data.
* No recursive folders: avoids accidentally pulling similarly-named files in subfolders.
* **Date matching by true dates**: avoids header text formatting pitfalls.

---

## 6) Performance tips

* Keep the folder local or on a fast share.
* Keep the **allow-list** (`HotelCodes`) tight to avoid opening unnecessary files.
* In Excel → **Data** → **Query Options**: uncheck “Detect data types and relationships” to skip extra inference.
* If source files are large, consider:

  * Removing unused sheets/regions from them.
  * Splitting historic vs. current files into separate folders.

---

## 7) Customization knobs

* **Sheet names:** change `"2025 IS"` / `"2025 CF"` in `ProcessFile` if your naming rolls forward annually.
* **Skip rows:** change `skipRows` (currently **4**) if your date headers move.
* **Allocate detection:** add more account codes or label variants in the IS normalizer.
* **Metrics set:** add/remove metric names in the `keepMet` filters for IS and CF.
* **Date formats:** the pivot uses `M/d/yyyy` text for stable column names. You can change the `Date.ToText` format string to your preferred format.

---

## 8) Using the results on hotel sheets

If **each sheet is named** with its hotel code (e.g., `12345`):

* Put your row labels in **D9 / D13 / D17 / D21 / D25**:

  * Gross Operating Profit (Loss)
  * USALI EBITDA
  * Debt Service
  * Allocate to Reserve Fund
  * Cash Generated (Used) After Debt Service
* Dates in **E6\:P6** (same as `PQ_Dates`).
* In **E9**, copy across and down (uses the **Long** table):

  ```excel
  =SUMIFS(Hotels_Metrics_2025_Long[Value],
          Hotels_Metrics_2025_Long[Hotel],  MID(CELL("filename",A1),FIND("]",CELL("filename",A1))+1,99),
          Hotels_Metrics_2025_Long[Metric], $D9,
          Hotels_Metrics_2025_Long[Date],   E$6)
  ```

This pulls the number for the current sheet’s hotel code, current row’s metric, and the column header date.

---

## 9) Troubleshooting

### Only January shows up

* Ensure your `PQ_Dates` named range includes **all months** and that the README’s version of the code **flattens all columns**:

  * It should use
    `List.Combine(Table.ToColumns(DatesTbl))`
    not `Table.Column(DatesTbl,"Column1")`.

### “We cannot convert null to Logical”

* You are likely on an older version that used `and` chains on possibly null values.
* Make sure the **date filter** in both normalizers uses:
  `Value.Is([DateHeader], type date) and List.Contains(DateListDates, [DateHeader])`.

### Allocate to Reserve Fund all zeros

* Confirm the **Col D** row label actually contains “Allocate to Reserve Fund” **or** the Col B account code equals **`46000-110`**.
* If your code is different in some hotels, add it to the detection logic:

  * `or b = "46000-115"` (for example).

### It’s pulling files from subfolders

* Confirm the folder step uses **`Folder.Contents()`** and filters `Attributes[Directory] = false`.

### Still getting errors on a specific file

* Temporarily filter `onlyAllow` to just that hotel.
* In the normalizer, temporarily return diagnostic subsets (e.g., `Table.SelectColumns(renamed, {"ColA","ColB","ColC","ColD"})`) to inspect labels, or `Table.Distinct` on `DateHeader` after `coerceD`.

---

## 10) Maintenance checklist

* New year: update sheet names in `ProcessFile` (`"2026 IS"`, `"2026 CF"`) or parameterize the year.
* Add/remove hotels: just edit the `HotelCodes` table.
* Add metrics: include them in `keepMet` lists and, if needed, extend the detection rules.
* Change date range: update `PQ_Dates`; query will adapt automatically.

---

If you want, I can tailor a “year parameter” or toggle that lets you point this same query at **any** year by just editing a single named cell (e.g., `PQ_Year`).
