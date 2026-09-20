# OPS KPI Dashboard

A JSON data layer and an HTML dashboard bolted onto the existing OPS KPI
reporting script. The KPI calculations, the SQL and the Excel output are
unchanged — the same DataFrames now feed a second output alongside Excel.

```
PostgreSQL
    |
OPS_KPI.py          existing SQL, existing Pandas frames
    |
    +--> Excel reports          (unchanged)
    |
    +--> data/ops_kpi_data.json (new, written from the same frames)
              |
        dashboard/index.html
```

---

## 1. Files

**Modified**

| File | What changed |
|---|---|
| `OPS_KPI.py` | Six edits, listed in section 7. No SQL, no formula, no filter, no exclusion list touched. |

**New**

| File | Purpose |
|---|---|
| `ops_kpi_json.py` | JSON serialisation, column/row metadata, validation, atomic write. Contains no KPI logic. |
| `validate_json_vs_excel.py` | Reconciles the JSON against the Excel workbooks from the same run. |
| `dashboard/index.html` | Dashboard shell. |
| `dashboard/styles.css` | Styling. |
| `dashboard/app.js` | Data loading, filters, sorting, chart, table. |
| `tests/test_json_layer.py` | Unit tests for the serialiser on synthetic frames. |
| `requirements.txt`, `.gitignore`, `.env.example`, `README.md` | Setup and configuration. |
| `data/` | Where `ops_kpi_data.json` lands. Git-ignored. |

Put this folder anywhere. `OPS_KPI.py`, `ops_kpi_json.py` and `dashboard/`
must stay siblings, because the script imports the module by name and the
dashboard reads `../data/ops_kpi_data.json` relative to itself.

---

## 2. Setup

```powershell
cd OPS_KPI_DASHBOARD
pip install -r requirements.txt
```

`requirements.txt` lists only what the script already needed —
pandas, SQLAlchemy, psycopg2-binary, openpyxl. Nothing new was added,
and the dashboard uses no JavaScript libraries at all.

### PostgreSQL configuration

The connection string was hardcoded, password and all. It now reads from
the environment first and falls back to the original literal so nothing
breaks on the first run:

```python
DB_URL = os.environ.get("OPS_KPI_DB_URL", "postgresql://postgres:...")
```

To take the password out of the file:

```powershell
copy .env.example .env
# edit .env with the real password (percent-encode @ as %40)
```

Load it before running, or set the variable permanently:

```powershell
[Environment]::SetEnvironmentVariable(
  "OPS_KPI_DB_URL",
  "postgresql://postgres:Miniso%4012345@localhost:5432/postgres",
  "User")
```

Once that works, delete the fallback string from `OPS_KPI.py`. Until you
do, the script prints a warning on every run. `.env` and `*.json` are
git-ignored; no credential ever reaches the JSON or the browser.

Optional variables: `OPS_KPI_BASE_FOLDER` (Excel output root),
`OPS_KPI_JSON_PATH` (JSON destination).

---

## 3. Refreshing the data

```powershell
python OPS_KPI.py
```

Unchanged behaviour for the first part of the run: connect, run the five
queries, rename the period columns to calendar dates, write five
timestamped workbooks. Then:

```
====================================================
JSON DATA LAYER
====================================================
   ABV                    212 rows  (198 stores, 13 SPoC totals, 1 MINISO total)
   UPT                    212 rows  ...
   ...
✅ JSON refreshed (1,240 KB)
   C:\OPS_KPI_DASHBOARD\data\ops_kpi_data.json
   Refreshed at 2026-09-20T09:14:33
```

**Refresh policy: atomic full refresh.** If any one report raises, no JSON
is written at all and the previous file is left exactly as it was — the
dashboard keeps showing the last complete data set rather than a mixture
of old and new. Partial refresh is deliberately not supported. The write
itself goes to a temp file in the same folder and is swapped in with
`os.replace()`, so an interrupted run cannot leave a half-written file.

Before the swap the payload is validated: every expected report present,
non-empty, exactly one MINISO row, at least one store row, no duplicate
store codes, no duplicate column labels, numerics still numeric, no NaN
or Infinity. A validation failure aborts the write and prints the reason.

---

## 4. Launching the dashboard

Browsers refuse `fetch()` on `file://`, so serve the folder:

```powershell
cd OPS_KPI_DASHBOARD
python -m http.server 8000
```

Open **http://localhost:8000/dashboard/**

Leave the server running; refreshing the page after a new
`python OPS_KPI.py` run picks up the new data (the fetch sets
`cache: "no-store"`).

---

## 5. JSON schema

```jsonc
{
  "metadata": {
    "report_name": "OPS KPI Dashboard",
    "schema_version": "1.0",
    "generated_at": "2026-09-20T09:14:33",
    "last_successful_refresh": "2026-09-20T09:14:33",
    "reporting_date": "2026-09-20",
    "status": "success",
    "refresh_mode": "atomic_full",
    "row_type_legend": { "store": "...", "spoc_total": "...", "miniso_total": "..." },
    "excluded_stores": ["IN3J", "INAI", "..."],
    "report_order": ["abv", "upt", "sales", "noh", "sales_ach"]
  },

  "reports": {
    "abv": {
      "key": "abv",
      "name": "ABV",
      "description": "Average bill value: ...",
      "unit": "currency",
      "value_format": "decimal2",
      "higher_is_better": true,
      "target_column": "target_abv",
      "achievement_basis": null,
      "excel_file": "abv_19_sep_20260920_091433.xlsx",
      "excel_sheet": "Data",
      "row_count": 212,
      "row_counts_by_type": { "store": 198, "spoc_total": 13, "miniso_total": 1 },

      "columns": [
        { "key": "store_code", "label": "store_code", "role": "dimension",
          "granularity": null, "period_index": null, "metric": null, "dtype": "string" },
        { "key": "target_abv", "label": "target_abv", "role": "target", ... },
        { "key": "day_1_abv",  "label": "19-Sep-2026", "role": "metric",
          "granularity": "day", "period_index": 1, "metric": "abv", "dtype": "number" }
      ],

      "periods": {
        "day":   [ { "key": "day_1_abv", "index": 1, "label": "19-Sep-2026" }, ... ],
        "week":  [ { "key": "week_1_abv", "index": 1, "label": "07-Sep to 13-Sep-2026" }, ... ],
        "month": [ { "key": "month_1_abv", "index": 1, "label": "Aug-2026" }, ... ]
      },

      "data": [
        { "store_code": "INAA", "store_name": "...", "spoc_name": "...",
          "target_abv": 470.0, "day_1_abv": 501.25, "day_2_abv": null, "...": null,
          "_row_type": "store", "_spoc": "..." }
      ]
    }
  }
}
```

Points worth knowing:

- **Records are keyed by the original SQL alias** (`day_1_abv`,
  `w3_neg_sku`). `columns[].label` carries the calendar label — the exact
  string in the Excel header, produced by the script's own
  `build_date_rename_map()`. Stable machine key and human label, both.
- **Row hierarchy** is tagged, not guessed. `_row_type` is derived from
  the value the SQL already emits in `store_code`: `MINISO` → grand
  total, `<spoc> - TOTAL` → SPoC subtotal, anything else → a store.
  `_spoc` is the SPoC a row belongs to, recovered from that same label
  for subtotal rows so they can be filtered; the original `spoc_name`
  stays NULL on those rows exactly as the query returns it.
- **Nulls stay null.** A store that did not trade gets `null`, never `0`.
  Stores with no activity at all are still present, one row each.
- **Numbers stay numbers.** Decimals from psycopg2 become floats, not
  strings; validation fails the run if any metric arrives as text.
- Reports are keyed by the same names as the folders: `abv`, `upt`,
  `sales`, `noh`, `sales_ach`. Adding a sixth report to the `reports`
  list in `OPS_KPI.py` — with a `key` and the display fields — puts it in
  the JSON and the dashboard with no other change.

### Period semantics

Taken verbatim from the existing SQL and renamer, not reinterpreted:

| Alias | Means | Label style |
|---|---|---|
| `day_1` / `d1` | yesterday, the last completed day | `19-Sep-2026` |
| `day_7` / `d7` | seven days back |  |
| `week_1` / `w1` | the last completed Monday–Sunday week, never the current part-week | `07-Sep to 13-Sep-2026` |
| `week_4` / `w4` | four completed weeks back |  |
| `month_1` / `m1` | last calendar month, never the current part-month | `Aug-2026` |
| `month_3` / `m3` | three calendar months back |  |

Labels come from the machine's local date. If the script host and the
Postgres host are in different time zones, labels can slip by a day
relative to `CURRENT_DATE` — this was already true of the Excel output
and is unchanged.

---

## 6. Dashboard behaviour

- **Left rail** lists the five reports with the MINISO figure for the most
  recent completed day, so switching reports is also a scoreboard.
- **MINISO overall** cards read the grand-total row the report produces.
  Nothing is re-added from store rows, and the MINISO row is excluded
  from the table so it can't be counted twice.
- **Period control** is built from `periods` in the JSON: day/week/month,
  then the actual dates. Selecting one highlights that column in the
  table and drives the chart and the status flags.
- **Filters**: store search (code or name), SPoC, row type (stores only /
  SPoC totals only / both), below-target only. Sorting on any column,
  nulls always sinking. Pagination at 25/50/100/all.
- **Chart**: highest and lowest five stores for the selected period, with
  a zero baseline. Stores with no value for that period are excluded from
  the chart and the count is stated next to it.
- **Status colours** use only what the source data defines:

  | Report | Basis | Rule |
  |---|---|---|
  | ABV | the store's own `target_abv` | at or above target = green, below = red |
  | UPT | the store's own `target_upt` | same |
  | Sales achievement | the metric is already a % of target | 100% or more = green, below = red |
  | Sales | none in the data | no colour |
  | Negative stock | none in the data | no colour |

  There is no amber band, because no amber threshold exists anywhere in
  the source. Colour is never the only signal: each flagged value carries
  a dot and a tooltip, and the target column is on screen. Subtotal and
  grand-total rows carry no per-store target, so they are not flagged
  against one.

- Column headers show the calendar labels exactly as Excel does. Three
  dimension headers are shown in plain English (`store_code` → "Store
  code", `store_name` → "Store", `spoc_name` → "SPoC") with the raw name
  in the tooltip; that is display only, the JSON label is untouched.
- Loading, error and empty states are handled. A report that arrives
  malformed is skipped rather than blanking the whole page.

---

## 7. Exactly what changed in `OPS_KPI.py`

Six edits. Diff against your original shows no other line removed.

1. `import json` and `import ops_kpi_json as jsonlayer` added.
2. `DB_URL` reads `OPS_KPI_DB_URL` from the environment, falling back to
   the original literal.
3. `base_folder` accepts an `OPS_KPI_BASE_FOLDER` override;
   `JSON_OUTPUT_PATH` added below it.
4. Each of the five report dicts gained display-only fields: `key`,
   `display_name`, `description`, `unit`, `value_format`,
   `higher_is_better`, `target_column`, `achievement_basis`. No query,
   folder or filename changed.
5. In the loop: `df_raw = pd.read_sql_query(...)` and
   `df = df_raw.rename(columns=rename_map)`. Excel still receives the
   renamed frame, byte for byte the same as before. On success the raw
   frame is handed to the JSON layer; on failure the report key is
   recorded.
6. After `engine.dispose()`: the atomic JSON refresh block and a run
   summary.

Untouched: every SQL query, `EXCLUDED_STORES`, `build_date_rename_map`,
`_shift_month`, the GROUPING SETS hierarchy, store-code mapping, SPoC
mapping, MINISO totals, null handling, targets, the Excel writer, the
timestamped filenames.

---

## 8. Validation

### Reconciliation against Excel — you must run this

```powershell
python validate_json_vs_excel.py --base "J:\...\Supporting_Folder_For_Complete_dashboard\ops kpi"
```

It opens each workbook the JSON names — the exact file from that run, not
the newest one it can find — and checks row counts, column count, order
and headers, store order and identity, every cell including nulls,
duplicate store codes, numeric types, and the store / SPoC-total / MINISO
counts. It prints the report, column and store for any mismatch, and
exits non-zero if anything fails.

**I have not run it.** This environment has no PostgreSQL, no network and
none of your data, so no Excel file and no JSON exist here to compare. I
am not going to tell you the outputs match when nothing has been
compared. Run it after your next `python OPS_KPI.py` and send me the
output if anything fails.

### What was tested here

`python tests/test_json_layer.py` — 30 checks, all passing, on synthetic
frames shaped like the real reports. They cover: NULL, NaN and Decimal
handling; numpy unwrapping; all-null stores surviving; row-type
classification; SPoC recovery on subtotal rows; JSON labels matching the
Excel headers produced by the same renamer; period ordering; day_1 being
yesterday, week_1 ending on the last completed Sunday, month_1 not being
the current month; every validation rule rejecting the payload it should;
and the atomic write leaving the previous file intact after a failure.

The dashboard was rendered in headless Chromium at 1440px and 390px
against a synthetic payload, with no console errors, and checked for
report switching, period switching, filtering, the empty state and
pagination. That exercises the code paths, not your numbers.

---

## 9. Assumptions

1. `folder` and report `key` are the same word for all five reports, so
   `validate_json_vs_excel.py` looks in `<base>/<key>/` first (it falls
   back to a recursive search).
2. A `store_code` of `MINISO`, or one ending in ` - TOTAL`, is an
   aggregate row. That is how the SQL labels them; if a real store were
   ever named that way it would be misclassified.
3. Excel and JSON are compared row-position by row-position, which holds
   because both come from one frame in the query's `ORDER BY` sequence.
4. Currency is Indian rupees and is formatted with the Indian digit
   grouping; sales figures over one lakh are abbreviated in the rail only,
   never in the table.
5. `higher_is_better: false` on negative stock reflects the metric's
   meaning. It has no target in the data, so it changes nothing on screen
   today.
6. The dashboard treats data older than 36 hours as stale and says so in
   the header. Display only.

## 10. Known limitations

1. **The Excel reconciliation has not been run.** See section 8.
2. The whole JSON loads into the browser at once. Five reports at roughly
   200 rows and 18 columns is around 1–2 MB, which is fine locally. Past
   a few thousand stores it would need paging or one file per report.
3. Only the five current reports have display metadata. A new report
   without a `key` will raise a `KeyError` in the loop — add the display
   fields alongside the query.
4. Negative stock and Sales have no thresholds in the source, so they are
   shown without status colour. Give me the thresholds and I will add
   them; I am not inventing any.
5. The MINISO row is shown in the summary cards only, not as a table row.
6. The chart shows the top and bottom five for one period. There is no
   trend chart, because the reports hold fixed rolling periods rather
   than a history to plot.
7. No auto-refresh in the browser — reload the page after a run.
8. The hardcoded password is still in `OPS_KPI.py` as a fallback. It stays
   a live credential in a file until you set `OPS_KPI_DB_URL` and delete
   it. If that file has ever been committed anywhere, rotate the password.
