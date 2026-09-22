import pandas as pd
from sqlalchemy import create_engine
import os
import json
import re
import traceback
from datetime import datetime, date, timedelta

# JSON data layer for the HTML dashboard. Serialisation only -- it
# contains no KPI logic and never touches the SQL or the DataFrames.
import ops_kpi_json as jsonlayer

# ==========================================
# PostgreSQL Connection
# ==========================================

# Credentials come from the environment when available so they are not
# stored in this file or in anything committed to git. The previous
# hardcoded URL is kept ONLY as a fallback so the existing workflow
# keeps running unchanged; set OPS_KPI_DB_URL (see .env.example) and
# then delete the fallback string below.
DB_URL = os.environ.get(
    "OPS_KPI_DB_URL",
    "postgresql://postgres:Miniso%4012345@localhost:5432/postgres",
)

if "OPS_KPI_DB_URL" not in os.environ:
    print("⚠  Using the hardcoded database URL. Set OPS_KPI_DB_URL to remove the password from this file.")

try:
    engine = create_engine(DB_URL)

    with engine.connect() as conn:
        print("✅ Connected to PostgreSQL Successfully")

except Exception as e:
    print("❌ Database Connection Error")
    traceback.print_exc()
    exit()

# ==========================================
# Base Folder
# ==========================================

base_folder = r"C:\Users\thinkpad\Downloads\OPS_KPI_HTML_DASHBOARD\reports"

if not os.path.exists(base_folder):
    print(f"❌ Base folder not found:\n{base_folder}")
    exit()

# ==========================================
# JSON output (dashboard data layer)
# ==========================================
# Written next to this script by default: OPS_KPI_DASHBOARD/data/ops_kpi_data.json
# Override with OPS_KPI_JSON_PATH if the dashboard lives elsewhere.

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

JSON_OUTPUT_PATH = os.environ.get(
    "OPS_KPI_JSON_PATH",
    os.path.join(SCRIPT_DIR, "data", "ops_kpi_data.json"),
)

# ==========================================
# SINGLE SOURCE OF TRUTH — EXCLUDED STORES
# ==========================================
# This is the ONLY place store exclusions are defined.
# Every report query below pulls from this same list, so a store
# excluded here is excluded everywhere — no more per-query drift.
# (Previously: abv/upt/sales had NO exclusions, sales_ach had a
#  12-store list, noh had a different 3-store list. That mismatch,
#  not store_code_mapping, is what let INFD/INFK keep appearing —
#  store_code_mapping only renames codes, it never filters them.)

EXCLUDED_STORES = [
    "IN3J",
    "INAI",
    "INAQ",
    "INB2",
    "INES",
    "INFD",
    "INFK",
    "INL7",
    "INP1",
    "MDF",
    "Puma_GGN_Sec_64",
    "Puma_Noida_Sec_49",
    "Puma_NSP",
    "U K LIFESTYLE-OMAXE WORLD STREET-FARIDABAD",
    "U K LIFESTYLE-SALT LAKE 5-KOLKATA",
]

# Pre-formatted for direct injection into SQL NOT IN (...) clauses
EXCLUDED_STORES_SQL = ",".join(f"'{s}'" for s in EXCLUDED_STORES)

# ==========================================
# COLUMN HEADER -> ACTUAL DATE RENAMING
# ==========================================
# The SQL aliases (day_1_abv, week_1_abv, month_1_abv, d1_neg_sku, ...)
# stay fixed in the query itself — that's what keeps the query
# reliable and easy to maintain. But the report your head reads
# should show the real calendar date/period as the column header
# instead of "day_1". This renames the DataFrame's columns AFTER
# the SQL runs, right before the Excel file is written, so the
# underlying SQL logic and column order never change — only the
# label a human sees.
#
# NOTE: this uses the machine's local date (date.today()), which
# must match the Postgres server's CURRENT_DATE for the labels to
# line up with the data. Since this script queries the DB and
# writes the file in the same run, that's true unless the script
# server and DB server are in different timezones — worth a
# one-time check if labels ever look off by a day.

def _shift_month(first_of_month, months_back):
    total = first_of_month.month - 1 - months_back
    year = first_of_month.year + total // 12
    month = total % 12 + 1
    return date(year, month, 1)

def build_date_rename_map(columns):
    today = date.today()
    week_start = today - timedelta(days=today.weekday())   # Monday of current (incomplete) week
    month_start = date(today.year, today.month, 1)

    day_pat   = re.compile(r"^(?:day_|d)(\d)_(.+)$")
    week_pat  = re.compile(r"^(?:week_|w)(\d)_(.+)$")
    month_pat = re.compile(r"^(?:month_|m)(\d)_(.+)$")

    rename = {}
    for col in columns:
        m = day_pat.match(col)
        if m:
            n = int(m.group(1))
            d = today - timedelta(days=n)
            rename[col] = d.strftime("%d-%b-%Y")
            continue

        m = week_pat.match(col)
        if m:
            n = int(m.group(1))
            w_end = week_start - timedelta(days=(n - 1) * 7 + 1)
            w_start = w_end - timedelta(days=6)
            rename[col] = f"{w_start.strftime('%d-%b')} to {w_end.strftime('%d-%b-%Y')}"
            continue

        m = month_pat.match(col)
        if m:
            n = int(m.group(1))
            mth = _shift_month(month_start, n)
            rename[col] = mth.strftime("%b-%Y")
            continue

    return rename

# ==========================================
# Reports Configuration
# ==========================================

reports = [
    {
        "key": "abv",
        "display_name": "ABV",
        "description": "Average bill value: sales value divided by bills, weighted at every level.",
        "unit": "currency",
        "value_format": "decimal2",
        "higher_is_better": True,
        "target_column": "target_abv",
        "achievement_basis": None,
        "folder": "abv",
        "file": "abv_22_sep.xlsx",
        "query": f"""
        WITH store_spoc AS (
    -- Authoritative store universe: every store with a SPOC/target
    -- assignment shows up here regardless of whether it billed
    -- anything in the lookback window. This is what previously made
    -- zero-recent-sales stores (e.g. INFF, INFN) vanish from the
    -- report entirely — they simply never had a row in b2c_sale for
    -- the queried dates, so the old base CTE never produced a group
    -- for them.
    SELECT DISTINCT
        COALESCE(scm.store_code, ot.store_code) AS store_code,
        ot.spoc_name
    FROM ops_targets ot
    LEFT JOIN store_code_mapping scm
        ON ot.store_code = scm.old_store_code
    WHERE COALESCE(scm.store_code, ot.store_code) NOT IN ({EXCLUDED_STORES_SQL})
),

sales_daily AS (
    SELECT
        COALESCE(scm.store_code, bs.store_code) AS store_code,
        safe_to_date(bs.date::text) AS dt,
        SUM(bs.amt_with_gst) AS sales_value,
        COUNT(DISTINCT bs.invoice_no) AS bills
    FROM b2c_sale bs
    LEFT JOIN store_code_mapping scm
        ON bs.store_code = scm.old_store_code
    WHERE COALESCE(scm.store_code, bs.store_code) NOT IN ({EXCLUDED_STORES_SQL})
    GROUP BY
        COALESCE(scm.store_code, bs.store_code),
        safe_to_date(bs.date::text)
),

base AS (
    -- every store in store_spoc appears at least once, even with
    -- NULL dt/sales_value/bills if it has zero transactions
    SELECT
        ss.store_code,
        ss.spoc_name,
        sd.dt,
        sd.sales_value,
        sd.bills
    FROM store_spoc ss
    LEFT JOIN sales_daily sd
        ON ss.store_code = sd.store_code

    UNION ALL

    -- safety net: a store that billed but is missing from
    -- ops_targets still shows up (with a blank SPOC) rather than
    -- being silently dropped
    SELECT
        sd.store_code,
        NULL AS spoc_name,
        sd.dt,
        sd.sales_value,
        sd.bills
    FROM sales_daily sd
    WHERE sd.store_code NOT IN (SELECT store_code FROM store_spoc)
),

abv_report AS (

SELECT
    store_code,
    spoc_name,
    GROUPING(spoc_name)  AS grp_spoc,   -- 1 only on the single grand-total row
    GROUPING(store_code) AS grp_store,  -- 1 on spoc-subtotal AND grand-total rows

    /* -------------------- LAST 7 COMPLETED DAYS -------------------- */

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '1 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '1 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS day_1_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '2 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '2 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS day_2_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '3 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '3 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS day_3_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '4 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '4 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS day_4_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '5 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '5 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS day_5_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '6 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '6 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS day_6_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '7 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '7 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS day_7_abv,

    /* -------------------- LAST 4 COMPLETED CALENDAR WEEKS -------------------- */

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                 AND dt <  date_trunc('week', CURRENT_DATE)::date
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                     AND dt <  date_trunc('week', CURRENT_DATE)::date
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS week_1_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                 AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                     AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS week_2_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                 AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                     AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS week_3_abv,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '28 day'
                 AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '28 day'
                     AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS week_4_abv,
	    /* -------------------- LAST 3 COMPLETED MONTHS -------------------- */

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '1 month'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '1 month'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS month_1_abv,

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '2 month'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '2 month'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS month_2_abv,

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '3 month'
                THEN sales_value
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '3 month'
                    THEN bills
                END
            ),
            0
        ),
        2
    ) AS month_3_abv

FROM base

GROUP BY GROUPING SETS
(
    (spoc_name, store_code),   -- store-level rows
    (spoc_name),               -- SPOC subtotal rows (weighted sum/sum, not avg-of-avg)
    ()                         -- MINISO grand total
)

),

store_master AS (

    SELECT
        store_code,
        MAX(store_name) AS store_name
    FROM b2c_sale
    GROUP BY store_code

),

target_master AS (

    SELECT
        store_code,
        MAX(abv) AS target_abv
    FROM ops_targets
    GROUP BY store_code

)

SELECT

    CASE
        WHEN a.grp_spoc  = 1 THEN 'MINISO'
        WHEN a.grp_store = 1 THEN COALESCE(a.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE a.store_code
    END AS store_code,

    CASE
        WHEN a.grp_spoc  = 1 THEN 'MINISO'
        WHEN a.grp_store = 1 THEN COALESCE(a.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE sm.store_name
    END AS store_name,

    CASE
        WHEN a.grp_store = 1 THEN NULL   -- covers both spoc-total & grand-total rows
        ELSE a.spoc_name
    END AS spoc_name,

    CASE
        WHEN a.grp_store = 1 THEN NULL
        ELSE tm.target_abv
    END AS target_abv,

    /* actual calendar date each rolling-day column refers to */
    a.day_1_abv,
    a.day_2_abv,
    a.day_3_abv,
    a.day_4_abv,
    a.day_5_abv,
    a.day_6_abv,
    a.day_7_abv,

    /* actual Mon-Sun calendar-week range each week column refers to */
    a.week_1_abv,
    a.week_2_abv,
    a.week_3_abv,
    a.week_4_abv,

    /* actual calendar month each month column refers to */
    a.month_1_abv,
    a.month_2_abv,
    a.month_3_abv

FROM abv_report a

LEFT JOIN store_master sm
    ON a.store_code = sm.store_code

LEFT JOIN target_master tm
    ON a.store_code = tm.store_code

ORDER BY
    a.grp_spoc,                                    -- MINISO grand total sorts absolutely last
    COALESCE(a.spoc_name, 'zzz_no_spoc_assigned'),  -- SPOCs alphabetically
    a.grp_store,                                    -- stores (0) before that SPOC's total row (1)
    a.store_code;


        """
    },

     {
        "key": "upt",
        "display_name": "UPT",
        "description": "Units per transaction: quantity divided by bills, weighted at every level.",
        "unit": "units",
        "value_format": "decimal2",
        "higher_is_better": True,
        "target_column": "target_upt",
        "achievement_basis": None,
        "folder": "upt",
        "file": "upt_22_sep.xlsx",
        "query": f"""
    WITH store_spoc AS (
    SELECT DISTINCT
        COALESCE(scm.store_code, ot.store_code) AS store_code,
        ot.spoc_name
    FROM ops_targets ot
    LEFT JOIN store_code_mapping scm
        ON ot.store_code = scm.old_store_code
    WHERE COALESCE(scm.store_code, ot.store_code) NOT IN ({EXCLUDED_STORES_SQL})
),

sales_daily AS (
    SELECT
        COALESCE(scm.store_code, bs.store_code) AS store_code,
        safe_to_date(bs.date::text) AS dt,
        SUM(bs.qty) AS total_qty,
        COUNT(DISTINCT bs.invoice_no) AS bills
    FROM b2c_sale bs
    LEFT JOIN store_code_mapping scm
        ON bs.store_code = scm.old_store_code
    WHERE COALESCE(scm.store_code, bs.store_code) NOT IN ({EXCLUDED_STORES_SQL})
    GROUP BY
        COALESCE(scm.store_code, bs.store_code),
        safe_to_date(bs.date::text)
),

base AS (
    SELECT
        ss.store_code,
        ss.spoc_name,
        sd.dt,
        sd.total_qty,
        sd.bills
    FROM store_spoc ss
    LEFT JOIN sales_daily sd
        ON ss.store_code = sd.store_code

    UNION ALL

    SELECT
        sd.store_code,
        NULL AS spoc_name,
        sd.dt,
        sd.total_qty,
        sd.bills
    FROM sales_daily sd
    WHERE sd.store_code NOT IN (SELECT store_code FROM store_spoc)
),

upt_report AS (

SELECT
    store_code,
    spoc_name,
    GROUPING(spoc_name)  AS grp_spoc,
    GROUPING(store_code) AS grp_store,

    /* -------------------- LAST 7 COMPLETED DAYS -------------------- */

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '1 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '1 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS day_1_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '2 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '2 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS day_2_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '3 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '3 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS day_3_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '4 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '4 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS day_4_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '5 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '5 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS day_5_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '6 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
               CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '6 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS day_6_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt = CURRENT_DATE - INTERVAL '7 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt = CURRENT_DATE - INTERVAL '7 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS day_7_upt,

    /* -------------------- LAST 4 COMPLETED CALENDAR WEEKS -------------------- */

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                 AND dt <  date_trunc('week', CURRENT_DATE)::date
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                     AND dt <  date_trunc('week', CURRENT_DATE)::date
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS week_1_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                 AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                     AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS week_2_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                 AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                     AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS week_3_upt,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '28 day'
                 AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '28 day'
                     AND dt <  date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS week_4_upt,
	    /* -------------------- LAST 3 COMPLETED MONTHS -------------------- */

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '1 month'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '1 month'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS month_1_upt,

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '2 month'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '2 month'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS month_2_upt,

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '3 month'
                THEN total_qty
            END
        )::numeric
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '3 month'
                    THEN bills
                END
            ),
            0
        ),
        1
    ) AS month_3_upt

FROM base

GROUP BY GROUPING SETS
(
    (spoc_name, store_code),
    (spoc_name),
    ()
)

),

store_master AS (

    SELECT
        store_code,
        MAX(store_name) AS store_name
    FROM b2c_sale
    GROUP BY store_code

),

target_master AS (

    SELECT
        store_code,
        MAX(upt) AS target_upt
    FROM ops_targets
    GROUP BY store_code

)

SELECT

    CASE
        WHEN a.grp_spoc  = 1 THEN 'MINISO'
        WHEN a.grp_store = 1 THEN COALESCE(a.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE a.store_code
    END AS store_code,

    CASE
        WHEN a.grp_spoc  = 1 THEN 'MINISO'
        WHEN a.grp_store = 1 THEN COALESCE(a.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE sm.store_name
    END AS store_name,

    CASE
        WHEN a.grp_store = 1 THEN NULL
        ELSE a.spoc_name
    END AS spoc_name,

    CASE
        WHEN a.grp_store = 1 THEN NULL
        ELSE tm.target_upt
    END AS target_upt,

    a.day_1_upt,
    a.day_2_upt,
    a.day_3_upt,
    a.day_4_upt,
    a.day_5_upt,
    a.day_6_upt,
    a.day_7_upt,

    a.week_1_upt,
    a.week_2_upt,
    a.week_3_upt,
    a.week_4_upt,

    a.month_1_upt,
    a.month_2_upt,
    a.month_3_upt

FROM upt_report a

LEFT JOIN store_master sm
    ON a.store_code = sm.store_code

LEFT JOIN target_master tm
    ON a.store_code = tm.store_code

ORDER BY
    a.grp_spoc,
    COALESCE(a.spoc_name, 'zzz_no_spoc_assigned'),
    a.grp_store,
    a.store_code;

        """
     },
     {
        "key": "sales",
        "display_name": "Sales",
        "description": "Sales value including GST.",
        "unit": "currency",
        "value_format": "currency0",
        "higher_is_better": True,
        "target_column": None,
        "achievement_basis": None,
        "folder": "sales",
        "file": "sales_22_sep.xlsx",
        "query": f"""
        WITH store_spoc AS (
    SELECT DISTINCT
        COALESCE(scm.store_code, ot.store_code) AS store_code,
        ot.spoc_name
    FROM ops_targets ot
    LEFT JOIN store_code_mapping scm
        ON ot.store_code = scm.old_store_code
    WHERE COALESCE(scm.store_code, ot.store_code) NOT IN ({EXCLUDED_STORES_SQL})
),

sales_daily AS (
    SELECT
        COALESCE(scm.store_code, bs.store_code) AS store_code,
        safe_to_date(bs.date::text) AS dt,
        SUM(bs.amt_with_gst) AS sales
    FROM b2c_sale bs
    LEFT JOIN store_code_mapping scm
        ON bs.store_code = scm.old_store_code
    WHERE COALESCE(scm.store_code, bs.store_code) NOT IN ({EXCLUDED_STORES_SQL})
    GROUP BY
        COALESCE(scm.store_code, bs.store_code),
        safe_to_date(bs.date::text)
),

base AS (
    SELECT
        ss.store_code,
        ss.spoc_name,
        sd.dt,
        sd.sales
    FROM store_spoc ss
    LEFT JOIN sales_daily sd
        ON ss.store_code = sd.store_code

    UNION ALL

    SELECT
        sd.store_code,
        NULL AS spoc_name,
        sd.dt,
        sd.sales
    FROM sales_daily sd
    WHERE sd.store_code NOT IN (SELECT store_code FROM store_spoc)
),

sales_report AS (

SELECT
    store_code,
    spoc_name,
    GROUPING(spoc_name)  AS grp_spoc,
    GROUPING(store_code) AS grp_store,

        /* =========================
       MONTH-TO-DATE (MTD) SALES
       From the 1st of the current month through yesterday
       (today excluded — still an incomplete day)
    ========================= */

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('month', CURRENT_DATE)::date
                 AND dt < CURRENT_DATE
                THEN sales
            END
        ),
        2
    ) AS mtd_sales,

    /* =========================
       LAST 7 COMPLETED DAYS
    ========================= */

    ROUND(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '1 day' THEN sales END), 2) AS day_1_sales,
    ROUND(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '2 day' THEN sales END), 2) AS day_2_sales,
    ROUND(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '3 day' THEN sales END), 2) AS day_3_sales,
    ROUND(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '4 day' THEN sales END), 2) AS day_4_sales,
    ROUND(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '5 day' THEN sales END), 2) AS day_5_sales,
    ROUND(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '6 day' THEN sales END), 2) AS day_6_sales,
    ROUND(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '7 day' THEN sales END), 2) AS day_7_sales,

    /* =========================
       LAST 4 COMPLETED CALENDAR WEEKS
    ========================= */

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                 AND dt < date_trunc('week', CURRENT_DATE)::date
                THEN sales
            END
        ),
        2
    ) AS week_1_sales,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                 AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                THEN sales
            END
        ),
        2
    ) AS week_2_sales,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                 AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                THEN sales
            END
        ),
        2
    ) AS week_3_sales,

    ROUND(
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '28 day'
                 AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                THEN sales
            END
        ),
        2
    ) AS week_4_sales,

    /* =========================
       LAST 3 COMPLETED MONTHS
    ========================= */

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '1 month'
                THEN sales
            END
        ),
        2
    ) AS month_1_sales,

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '2 month'
                THEN sales
            END
        ),
        2
    ) AS month_2_sales,

    ROUND(
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '3 month'
                THEN sales
            END
        ),
        2
    ) AS month_3_sales

FROM base
GROUP BY GROUPING SETS
(
    (spoc_name, store_code),
    (spoc_name),
    ()
)

),

store_master AS (

    SELECT
        store_code,
        MAX(store_name) AS store_name
    FROM b2c_sale
    GROUP BY store_code

)

SELECT

    CASE
        WHEN s.grp_spoc  = 1 THEN 'MINISO'
        WHEN s.grp_store = 1 THEN COALESCE(s.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE s.store_code
    END AS store_code,

    CASE
        WHEN s.grp_spoc  = 1 THEN 'MINISO'
        WHEN s.grp_store = 1 THEN COALESCE(s.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE sm.store_name
    END AS store_name,

    CASE
        WHEN s.grp_store = 1 THEN NULL
        ELSE s.spoc_name
    END AS spoc_name,

    s.mtd_sales,

    s.day_1_sales,
    s.day_2_sales,
    s.day_3_sales,
    s.day_4_sales,
    s.day_5_sales,
    s.day_6_sales,
    s.day_7_sales,

    s.week_1_sales,
    s.week_2_sales,
    s.week_3_sales,
    s.week_4_sales,

    s.month_1_sales,
    s.month_2_sales,
    s.month_3_sales

FROM sales_report s

LEFT JOIN store_master sm
    ON s.store_code = sm.store_code

ORDER BY
    s.grp_spoc,
    COALESCE(s.spoc_name, 'zzz_no_spoc_assigned'),
    s.grp_store,
    s.store_code;

        """
     },
     {
        "key": "noh",
        "display_name": "Negative stock",
        "description": "Average count of SKUs with negative available stock per day in the period.",
        "unit": "skus",
        "value_format": "decimal1",
        "higher_is_better": False,
        "target_column": None,
        "achievement_basis": None,
        "folder": "noh",
        "file": "noh_22_sep.xlsx",
        "query": f"""
WITH store_spoc AS (
    SELECT DISTINCT
        COALESCE(scm.store_code, ot.store_code) AS store_code,
        ot.spoc_name
    FROM ops_targets ot
    LEFT JOIN store_code_mapping scm
        ON ot.store_code = scm.old_store_code
    WHERE COALESCE(scm.store_code, ot.store_code) NOT IN ({EXCLUDED_STORES_SQL})
),

inv_daily AS (

    SELECT

        COALESCE(
            NULLIF(TRIM(scm.store_code::text), ''),
            TRIM(i.store_code::text)
        ) AS store_code,

        safe_to_date(LEFT(i.date::text,10)) AS dt,

        COUNT(
            DISTINCT CASE
                WHEN COALESCE(i.available_stock,0) < 0
                THEN i.product_code
            END
        ) AS negative_sku_count

    FROM inv_90_days i

    LEFT JOIN store_code_mapping scm
        ON TRIM(i.store_code::text) =
           TRIM(scm.old_store_code::text)

    WHERE
        safe_to_date(LEFT(i.date::text,10))
            >= CURRENT_DATE - INTERVAL '120 days'

        AND COALESCE(
                NULLIF(TRIM(scm.store_code::text), ''),
                TRIM(i.store_code::text)
            ) NOT IN ({EXCLUDED_STORES_SQL})

    GROUP BY

        COALESCE(
            NULLIF(TRIM(scm.store_code::text), ''),
            TRIM(i.store_code::text)
        ),

        safe_to_date(LEFT(i.date::text,10))

),

base AS (
    SELECT
        ss.store_code,
        ss.spoc_name,
        idl.dt,
        idl.negative_sku_count
    FROM store_spoc ss
    LEFT JOIN inv_daily idl
        ON ss.store_code = idl.store_code

    UNION ALL

    SELECT
        idl.store_code,
        NULL AS spoc_name,
        idl.dt,
        idl.negative_sku_count
    FROM inv_daily idl
    WHERE idl.store_code NOT IN (SELECT store_code FROM store_spoc)
),

negative_sku_report AS (

SELECT

    store_code,
    spoc_name,
    GROUPING(spoc_name)  AS grp_spoc,
    GROUPING(store_code) AS grp_store,

    /* ================= DAILY ================= */

    SUM(CASE WHEN dt=CURRENT_DATE-INTERVAL '1 day' THEN negative_sku_count ELSE 0 END) AS d1_neg_sku,
    SUM(CASE WHEN dt=CURRENT_DATE-INTERVAL '2 day' THEN negative_sku_count ELSE 0 END) AS d2_neg_sku,
    SUM(CASE WHEN dt=CURRENT_DATE-INTERVAL '3 day' THEN negative_sku_count ELSE 0 END) AS d3_neg_sku,
    SUM(CASE WHEN dt=CURRENT_DATE-INTERVAL '4 day' THEN negative_sku_count ELSE 0 END) AS d4_neg_sku,
    SUM(CASE WHEN dt=CURRENT_DATE-INTERVAL '5 day' THEN negative_sku_count ELSE 0 END) AS d5_neg_sku,
    SUM(CASE WHEN dt=CURRENT_DATE-INTERVAL '6 day' THEN negative_sku_count ELSE 0 END) AS d6_neg_sku,
    SUM(CASE WHEN dt=CURRENT_DATE-INTERVAL '7 day' THEN negative_sku_count ELSE 0 END) AS d7_neg_sku,

    /* ================= WEEKLY ================= */

    ROUND(
        AVG(
            CASE
                WHEN dt>=date_trunc('week',CURRENT_DATE)::date-INTERVAL '7 day'
                 AND dt<date_trunc('week',CURRENT_DATE)::date
                THEN negative_sku_count
            END
        ),
        1
    ) AS w1_neg_sku,

    ROUND(
        AVG(
            CASE
                WHEN dt>=date_trunc('week',CURRENT_DATE)::date-INTERVAL '14 day'
                 AND dt<date_trunc('week',CURRENT_DATE)::date-INTERVAL '7 day'
                THEN negative_sku_count
            END
        ),
        1
    ) AS w2_neg_sku,

    ROUND(
        AVG(
            CASE
                WHEN dt>=date_trunc('week',CURRENT_DATE)::date-INTERVAL '21 day'
                 AND dt<date_trunc('week',CURRENT_DATE)::date-INTERVAL '14 day'
                THEN negative_sku_count
            END
        ),
        1
    ) AS w3_neg_sku,

    ROUND(
        AVG(
            CASE
                WHEN dt>=date_trunc('week',CURRENT_DATE)::date-INTERVAL '28 day'
                 AND dt<date_trunc('week',CURRENT_DATE)::date-INTERVAL '21 day'
                THEN negative_sku_count
            END
        ),
        1
    ) AS w4_neg_sku,
	    /* ================= MONTHLY ================= */

    ROUND(
        AVG(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE - INTERVAL '1 month')
                THEN negative_sku_count
            END
        ),
        1
    ) AS m1_neg_sku,

    ROUND(
        AVG(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE - INTERVAL '2 month')
                THEN negative_sku_count
            END
        ),
        1
    ) AS m2_neg_sku,

    ROUND(
        AVG(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE - INTERVAL '3 month')
                THEN negative_sku_count
            END
        ),
        1
    ) AS m3_neg_sku

FROM base

GROUP BY GROUPING SETS
(
    (spoc_name, store_code),
    (spoc_name),
    ()
)

),

store_master AS (

    SELECT
        store_code,
        MAX(store_name) AS store_name
    FROM inv_90_days
    WHERE store_code NOT IN ({EXCLUDED_STORES_SQL})
    GROUP BY store_code

)

SELECT

    CASE
        WHEN r.grp_spoc  = 1 THEN 'MINISO'
        WHEN r.grp_store = 1 THEN COALESCE(r.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE r.store_code
    END AS store_code,

    CASE
        WHEN r.grp_spoc  = 1 THEN 'MINISO'
        WHEN r.grp_store = 1 THEN COALESCE(r.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE sm.store_name
    END AS store_name,

    CASE
        WHEN r.grp_store = 1 THEN NULL
        ELSE r.spoc_name
    END AS spoc_name,

    r.d1_neg_sku,
    r.d2_neg_sku,
    r.d3_neg_sku,
    r.d4_neg_sku,
    r.d5_neg_sku,
    r.d6_neg_sku,
    r.d7_neg_sku,

    r.w1_neg_sku,
    r.w2_neg_sku,
    r.w3_neg_sku,
    r.w4_neg_sku,

    r.m1_neg_sku,
    r.m2_neg_sku,
    r.m3_neg_sku

FROM negative_sku_report r

LEFT JOIN store_master sm
    ON r.store_code = sm.store_code

ORDER BY
    r.grp_spoc,
    COALESCE(r.spoc_name, 'zzz_no_spoc_assigned'),
    r.grp_store,
    r.store_code;
        """
     },
      {
        "key": "sales_ach",
        "display_name": "Sales achievement",
        "description": "Sales as a percentage of sales target.",
        "unit": "percent",
        "value_format": "percent2",
        "higher_is_better": True,
        "target_column": None,
        "achievement_basis": 100,
        "folder": "sales_ach",
        "file": "sales_ach_22_sep.xlsx",
        "query": f"""
         WITH store_spoc AS (
    SELECT DISTINCT
        COALESCE(scm.store_code, ot.store_code) AS store_code,
        ot.spoc_name
    FROM ops_targets ot
    LEFT JOIN store_code_mapping scm
        ON ot.store_code = scm.old_store_code
    WHERE COALESCE(scm.store_code, ot.store_code) NOT IN ({EXCLUDED_STORES_SQL})
),

base AS (

    SELECT
        COALESCE(scm.store_code, bs.store_code) AS store_code,
        safe_to_date(bs.date::text) AS dt,
        SUM(bs.amt_with_gst) AS sales

    FROM b2c_sale bs

    LEFT JOIN store_code_mapping scm
        ON bs.store_code = scm.old_store_code

    WHERE COALESCE(scm.store_code, bs.store_code) NOT IN ({EXCLUDED_STORES_SQL})

    GROUP BY
        COALESCE(scm.store_code, bs.store_code),
        safe_to_date(bs.date::text)

),

target_base AS (

    SELECT
        COALESCE(scm.store_code, st.store_code) AS store_code,
        safe_to_date(st.date::text) AS dt,
        SUM(st.target) AS target

    FROM sales_targets st

    LEFT JOIN store_code_mapping scm
        ON st.store_code = scm.old_store_code

    WHERE COALESCE(scm.store_code, st.store_code) NOT IN ({EXCLUDED_STORES_SQL})
    AND st.date IS NOT NULL
    AND st.date::text <> 'NAN'

    GROUP BY
        COALESCE(scm.store_code, st.store_code),
        safe_to_date(st.date::text)

),

combined AS (

    SELECT
        COALESCE(b.store_code, t.store_code) AS store_code,
        COALESCE(b.dt, t.dt) AS dt,
        COALESCE(b.sales, 0) AS sales,
        COALESCE(t.target, 0) AS target

    FROM base b

    FULL OUTER JOIN target_base t
        ON b.store_code = t.store_code
       AND b.dt = t.dt

),

combined_with_spoc AS (

    SELECT
        c.store_code,
        ss.spoc_name,
        c.dt,
        c.sales,
        c.target
    FROM combined c
    LEFT JOIN store_spoc ss
        ON c.store_code = ss.store_code

    UNION ALL

    -- stores with zero sales AND zero target activity in the whole
    -- window still show up (blank/NULL achievement) instead of
    -- vanishing from the report entirely
    SELECT
        ss.store_code,
        ss.spoc_name,
        NULL AS dt,
        NULL AS sales,
        NULL AS target
    FROM store_spoc ss
    WHERE ss.store_code NOT IN (SELECT store_code FROM combined)

),

achievement_report AS (

SELECT

    store_code,
    spoc_name,
    GROUPING(spoc_name)  AS grp_spoc,
    GROUPING(store_code) AS grp_store,

    /* =========================
       LAST 7 COMPLETED DAYS
    ========================= */

    ROUND(
        100.0 *
        SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '1 day' THEN sales END)
        /
        NULLIF(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '1 day' THEN target END),0),
    2) AS day_1_achievement,

    ROUND(
        100.0 *
        SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '2 day' THEN sales END)
        /
        NULLIF(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '2 day' THEN target END),0),
    2) AS day_2_achievement,

    ROUND(
        100.0 *
        SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '3 day' THEN sales END)
        /
        NULLIF(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '3 day' THEN target END),0),
    2) AS day_3_achievement,

    ROUND(
        100.0 *
        SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '4 day' THEN sales END)
        /
        NULLIF(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '4 day' THEN target END),0),
    2) AS day_4_achievement,

    ROUND(
        100.0 *
        SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '5 day' THEN sales END)
        /
        NULLIF(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '5 day' THEN target END),0),
    2) AS day_5_achievement,

    ROUND(
        100.0 *
        SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '6 day' THEN sales END)
        /
        NULLIF(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '6 day' THEN target END),0),
    2) AS day_6_achievement,

    ROUND(
        100.0 *
        SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '7 day' THEN sales END)
        /
        NULLIF(SUM(CASE WHEN dt = CURRENT_DATE - INTERVAL '7 day' THEN target END),0),
    2) AS day_7_achievement,

    /* =========================
       LAST 4 COMPLETED CALENDAR WEEKS
    ========================= */

    ROUND(
        100.0 *
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                 AND dt < date_trunc('week', CURRENT_DATE)::date
                THEN sales
            END
        )
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                     AND dt < date_trunc('week', CURRENT_DATE)::date
                    THEN target
                END
            ),
            0
        ),
    2) AS week_1_achievement,

    ROUND(
        100.0 *
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                 AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                THEN sales
            END
        )
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                     AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '7 day'
                    THEN target
                END
            ),
            0
        ),
    2) AS week_2_achievement,

    ROUND(
        100.0 *
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                 AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                THEN sales
            END
        )
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                     AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '14 day'
                    THEN target
                END
            ),
            0
        ),
    2) AS week_3_achievement,

    ROUND(
        100.0 *
        SUM(
            CASE
                WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '28 day'
                 AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                THEN sales
            END
        )
        /
        NULLIF(
            SUM(
                CASE
                    WHEN dt >= date_trunc('week', CURRENT_DATE)::date - INTERVAL '28 day'
                     AND dt < date_trunc('week', CURRENT_DATE)::date - INTERVAL '21 day'
                    THEN target
                END
            ),
            0
        ),
    2) AS week_4_achievement,
	    /* =========================
       LAST 3 COMPLETED MONTHS
    ========================= */

    ROUND(
        100.0 *
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '1 month'
                THEN sales
            END
        )
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '1 month'
                    THEN target
                END
            ),
            0
        ),
    2) AS month_1_achievement,

    ROUND(
        100.0 *
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '2 month'
                THEN sales
            END
        )
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '2 month'
                    THEN target
                END
            ),
            0
        ),
    2) AS month_2_achievement,

    ROUND(
        100.0 *
        SUM(
            CASE
                WHEN date_trunc('month', dt)
                     = date_trunc('month', CURRENT_DATE) - INTERVAL '3 month'
                THEN sales
            END
        )
        /
        NULLIF(
            SUM(
                CASE
                    WHEN date_trunc('month', dt)
                         = date_trunc('month', CURRENT_DATE) - INTERVAL '3 month'
                    THEN target
                END
            ),
            0
        ),
    2) AS month_3_achievement

FROM combined_with_spoc

GROUP BY GROUPING SETS
(
    (spoc_name, store_code),
    (spoc_name),
    ()
)

),

store_master AS (

    SELECT
        store_code,
        MAX(store_name) AS store_name
    FROM b2c_sale
    WHERE store_code NOT IN ({EXCLUDED_STORES_SQL})
    GROUP BY store_code

)

SELECT

    CASE
        WHEN a.grp_spoc  = 1 THEN 'MINISO'
        WHEN a.grp_store = 1 THEN COALESCE(a.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE a.store_code
    END AS store_code,

    CASE
        WHEN a.grp_spoc  = 1 THEN 'MINISO'
        WHEN a.grp_store = 1 THEN COALESCE(a.spoc_name, '(No SPOC)') || ' - TOTAL'
        ELSE sm.store_name
    END AS store_name,

    CASE
        WHEN a.grp_store = 1 THEN NULL
        ELSE a.spoc_name
    END AS spoc_name,

    a.day_1_achievement,
    a.day_2_achievement,
    a.day_3_achievement,
    a.day_4_achievement,
    a.day_5_achievement,
    a.day_6_achievement,
    a.day_7_achievement,

    a.week_1_achievement,
    a.week_2_achievement,
    a.week_3_achievement,
    a.week_4_achievement,

    a.month_1_achievement,
    a.month_2_achievement,
    a.month_3_achievement

FROM achievement_report a

LEFT JOIN store_master sm
    ON a.store_code = sm.store_code

ORDER BY
    a.grp_spoc,
    COALESCE(a.spoc_name, 'zzz_no_spoc_assigned'),
    a.grp_store,
    a.store_code;
        """
      }
]

# ==========================================
# Timestamp
# ==========================================

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# ==========================================
# Generate Reports
# ==========================================

# Collected in memory so the JSON refresh can be atomic: nothing is
# published unless EVERY report succeeded.
collected_reports = {}
failed_reports = []

for report in reports:

    try:

        print(f"\n▶ Running Report : {report['file']}")

        # Create folder if not exists
        folder_path = os.path.join(base_folder, report["folder"])
        os.makedirs(folder_path, exist_ok=True)

        # Execute SQL Query
        # df_raw keeps the original SQL aliases (day_1_abv, ...) so the
        # JSON can carry a stable machine key; df below is the identical
        # frame with human date headers, exactly as before, for Excel.
        df_raw = pd.read_sql_query(
            report["query"],
            engine
        )

        df = df_raw

        print(f"✅ Rows Retrieved : {len(df):,}")

        # Rename day_1/week_1/month_1-style columns to their actual
        # calendar date/period for the human-facing Excel file.
        # store_code / store_name / spoc_name / target_* columns are
        # untouched — only the 14 metric columns get relabeled.
        rename_map = build_date_rename_map(df_raw.columns)
        if rename_map:
            df = df_raw.rename(columns=rename_map)

        # Create timestamp file name
        file_name = report["file"].replace(
            ".xlsx",
            f"_{timestamp}.xlsx"
        )

        # Full output path
        file_path = os.path.join(
            folder_path,
            file_name
        )

        # Save Excel
        with pd.ExcelWriter(
            file_path,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                sheet_name="Data",
                index=False
            )

        print(f"✅ Saved Successfully")
        print(file_path)

        # Hand the SAME frame that was just written to Excel to the
        # JSON layer. No recomputation anywhere.
        collected_reports[report["key"]] = jsonlayer.build_report_block(
            report_key=report["key"],
            config=report,
            df=df_raw,
            rename_map=rename_map,
            excel_path=file_path,
        )

    except Exception:

        print(f"\n❌ ERROR IN REPORT : {report['file']}")
        traceback.print_exc()
        failed_reports.append(report["key"])

# ==========================================
# Close Connection
# ==========================================

engine.dispose()

# ==========================================
# JSON Refresh  (atomic, all-or-nothing)
# ==========================================
# Policy: a full refresh only. If ANY report failed, the previous
# ops_kpi_data.json is left exactly as it was -- the dashboard keeps
# showing the last complete data set rather than a partial one.

expected_report_keys = [r["key"] for r in reports]

print("\n" + "=" * 52)
print("JSON DATA LAYER")
print("=" * 52)

if failed_reports:
    print(
        "❌ JSON NOT refreshed. These reports failed: "
        + ", ".join(failed_reports)
    )
    print(f"   Previous JSON left untouched: {JSON_OUTPUT_PATH}")
else:
    try:
        payload = jsonlayer.build_payload(
            report_blocks={k: collected_reports[k] for k in expected_report_keys},
            reporting_date=date.today(),
            excluded_stores=EXCLUDED_STORES,
        )

        warnings = jsonlayer.validate_payload(payload, expected_report_keys)

        jsonlayer.write_atomic(JSON_OUTPUT_PATH, payload)

        for report_key in expected_report_keys:
            block = payload["reports"][report_key]
            counts = block["row_counts_by_type"]
            print(
                f"   {block['name']:<20} "
                f"{block['row_count']:>5} rows  "
                f"({counts['store']} stores, "
                f"{counts['spoc_total']} SPoC totals, "
                f"{counts['miniso_total']} MINISO total)"
            )

        for warning in warnings:
            print(f"   ⚠  {warning}")

        size_kb = os.path.getsize(JSON_OUTPUT_PATH) / 1024
        print(f"\n✅ JSON refreshed ({size_kb:,.0f} KB)")
        print(f"   {JSON_OUTPUT_PATH}")
        print(
            f"   Refreshed at "
            f"{payload['metadata']['last_successful_refresh']}"
        )

    except jsonlayer.JsonValidationError as exc:
        print(f"❌ JSON validation failed: {exc}")
        print(f"   Previous JSON left untouched: {JSON_OUTPUT_PATH}")

    except Exception:
        print("❌ JSON refresh failed")
        traceback.print_exc()
        print(f"   Previous JSON left untouched: {JSON_OUTPUT_PATH}")

# ==========================================
# Summary
# ==========================================

if failed_reports:
    print(
        f"\n⚠  Finished with {len(failed_reports)} failed report(s): "
        + ", ".join(failed_reports)
    )
else:
    print("\n🎉 All Reports Generated Successfully")