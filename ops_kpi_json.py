"""
ops_kpi_json.py
===============

JSON data layer for the OPS KPI reporting system.

This module does NOT calculate any KPI. It is a pure serialisation /
validation layer: it takes the SAME pandas DataFrames that OPS_KPI.py
already writes to Excel and turns them into a single JSON document that
the HTML dashboard reads.

Design rules honoured here
--------------------------
1. No KPI logic. No SQL. No re-aggregation. Values are copied verbatim
   from the DataFrame produced by pd.read_sql_query().
2. Column identity is preserved twice:
      key   -> the original SQL alias  (day_1_abv, w3_neg_sku, ...)
      label -> the human calendar label the Excel header shows
               (produced by build_date_rename_map() in OPS_KPI.py and
               passed in here, so both outputs use the SAME labels).
3. Row identity is preserved. Individual stores, SPoC subtotal rows and
   the single MINISO grand-total row are tagged with `row_type` so the
   dashboard can never mistake an aggregate for a store. The tag is
   derived from the store_code value the SQL itself emits
   ('MINISO' / '<spoc> - TOTAL'), it is not a new business rule.
4. Atomic refresh. The payload is validated in memory, written to a
   temp file in the same directory, and only then os.replace()'d over
   the live file. A failed run leaves the previous JSON untouched.
5. NaN / NaT / numpy scalars / Decimal / date / datetime are converted
   to JSON-safe values. SQL NULL stays null — it is never coerced to 0.
"""

from __future__ import annotations

import json
import math
import os
import re
import tempfile
from datetime import date, datetime, time
from decimal import Decimal

import pandas as pd

SCHEMA_VERSION = "1.0"

# Same alias grammar the Excel renamer uses, so the two never drift.
_DAY_PAT = re.compile(r"^(?:day_|d)(\d)_(.+)$")
_WEEK_PAT = re.compile(r"^(?:week_|w)(\d)_(.+)$")
_MONTH_PAT = re.compile(r"^(?:month_|m)(\d)_(.+)$")

DIMENSION_COLUMNS = ("store_code", "store_name", "spoc_name")

MINISO_ROW_LABEL = "MINISO"
SPOC_TOTAL_SUFFIX = " - TOTAL"


# ----------------------------------------------------------------------
# value conversion
# ----------------------------------------------------------------------

def to_jsonable(value):
    """Convert one DataFrame cell into a JSON-safe Python value.

    NULL-ish input (None, NaN, NaT, pd.NA) -> None, never 0 and never "".
    Numerics stay numerics -- they are never stringified.
    """
    if value is None:
        return None

    # pd.isna raises on list-likes; guard it.
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, (bool,)):
        return bool(value)

    if isinstance(value, Decimal):
        f = float(value)
        return None if math.isnan(f) else f

    if isinstance(value, (datetime, pd.Timestamp)):
        return value.isoformat()

    if isinstance(value, date):
        return value.isoformat()

    if isinstance(value, time):
        return value.isoformat()

    # numpy scalars expose .item()
    item = getattr(value, "item", None)
    if callable(item) and hasattr(value, "dtype"):
        value = item()

    if isinstance(value, float):
        return None if (math.isnan(value) or math.isinf(value)) else value

    if isinstance(value, int):
        return value

    if isinstance(value, str):
        return value

    return str(value)


def _is_numeric_series(series: pd.Series) -> bool:
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        return True
    # object columns holding Decimal come back from some drivers
    non_null = series.dropna()
    if len(non_null) == 0:
        return False
    return all(isinstance(v, (int, float, Decimal)) and not isinstance(v, bool)
               for v in non_null.head(50))


# ----------------------------------------------------------------------
# column + row metadata
# ----------------------------------------------------------------------

def parse_metric_key(column: str):
    """Return (granularity, period_index, metric) for a metric alias.

    'day_1_abv'   -> ('day', 1, 'abv')
    'w3_neg_sku'  -> ('week', 3, 'neg_sku')
    'target_abv'  -> ('target', None, 'abv')
    'store_code'  -> (None, None, None)
    """
    m = _DAY_PAT.match(column)
    if m:
        return "day", int(m.group(1)), m.group(2)
    m = _WEEK_PAT.match(column)
    if m:
        return "week", int(m.group(1)), m.group(2)
    m = _MONTH_PAT.match(column)
    if m:
        return "month", int(m.group(1)), m.group(2)
    if column.startswith("target_"):
        return "target", None, column[len("target_"):]
    return None, None, None


def build_column_metadata(df: pd.DataFrame, rename_map: dict) -> list:
    """Describe every column of the raw (pre-rename) DataFrame.

    `rename_map` is exactly the dict OPS_KPI.py feeds into df.rename()
    before writing Excel, so `label` here == the Excel header.
    """
    columns = []
    for col in df.columns:
        granularity, index, metric = parse_metric_key(col)

        if col in DIMENSION_COLUMNS:
            role = "dimension"
        elif granularity == "target":
            role = "target"
        elif granularity in ("day", "week", "month"):
            role = "metric"
        else:
            role = "attribute"

        columns.append({
            "key": col,
            "label": rename_map.get(col, col),
            "role": role,
            "granularity": granularity,
            "period_index": index,
            "metric": metric,
            "dtype": "number" if _is_numeric_series(df[col]) else "string",
        })
    return columns


def classify_row(store_code) -> str:
    """store | spoc_total | miniso_total, from the value SQL emitted."""
    if store_code is None:
        return "store"
    text = str(store_code)
    if text == MINISO_ROW_LABEL:
        return "miniso_total"
    if text.endswith(SPOC_TOTAL_SUFFIX):
        return "spoc_total"
    return "store"


def _spoc_for_row(row_type: str, store_code, spoc_name):
    """SPoC a row belongs to, for dashboard filtering only.

    Subtotal rows carry spoc_name = NULL by design (the SQL nulls it so
    Excel doesn't read as if the total were a store's own row). The SPoC
    name is still recoverable from the store_code label the SQL built,
    so the dashboard can group without changing any stored value.
    spoc_name itself is left exactly as the query returned it.
    """
    if row_type == "spoc_total":
        return str(store_code)[: -len(SPOC_TOTAL_SUFFIX)]
    if row_type == "miniso_total":
        return None
    return spoc_name


def dataframe_to_records(df: pd.DataFrame) -> list:
    """Row-wise JSON-safe records keyed by the ORIGINAL SQL column names."""
    records = []
    columns = list(df.columns)
    for row in df.itertuples(index=False, name=None):
        record = {col: to_jsonable(val) for col, val in zip(columns, row)}
        row_type = classify_row(record.get("store_code"))
        record["_row_type"] = row_type
        record["_spoc"] = _spoc_for_row(
            row_type, record.get("store_code"), record.get("spoc_name")
        )
        records.append(record)
    return records


# ----------------------------------------------------------------------
# payload assembly
# ----------------------------------------------------------------------

def build_report_block(report_key: str, config: dict, df: pd.DataFrame,
                       rename_map: dict, excel_path: str | None) -> dict:
    records = dataframe_to_records(df)
    counts = {"store": 0, "spoc_total": 0, "miniso_total": 0}
    for r in records:
        counts[r["_row_type"]] += 1

    columns = build_column_metadata(df, rename_map)

    periods = {}
    for col in columns:
        if col["granularity"] in ("day", "week", "month"):
            periods.setdefault(col["granularity"], []).append({
                "key": col["key"],
                "index": col["period_index"],
                "label": col["label"],
            })
    for granularity in periods:
        periods[granularity].sort(key=lambda p: p["index"])

    return {
        "key": report_key,
        "name": config.get("display_name", report_key),
        "description": config.get("description", ""),
        "unit": config.get("unit"),
        "value_format": config.get("value_format", "number"),
        "higher_is_better": config.get("higher_is_better"),
        "target_column": config.get("target_column"),
        "achievement_basis": config.get("achievement_basis"),
        "excel_file": os.path.basename(excel_path) if excel_path else None,
        "excel_sheet": "Data",
        "row_count": len(records),
        "row_counts_by_type": counts,
        "columns": columns,
        "periods": periods,
        "data": records,
    }


def build_payload(report_blocks: dict, reporting_date, excluded_stores,
                  generated_at=None) -> dict:
    generated_at = generated_at or datetime.now()
    return {
        "metadata": {
            "report_name": "OPS KPI Dashboard",
            "schema_version": SCHEMA_VERSION,
            "generated_at": generated_at.isoformat(timespec="seconds"),
            "last_successful_refresh": generated_at.isoformat(timespec="seconds"),
            "reporting_date": reporting_date.isoformat()
                if hasattr(reporting_date, "isoformat") else str(reporting_date),
            "status": "success",
            "refresh_mode": "atomic_full",
            "row_type_legend": {
                "store": "An individual store.",
                "spoc_total": "Subtotal for one SPoC. Not a store.",
                "miniso_total": "MINISO grand total. Not a store.",
            },
            "excluded_stores": list(excluded_stores),
            "report_order": list(report_blocks.keys()),
        },
        "reports": report_blocks,
    }


# ----------------------------------------------------------------------
# validation + atomic write
# ----------------------------------------------------------------------

class JsonValidationError(Exception):
    pass


def validate_payload(payload: dict, expected_reports) -> list:
    """Raise JsonValidationError on anything that would publish bad data.

    Returns a list of non-fatal warnings.
    """
    warnings = []
    expected = list(expected_reports)

    if "metadata" not in payload or "reports" not in payload:
        raise JsonValidationError("payload is missing 'metadata' or 'reports'")

    reports = payload["reports"]
    missing = [r for r in expected if r not in reports]
    if missing:
        raise JsonValidationError(f"reports missing from payload: {missing}")

    for key in expected:
        block = reports[key]

        if not block.get("columns"):
            raise JsonValidationError(f"[{key}] has no column metadata")

        rows = block.get("data")
        if rows is None:
            raise JsonValidationError(f"[{key}] has no data array")
        if len(rows) == 0:
            raise JsonValidationError(f"[{key}] returned zero rows")
        if len(rows) != block.get("row_count"):
            raise JsonValidationError(
                f"[{key}] row_count {block.get('row_count')} != {len(rows)} rows"
            )

        col_keys = [c["key"] for c in block["columns"]]
        if len(col_keys) != len(set(col_keys)):
            raise JsonValidationError(f"[{key}] duplicate column keys")
        labels = [c["label"] for c in block["columns"]]
        if len(labels) != len(set(labels)):
            raise JsonValidationError(
                f"[{key}] duplicate column labels -- Excel headers would collide"
            )

        counts = block["row_counts_by_type"]
        if counts["miniso_total"] != 1:
            raise JsonValidationError(
                f"[{key}] expected exactly 1 MINISO total row, "
                f"found {counts['miniso_total']}"
            )
        if counts["store"] == 0:
            raise JsonValidationError(f"[{key}] contains no store-level rows")

        seen = set()
        for row in rows:
            identity = (row.get("store_code"), row["_row_type"])
            if identity in seen:
                raise JsonValidationError(
                    f"[{key}] duplicate row for {identity[0]}"
                )
            seen.add(identity)

        numeric_cols = [c["key"] for c in block["columns"] if c["dtype"] == "number"]
        for row in rows:
            for col in numeric_cols:
                v = row.get(col)
                if v is not None and not isinstance(v, (int, float)):
                    raise JsonValidationError(
                        f"[{key}] {col} became {type(v).__name__} for "
                        f"{row.get('store_code')} -- numbers must stay numbers"
                    )

        all_null = [c for c in numeric_cols
                    if all(row.get(c) is None for row in rows)]
        if all_null:
            warnings.append(f"[{key}] every value is null in: {', '.join(all_null)}")

    # payload must actually serialise
    try:
        json.dumps(payload, allow_nan=False)
    except (TypeError, ValueError) as exc:
        raise JsonValidationError(f"payload is not JSON-serialisable: {exc}")

    return warnings


def write_atomic(path: str, payload: dict) -> None:
    """Write JSON via temp file + os.replace so a crash can't corrupt the
    previous good file."""
    directory = os.path.dirname(os.path.abspath(path))
    os.makedirs(directory, exist_ok=True)

    text = json.dumps(payload, ensure_ascii=False, allow_nan=False, indent=1)

    fd, tmp_path = tempfile.mkstemp(
        prefix=".ops_kpi_data_", suffix=".json.tmp", dir=directory
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp_path, path)
    except BaseException:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise
