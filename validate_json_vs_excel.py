"""
validate_json_vs_excel.py
=========================

Reconciles data/ops_kpi_data.json against the Excel workbooks produced by
the same OPS_KPI.py run. Read-only: it never writes to the database, the
Excel files or the JSON.

The JSON records the exact workbook filename each report was written to,
so this compares like with like rather than guessing at the newest file.

Checks, per report:
    1.  the report exists in JSON
    2.  the workbook exists on disk
    3.  row counts match
    4.  column count and order match
    5.  JSON column labels == Excel headers
    6.  every cell matches, store by store, including nulls
    7.  numerics are still numerics in JSON
    8.  store / SPoC-total / MINISO rows line up one-for-one
    9.  no duplicate store codes introduced
   10.  MINISO and SPoC totals match cell for cell

Run:
    python validate_json_vs_excel.py
    python validate_json_vs_excel.py --json data/ops_kpi_data.json --base "J:\\...\\ops kpi"
"""

import argparse
import json
import math
import os
import sys

import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOLERANCE = 1e-9   # exact-match expected; this only absorbs float repr noise


def find_workbook(base_folder, report_key, filename):
    """The run wrote <base>/<folder>/<filename>; folder == report key."""
    direct = os.path.join(base_folder, report_key, filename)
    if os.path.exists(direct):
        return direct
    for root, _dirs, files in os.walk(base_folder):
        if filename in files:
            return os.path.join(root, filename)
    return None


def values_equal(json_value, excel_value):
    json_null = json_value is None
    excel_null = excel_value is None or (
        isinstance(excel_value, float) and math.isnan(excel_value)
    ) or pd.isna(excel_value)

    if json_null or excel_null:
        return json_null and excel_null

    if isinstance(json_value, (int, float)) and isinstance(excel_value, (int, float)):
        if math.isclose(float(json_value), float(excel_value),
                        rel_tol=TOLERANCE, abs_tol=TOLERANCE):
            return True
        return False

    return str(json_value).strip() == str(excel_value).strip()


def reconcile_report(report_key, block, base_folder, max_diffs=10):
    issues = []
    notes = []

    filename = block.get("excel_file")
    if not filename:
        issues.append("JSON does not record an Excel filename for this report")
        return issues, notes

    path = find_workbook(base_folder, report_key, filename)
    if not path:
        issues.append(f"workbook not found anywhere under the base folder: {filename}")
        return issues, notes

    notes.append(f"workbook: {path}")
    excel = pd.read_excel(path, sheet_name=block.get("excel_sheet", "Data"))
    rows = block["data"]

    # --- shape
    if len(excel) != len(rows):
        issues.append(f"row count: Excel {len(excel)} vs JSON {len(rows)}")
    else:
        notes.append(f"rows: {len(rows)} (match)")

    excel_headers = [str(c) for c in excel.columns]
    json_labels = [c["label"] for c in block["columns"]]
    if excel_headers != json_labels:
        only_excel = [c for c in excel_headers if c not in json_labels]
        only_json = [c for c in json_labels if c not in excel_headers]
        issues.append(
            "column headers differ. "
            f"Excel-only: {only_excel or 'none'}; JSON-only: {only_json or 'none'}; "
            f"order {'differs' if sorted(excel_headers) == sorted(json_labels) else 'n/a'}"
        )
        return issues, notes
    notes.append(f"columns: {len(json_labels)} (labels and order match)")

    # --- row alignment: both are in the same SQL ORDER BY sequence
    keys = [c["key"] for c in block["columns"]]
    label_to_key = dict(zip(json_labels, keys))

    excel_codes = [str(v) for v in excel[excel_headers[0]].tolist()]
    json_codes = [str(r.get("store_code")) for r in rows]
    if excel_codes != json_codes:
        mismatched = [(i, a, b) for i, (a, b) in
                      enumerate(zip(excel_codes, json_codes)) if a != b][:5]
        issues.append(f"store order/identity differs at rows {mismatched}")
        return issues, notes

    dupes = {c for c in json_codes if json_codes.count(c) > 1}
    if dupes:
        issues.append(f"duplicate store codes in JSON: {sorted(dupes)[:10]}")

    # --- cell-by-cell
    diffs = []
    compared = 0
    for position, (json_row, (_, excel_row)) in enumerate(zip(rows, excel.iterrows())):
        for label in excel_headers:
            key = label_to_key[label]
            compared += 1
            if not values_equal(json_row.get(key), excel_row[label]):
                diffs.append(
                    f"row {position + 2} [{json_row.get('store_code')}] "
                    f"column '{label}' ({key}): "
                    f"Excel={excel_row[label]!r} JSON={json_row.get(key)!r}"
                )
    if diffs:
        issues.append(f"{len(diffs)} cell mismatch(es):")
        issues.extend("      " + d for d in diffs[:max_diffs])
        if len(diffs) > max_diffs:
            issues.append(f"      ... and {len(diffs) - max_diffs} more")
    else:
        notes.append(f"cells: {compared:,} compared, all equal")

    # --- hierarchy
    counts = block["row_counts_by_type"]
    excel_miniso = sum(1 for c in excel_codes if c == "MINISO")
    excel_totals = sum(1 for c in excel_codes if c.endswith(" - TOTAL"))
    if excel_miniso != counts["miniso_total"]:
        issues.append(
            f"MINISO rows: Excel {excel_miniso} vs JSON {counts['miniso_total']}"
        )
    if excel_totals != counts["spoc_total"]:
        issues.append(
            f"SPoC total rows: Excel {excel_totals} vs JSON {counts['spoc_total']}"
        )
    if not issues:
        notes.append(
            f"hierarchy: {counts['store']} stores, {counts['spoc_total']} SPoC "
            f"totals, {counts['miniso_total']} MINISO total (match)"
        )

    # --- types
    bad_types = []
    for col in block["columns"]:
        if col["dtype"] != "number":
            continue
        for row in rows:
            v = row.get(col["key"])
            if v is not None and not isinstance(v, (int, float)):
                bad_types.append(f"{col['key']} on {row.get('store_code')}")
                break
    if bad_types:
        issues.append(f"numeric values stored as text: {bad_types[:5]}")

    return issues, notes


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--json",
                        default=os.environ.get(
                            "OPS_KPI_JSON_PATH",
                            os.path.join(SCRIPT_DIR, "data", "ops_kpi_data.json")))
    parser.add_argument("--base", default=os.environ.get("OPS_KPI_BASE_FOLDER"))
    args = parser.parse_args()

    if not os.path.exists(args.json):
        print(f"JSON not found: {args.json}")
        print("Run OPS_KPI.py first.")
        return 2

    with open(args.json, encoding="utf-8") as fh:
        payload = json.load(fh)

    base_folder = args.base
    if not base_folder:
        print("Base folder not given. Pass --base \"<path to ops kpi folder>\" "
              "or set OPS_KPI_BASE_FOLDER.")
        return 2
    if not os.path.isdir(base_folder):
        print(f"Base folder not found: {base_folder}")
        return 2

    meta = payload["metadata"]
    print("=" * 64)
    print("JSON <-> EXCEL RECONCILIATION")
    print("=" * 64)
    print(f"JSON        : {args.json}")
    print(f"Refreshed   : {meta.get('last_successful_refresh')}")
    print(f"Base folder : {base_folder}")

    failed = []
    for report_key in meta.get("report_order", list(payload["reports"])):
        block = payload["reports"].get(report_key)
        print(f"\n--- {report_key} ({block['name'] if block else 'MISSING'})")
        if block is None:
            print("  FAIL  report missing from JSON")
            failed.append(report_key)
            continue

        issues, notes = reconcile_report(report_key, block, base_folder)
        for note in notes:
            print(f"  ok    {note}")
        for issue in issues:
            print(f"  FAIL  {issue}")
        if issues:
            failed.append(report_key)

    print("\n" + "=" * 64)
    if failed:
        print(f"RECONCILIATION FAILED for: {', '.join(failed)}")
        return 1
    print("RECONCILED: JSON matches Excel for every report, row and cell.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
