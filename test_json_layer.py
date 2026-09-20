"""
Unit tests for the JSON serialisation layer.

These use SYNTHETIC frames with the same *shape* as the real reports.
They prove the serialiser, the row classifier, the validator and the
atomic write behave correctly. They say nothing about your real KPI
values -- that reconciliation is validate_json_vs_excel.py, which runs
against the actual Excel files the script produces.

Run:  python tests/test_json_layer.py
"""

import json
import os
import sys
from datetime import date
from decimal import Decimal

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ops_kpi_json as jl  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
FAILURES = []


def check(condition, message):
    if condition:
        print(f"  pass  {message}")
    else:
        print(f"  FAIL  {message}")
        FAILURES.append(message)


def make_abv_frame():
    """Shape of the abv report: dimensions + target + 7d/4w/3m."""
    rows = [
        # store rows
        ("INAA", "Store A", "Spoc One", Decimal("450.00"),
         Decimal("501.25"), None, np.float64(487.5), 470.0, 455.1, 460.0, 462.0,
         480.0, 478.0, 476.0, 474.0, 479.0, 481.0, 483.0),
        ("INAB", "Store B", "Spoc One", Decimal("450.00"),
         np.nan, 399.0, 401.0, 402.0, 403.0, 404.0, 405.0,
         410.0, 411.0, 412.0, 413.0, 414.0, 415.0, 416.0),
        ("INAC", "Store C", "Spoc Two", None,
         None, None, None, None, None, None, None,
         None, None, None, None, None, None, None),
        # spoc subtotal rows: spoc_name and target NULL by design
        ("Spoc One - TOTAL", "Spoc One - TOTAL", None, None,
         450.1, 449.0, 448.0, 447.0, 446.0, 445.0, 444.0,
         443.0, 442.0, 441.0, 440.0, 439.0, 438.0, 437.0),
        ("Spoc Two - TOTAL", "Spoc Two - TOTAL", None, None,
         None, None, None, None, None, None, None,
         None, None, None, None, None, None, None),
        # grand total
        ("MINISO", "MINISO", None, None,
         460.0, 459.0, 458.0, 457.0, 456.0, 455.0, 454.0,
         453.0, 452.0, 451.0, 450.0, 449.0, 448.0, 447.0),
    ]
    columns = (
        ["store_code", "store_name", "spoc_name", "target_abv"]
        + [f"day_{i}_abv" for i in range(1, 8)]
        + [f"week_{i}_abv" for i in range(1, 5)]
        + [f"month_{i}_abv" for i in range(1, 4)]
    )
    return pd.DataFrame(rows, columns=columns)


def make_noh_frame():
    """Shape of the noh report: short aliases, no target column."""
    columns = (
        ["store_code", "store_name", "spoc_name"]
        + [f"d{i}_neg_sku" for i in range(1, 8)]
        + [f"w{i}_neg_sku" for i in range(1, 5)]
        + [f"m{i}_neg_sku" for i in range(1, 4)]
    )
    rows = [
        ["INAA", "Store A", "Spoc One"] + [float(i) for i in range(14)],
        ["Spoc One - TOTAL", "Spoc One - TOTAL", None] + [float(i) for i in range(14)],
        ["MINISO", "MINISO", None] + [float(i) for i in range(14)],
    ]
    return pd.DataFrame(rows, columns=columns)


def build_date_rename_map_reference(columns):
    """Copy of the renamer in OPS_KPI.py, so the test exercises the real
    label grammar without importing the script (which connects to a DB)."""
    import re
    from datetime import timedelta

    def shift_month(first_of_month, months_back):
        total = first_of_month.month - 1 - months_back
        return date(first_of_month.year + total // 12, total % 12 + 1, 1)

    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    month_start = date(today.year, today.month, 1)
    day_pat = re.compile(r"^(?:day_|d)(\d)_(.+)$")
    week_pat = re.compile(r"^(?:week_|w)(\d)_(.+)$")
    month_pat = re.compile(r"^(?:month_|m)(\d)_(.+)$")

    rename = {}
    for col in columns:
        m = day_pat.match(col)
        if m:
            rename[col] = (today - timedelta(days=int(m.group(1)))).strftime("%d-%b-%Y")
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
            rename[col] = shift_month(month_start, int(m.group(1))).strftime("%b-%Y")
    return rename


def main():
    abv = make_abv_frame()
    noh = make_noh_frame()

    abv_map = build_date_rename_map_reference(abv.columns)
    noh_map = build_date_rename_map_reference(noh.columns)

    print("\n1. Column renaming covers every metric column")
    check(len(abv_map) == 14, "abv: 14 metric columns relabelled")
    check(len(noh_map) == 14, "noh: 14 short-alias columns relabelled")
    check("target_abv" not in abv_map, "abv: target column not relabelled")
    check("store_code" not in abv_map, "abv: dimensions not relabelled")

    print("\n2. Period labels point at completed periods")
    from datetime import timedelta
    today = date.today()
    check(abv_map["day_1_abv"] == (today - timedelta(days=1)).strftime("%d-%b-%Y"),
          "day_1 is yesterday, not today")
    week_start = today - timedelta(days=today.weekday())
    expected_end = week_start - timedelta(days=1)
    check(abv_map["week_1_abv"].endswith(expected_end.strftime("%d-%b-%Y")),
          "week_1 ends on the last completed Sunday")
    check(abv_map["month_1_abv"] != today.strftime("%b-%Y"),
          "month_1 is not the current incomplete month")

    print("\n3. Row classification")
    block = jl.build_report_block("abv", {"display_name": "ABV",
                                          "target_column": "target_abv"},
                                  abv, abv_map, "/tmp/abv.xlsx")
    counts = block["row_counts_by_type"]
    check(counts == {"store": 3, "spoc_total": 2, "miniso_total": 1},
          f"abv row types {counts}")
    total_row = [r for r in block["data"] if r["_row_type"] == "miniso_total"][0]
    check(total_row["store_code"] == "MINISO", "MINISO row identified")
    spoc_row = [r for r in block["data"] if r["_row_type"] == "spoc_total"][0]
    check(spoc_row["_spoc"] == "Spoc One", "SPoC recovered for subtotal row")
    check(spoc_row["spoc_name"] is None, "original spoc_name left as NULL")

    print("\n4. Value fidelity")
    a = [r for r in block["data"] if r["store_code"] == "INAA"][0]
    check(a["target_abv"] == 450.0 and isinstance(a["target_abv"], float),
          "Decimal target -> float")
    check(a["day_1_abv"] == 501.25, "Decimal metric keeps precision")
    check(a["day_2_abv"] is None, "SQL NULL stays null, not 0")
    check(a["day_3_abv"] == 487.5, "numpy float unwrapped")
    b = [r for r in block["data"] if r["store_code"] == "INAB"][0]
    check(b["day_1_abv"] is None, "NaN -> null")
    c = [r for r in block["data"] if r["store_code"] == "INAC"][0]
    check(all(c[f"day_{i}_abv"] is None for i in range(1, 8)),
          "all-null store is kept, not dropped")

    print("\n5. Labels in JSON == Excel headers")
    excel_headers = list(abv.rename(columns=abv_map).columns)
    json_labels = [c["label"] for c in block["columns"]]
    check(json_labels == excel_headers, "column labels match Excel exactly")
    check([c["key"] for c in block["columns"]] == list(abv.columns),
          "machine keys preserved alongside labels")

    print("\n6. Period index")
    days = block["periods"]["day"]
    check([p["index"] for p in days] == [1, 2, 3, 4, 5, 6, 7], "days ordered 1..7")
    check(len(block["periods"]["week"]) == 4 and len(block["periods"]["month"]) == 3,
          "4 weeks, 3 months")

    print("\n7. Validation catches bad payloads")
    noh_block = jl.build_report_block("noh", {"display_name": "Negative stock"},
                                      noh, noh_map, "/tmp/noh.xlsx")
    payload = jl.build_payload({"abv": block, "noh": noh_block},
                               date.today(), ["INFD"])
    warnings = jl.validate_payload(payload, ["abv", "noh"])
    check(True, f"valid payload accepted ({len(warnings)} warning(s))")

    def expect_error(mutate, label):
        import copy
        bad = copy.deepcopy(payload)
        mutate(bad)
        try:
            jl.validate_payload(bad, ["abv", "noh"])
        except jl.JsonValidationError:
            check(True, label)
        else:
            check(False, label)

    expect_error(lambda p: p["reports"].pop("noh"), "missing report rejected")
    expect_error(lambda p: p["reports"]["abv"].__setitem__("data", []),
                 "empty report rejected")
    expect_error(lambda p: p["reports"]["abv"]["data"].append(
        dict(p["reports"]["abv"]["data"][0])), "duplicate row rejected")
    expect_error(lambda p: p["reports"]["abv"]["data"][0].__setitem__(
        "day_1_abv", "501.25"), "stringified number rejected")
    expect_error(lambda p: p["reports"]["abv"]["data"].remove(
        [r for r in p["reports"]["abv"]["data"]
         if r["_row_type"] == "miniso_total"][0]),
        "missing MINISO total rejected")
    expect_error(lambda p: p["reports"]["abv"]["data"][0].__setitem__(
        "day_1_abv", float("nan")), "NaN in payload rejected")

    print("\n8. Atomic write protects the previous file")
    target = os.path.join(HERE, "_tmp_atomic.json")
    jl.write_atomic(target, payload)
    first = os.path.getsize(target)
    try:
        jl.write_atomic(target, {"bad": {1, 2}})  # a set is not serialisable
    except TypeError:
        pass
    check(os.path.getsize(target) == first, "failed write left the old file intact")
    leftovers = [f for f in os.listdir(HERE) if f.startswith(".ops_kpi_data_")]
    check(not leftovers, "no temp files left behind")
    with open(target, encoding="utf-8") as fh:
        reloaded = json.load(fh)
    check(reloaded["reports"]["abv"]["data"][0]["day_1_abv"] == 501.25,
          "round-trip through disk preserves values")
    os.remove(target)

    print()
    if FAILURES:
        print(f"{len(FAILURES)} FAILED")
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
