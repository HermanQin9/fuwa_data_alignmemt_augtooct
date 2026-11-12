"""Independently audit every original cell and reconcile a built report or prepared draft.

The report window is read from the report's own Validation sheet, so the same audit works for the
all-dates draft and for a dated final report: every source row is still checked cell by cell, and
rows outside the window must be marked as such and kept out of the pivot.
"""
import argparse
import datetime as dt
import json
from collections import Counter
from itertools import zip_longest
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / "USA PMix.xlsx"
RESULT = ROOT.parent / "Result"
# Restated here rather than imported, so the audit stays independent of the workflow.
LOCATION_ALIASES = {"Fluffy Fluffy Tucker GA": "Fluffy Fluffy - Tucker"}
# The source column that holds the time of sale, per POS sheet.
TIME_COLUMNS = {"TOAST": "Sent Date - Copy", "SQUARE": "Time"}
MONEY = ["Gross Sales", "Discount", "Net Sales", "Tax"]


def records(sheet):
    rows = sheet.iter_rows(values_only=True)
    headers = next(rows)
    for row in rows:
        yield dict(zip(headers, row))


def blank(value):
    return "" if value is None else value


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mapping", type=Path, default=RESULT / "USA PMix - Menu Mapping.xlsx")
    parser.add_argument("--output", type=Path, default=RESULT / "USA PMix - PMIX Report DRAFT.xlsx")
    parser.add_argument("--source", type=Path, default=SOURCE)
    args = parser.parse_args()
    source = openpyxl.load_workbook(args.source, read_only=True, data_only=True)
    mapping = openpyxl.load_workbook(args.mapping, read_only=True, data_only=True)
    output = openpyxl.load_workbook(args.output, read_only=True, data_only=True)
    try:
        stated = {r["Check"]: r["Value"] for r in records(output["Validation"])}
        first, last = (dt.date.fromisoformat(str(stated[k]))
                       for k in ("Inclusive start date", "Inclusive end date"))
        master = {}
        mapping_rows = mapping_qty = 0
        for row in records(mapping["Menu Category"]):
            key = tuple(blank(row[c]) for c in ("POS Sheet", "Source Group", "Source Item"))
            assert key not in master, f"Duplicate mapping source identity: {key!r}"
            master[key] = row
            mapping_rows += row["Rows"]
            mapping_qty += row["Net Qty"]

        expected = {}
        expected_money = {}
        category_totals = Counter()
        report_totals = Counter()
        mapping_counts = Counter()
        mapping_quantities = Counter()
        source_counts = Counter()
        expected_exceptions = set()
        original_cells = 0
        source_qty = 0
        all_qty = 0  # net Qty inside the window
        for name, group, item, date in [
            ("TOAST", "Menu Group", "Menu Item", "Sent Date"),
            ("SQUARE", "Category", "Item", "Date"),
        ]:
            time_column = TIME_COLUMNS[name]
            src_rows = source[name].iter_rows(values_only=True)
            dst_rows = output[name].iter_rows(values_only=True)
            headers, dest_headers = next(src_rows), next(dst_rows)
            assert dest_headers[:len(headers)] == headers
            indices = {c: headers.index(c) for c in headers}
            sentinel = object()
            for excel_row, pair in enumerate(zip_longest(src_rows, dst_rows, fillvalue=sentinel), 2):
                original, copied = pair
                assert original is not sentinel and copied is not sentinel, f"{name}: different row counts"
                for col, value in enumerate(original):
                    actual = copied[col]
                    original_cells += 1
                    if value in ("", None) and actual in ("", None):
                        continue  # Excel stores an empty string and an empty cell alike
                    assert value == actual and type(value) is type(actual), (
                        f"{name}!row {excel_row}, column {col + 1}: "
                        f"{value!r} ({type(value).__name__}) != {actual!r} ({type(actual).__name__})"
                    )
                key = (name, blank(original[indices[group]]), original[indices[item]])
                assert key in master, f"Missing mapping: {key!r}"
                rule = master[key]
                qty = original[indices["Qty"]]
                sale_date = original[indices[date]].date()
                location = original[indices["Location"]].strip()
                location = LOCATION_ALIASES.get(location, location)
                row_id = (name, excel_row)
                source_counts[name] += 1
                mapping_counts[key] += 1
                mapping_quantities[key] += qty
                source_qty += qty
                in_window = first <= sale_date <= last
                for column, value in (
                    ("Mapped Category", blank(rule["Category"]) if in_window else ""),
                    ("Mapped Review Status", rule["Review Status"] if in_window else ""),
                    ("In Report Window", in_window),
                ):
                    assert blank(copied[dest_headers.index(column)]) == value, (
                        f"{name}!row {excel_row}: wrong {column}"
                    )
                if not in_window:
                    continue
                when = original[indices[time_column]] if time_column in indices else original[indices[date]]
                sale_time = when.time() if isinstance(when, dt.datetime) else when
                expected[row_id] = (rule, qty, sale_date, location, sale_time)
                expected_money[row_id] = tuple(
                    original[indices[c]] if c in indices else None for c in MONEY)
                all_qty += qty
                if rule["Category"] and rule["Review Status"] != "Excluded":
                    quarter = (sale_date.month - 1) // 3 + 1
                    period = f"{sale_date.year} Q{quarter}"
                    category_totals[(rule["Category"], period)] += qty
                    quarter_name = f"{quarter}{('st', 'nd', 'rd', 'th')[quarter - 1]} Quarter"
                    report_totals[(
                        location, sale_date.year, quarter_name, rule["Category"],
                        rule["Menu"], rule["Menu group"], rule["Item"],
                    )] += qty
                else:
                    expected_exceptions.add(row_id)
        assert mapping_rows == sum(source_counts.values())
        assert mapping_qty == source_qty
        for key, rule in master.items():
            assert rule["Rows"] == mapping_counts[key], f"Wrong mapping row count: {key!r}"
            assert rule["Net Qty"] == mapping_quantities[key], f"Wrong mapping Qty: {key!r}"

        seen = set()
        actual_categories = Counter()
        pivot_qty = 0
        for row in records(output["PMIXData"]):
            row_id = (row["POS Sheet"], row["Source Row"])
            assert row_id not in seen, f"Duplicate pivot row {row_id}"
            seen.add(row_id)
            rule, qty, sale_date, location, sale_time = expected[row_id]
            assert rule["Category"], f"Unresolved row incorrectly in pivot: {row_id}"
            for field in ("Category", "Menu", "Menu group", "Item", "Mapping ID", "Review Status"):
                assert row[field] == rule[field], f"Mapping not applied: {row_id} {field}"
            assert row["Time"] == sale_time, f"Wrong sale time: {row_id}"
            for field in ("Location", "Category", "Menu", "Menu group", "Item"):
                assert str(row[field]) == str(row[field]).strip(), f"Untrimmed {field}: {row_id}"
            if "Menu List Section" in row and blank(rule.get("Menu List Section")):
                assert row["Menu List Section"] == rule["Menu List Section"], f"Wrong section: {row_id}"
            assert row["Qty"] == qty and row["Location"] == location
            for column, value in zip(MONEY, expected_money[row_id]):
                assert blank(row.get(column)) == blank(value), f"Wrong {column}: {row_id}"
            assert isinstance(row["Date"], dt.datetime) and row["Date"].date() == sale_date
            quarter = (sale_date.month - 1) // 3 + 1
            assert row["Year"] == sale_date.year
            assert row["Week starting"] == (sale_date - dt.timedelta(days=sale_date.weekday())).isoformat()
            assert row["Month"] == sale_date.strftime("%Y-%m")
            assert row["Period"] == f"{sale_date.year} Q{quarter}"
            actual_categories[(row["Category"], row["Period"])] += qty
            pivot_qty += qty
        assert actual_categories == category_totals
        exceptions = list(records(output["Exceptions"]))
        exception_ids = {(r["POS Sheet"], r["Source Row"]) for r in exceptions}
        assert exception_ids == expected_exceptions
        assert len(exception_ids) == len(exceptions)
        assert seen.isdisjoint(exception_ids)
        assert seen | exception_ids == set(expected)
        for row in exceptions:
            row_id = (row["POS Sheet"], row["Source Row"])
            rule, qty, sale_date, location, sale_time = expected[row_id]
            assert row["Qty"] == qty and row["Location"] == location
            for column, value in zip(MONEY, expected_money[row_id]):
                assert blank(row.get(column)) == blank(value), f"Wrong {column} in Exceptions: {row_id}"
            assert row["Time"] == sale_time, f"Wrong sale time in Exceptions: {row_id}"
            assert row["Date"].date() == sale_date
            for field in ("Category", "Menu", "Menu group", "Item", "Mapping ID", "Review Status"):
                assert blank(row[field]) == blank(rule[field]), (
                    f"Exception changed source/mapping value: {row_id} {field}"
                )
            reason = (rule["Review Notes"] if rule["Review Status"] == "Excluded"
                      else "Unresolved mapping")
            assert row["Exception Reason"] == reason
        exception_qty = sum(r["Qty"] for r in exceptions)
        assert pivot_qty + exception_qty == all_qty

        actual_reports = Counter()
        report_keys = set()
        for row in records(output["Report Data"]):
            key = tuple(row[c] for c in (
                "Location", "Year", "Quarter", "Category", "Menu", "Menu group", "Item",
            ))
            assert key not in report_keys, f"Duplicate Report Data key: {key!r}"
            report_keys.add(key)
            actual_reports[key] += row["Qty sold"]
        assert actual_reports == report_totals
        metadata = stated
        for field, value in (
            ("Source item-sales rows", sum(source_counts.values())),
            ("Selected item-sales rows", len(expected)),
            ("Selected net Qty", all_qty),
            ("Pivot item-sales rows", len(seen)),
            ("Pivot net Qty", pivot_qty),
            ("Exception rows (not in pivot)", len(exceptions)),
            ("Exception net Qty", exception_qty),
        ):
            assert metadata[field] == value, f"Incorrect Validation value: {field}"
        result = {
            "report_window": [first.isoformat(), last.isoformat()],
            "source_rows": dict(source_counts), "source_net_qty": source_qty,
            "selected_rows": len(expected), "selected_net_qty": all_qty,
            "original_cells_compared_with_type_check": original_cells,
            "mapping_entries": len(master), "pivot_source_rows": len(seen),
            "pivot_source_net_qty": pivot_qty, "exception_rows": len(exceptions),
            "exception_net_qty": exception_qty,
            "observed_location_labels": len({v[3] for v in expected.values()}),
            "detailed_report_keys_checked": len(report_keys),
            "per_mapping_counts_and_qty": "PASS",
            "exceptions_field_level_checks": "PASS",
            "year_quarter_month_checks": "PASS",
            "validation_metadata_checks": "PASS",
            "data_checks": "PASS",
            "semantic_mapping_approval": (
                "Pending business review" if any(r["Review Status"] == "Pending" for r in master.values())
                else "No Pending entries"),
            "native_pivot_and_filter_tests": "See verify_result.py",
        }
        print(json.dumps(result, indent=2))
    finally:
        source.close()
        mapping.close()
        output.close()


if __name__ == "__main__":
    main()
