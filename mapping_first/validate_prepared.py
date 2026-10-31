"""Independently audit every original cell and reconcile the prepared draft."""
import datetime as dt
import json
from collections import Counter
from itertools import zip_longest
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / "USA PMix.xlsx"
MAPPING = ROOT / "USA PMix - Menu Mapping.xlsx"
OUTPUT = ROOT / "USA PMix - Pivot Source.xlsx"


def records(sheet):
    rows = sheet.iter_rows(values_only=True)
    headers = next(rows)
    for row in rows:
        yield dict(zip(headers, row))


def blank(value):
    return "" if value is None else value


def main():
    source = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
    mapping = openpyxl.load_workbook(MAPPING, read_only=True, data_only=True)
    output = openpyxl.load_workbook(OUTPUT, read_only=True, data_only=True)
    try:
        master = {}
        mapping_rows = mapping_qty = 0
        for row in records(mapping["Menu Category"]):
            key = tuple(blank(row[c]) for c in ("POS Sheet", "Source Group", "Source Item"))
            assert key not in master, f"Duplicate mapping source identity: {key!r}"
            master[key] = row
            mapping_rows += row["Rows"]
            mapping_qty += row["Net Qty"]

        expected = {}
        category_totals = Counter()
        report_totals = Counter()
        mapping_counts = Counter()
        mapping_quantities = Counter()
        source_counts = Counter()
        expected_exceptions = set()
        original_cells = 0
        all_qty = 0
        for name, group, item, date in [
            ("TOAST", "Menu Group", "Menu Item", "Sent Date"),
            ("SQUARE", "Category", "Item", "Date"),
        ]:
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
                    assert value == actual and type(value) is type(actual), (
                        f"{name}!row {excel_row}, column {col + 1}: "
                        f"{value!r} ({type(value).__name__}) != {actual!r} ({type(actual).__name__})"
                    )
                    original_cells += 1
                key = (name, blank(original[indices[group]]), original[indices[item]])
                assert key in master, f"Missing mapping: {key!r}"
                rule = master[key]
                qty = original[indices["Qty"]]
                sale_date = original[indices[date]].date()
                location = original[indices["Location"]].strip()
                row_id = (name, excel_row)
                expected[row_id] = (rule, qty, sale_date, location)
                source_counts[name] += 1
                mapping_counts[key] += 1
                mapping_quantities[key] += qty
                all_qty += qty
                for column, value in (
                    ("Mapped Category", blank(rule["Category"])),
                    ("Mapped Review Status", rule["Review Status"]),
                    ("In Report Window", True),
                ):
                    assert blank(copied[dest_headers.index(column)]) == value, (
                        f"{name}!row {excel_row}: wrong {column}"
                    )
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
        assert mapping_qty == all_qty
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
            rule, qty, sale_date, location = expected[row_id]
            assert rule["Category"], f"Unresolved row incorrectly in pivot: {row_id}"
            for field in ("Category", "Menu", "Menu group", "Item", "Mapping ID", "Review Status"):
                assert row[field] == rule[field], f"Mapping not applied: {row_id} {field}"
            assert row["Qty"] == qty and row["Location"] == location
            assert isinstance(row["Date"], dt.datetime) and row["Date"].date() == sale_date
            quarter = (sale_date.month - 1) // 3 + 1
            assert row["Year"] == sale_date.year
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
            rule, qty, sale_date, location = expected[row_id]
            assert row["Qty"] == qty and row["Location"] == location
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
        metadata = {r["Check"]: r["Value"] for r in records(output["Validation"])}
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
        assert source_counts == {"TOAST": 185823, "SQUARE": 6984}
        assert all_qty == 202865
        result = {
            "source_rows": dict(source_counts), "source_net_qty": all_qty,
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
            "semantic_mapping_approval": "Pending business review",
            "native_pivot_and_filter_tests": "NOT RUN - Excel COM sandbox access denied",
        }
        print(json.dumps(result, indent=2))
    finally:
        source.close()
        mapping.close()
        output.close()


if __name__ == "__main__":
    main()
