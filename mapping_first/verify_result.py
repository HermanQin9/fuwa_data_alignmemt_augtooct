"""Check the built report in live Excel against an independent recompute from the source.

Works on a temporary copy, so the delivered workbook is never saved or changed. It drives the
real slicers, then compares every result with totals calculated directly from
USA PMix.xlsx plus the saved mapping. Writes Result/VALIDATION_RESULTS.json, which holds overall
totals and pass/fail results only.
"""
import datetime as dt
import hashlib
import json
import shutil
import tempfile
from pathlib import Path

import re

import pandas as pd
import win32com.client

BASE = Path(__file__).resolve().parent.parent
SOURCE = BASE / "USA PMix.xlsx"
RESULT = BASE / "Result"
REPORT = RESULT / "USA PMix - PMIX Report DRAFT.xlsx"
MAPPING = RESULT / "USA PMix - Menu Mapping.xlsx"
TUCKER = {"Fluffy Fluffy Tucker GA": "Fluffy Fluffy - Tucker"}
COLUMNS = {"TOAST": ("Menu Group", "Menu Item", "Sent Date"), "SQUARE": ("Category", "Item", "Date")}


def mapping_id(sheet, group, item):
    key = json.dumps([sheet, "" if pd.isna(group) else str(group), item], ensure_ascii=True)
    return hashlib.sha256(key.encode("utf-8")).hexdigest()


def expected_rows():
    parts = []
    for sheet, (group, item, date) in COLUMNS.items():
        frame = pd.read_excel(SOURCE, sheet_name=sheet, dtype=object)
        dates = pd.to_datetime(frame[date])
        parts.append(pd.DataFrame({
            "id": [mapping_id(sheet, g, i) for g, i in zip(frame[group], frame[item])],
            "location": frame["Location"].astype(str).str.strip().replace(TUCKER),
            "qty": frame["Qty"].astype(int),
            "day": dates.dt.normalize(),
            "month": dates.dt.strftime("%Y-%m"),
            "week": (dates.dt.normalize() - pd.to_timedelta(dates.dt.weekday, unit="D")).dt.strftime("%Y-%m-%d"),
        }))
    rows = pd.concat(parts, ignore_index=True)
    mapping = pd.read_excel(MAPPING, dtype=object).fillna("")
    rows = rows.merge(mapping[["Mapping ID", "Category", "Review Status"]],
                      left_on="id", right_on="Mapping ID", how="left", validate="many_to_one")
    assert rows["Category"].notna().all(), "A source row has no mapping"
    return rows, mapping


class Live:
    """The open report, with the slicers addressed by the field they filter."""

    def __init__(self, book):
        self.sheet = book.Worksheets("PMIX")
        self.pivot = self.sheet.PivotTables("MenuPMIX")
        self.caches = {cache.SourceName: cache for cache in book.SlicerCaches}

    def totals_row(self, offset):
        """Sum the Total row across periods: offset 0 is Qty (#), offset 1 is Qty (%)."""
        body = self.pivot.DataBodyRange
        last = body.Rows.Count
        values = [body.Cells(last, c).Value for c in range(1 + offset, body.Columns.Count + 1, 2)]
        return sum(v for v in values if isinstance(v, (int, float)))

    def total(self):
        return int(round(self.totals_row(0)))

    def percent(self):
        return float(self.totals_row(1))

    def clear(self):
        for name in ("Category", "Location", "Month", "Week starting"):
            self.caches[name].ClearManualFilter()

    def choose(self, field, names):
        items = self.caches[field].SlicerItems
        for item in items:
            if item.Name in names:
                item.Selected = True
        for item in items:
            if item.Name not in names:
                item.Selected = False


CHINESE = re.compile("[⺀-鿿＀-￯]")


def chinese_cells(sheet):
    """Texts in a sheet that contain Chinese characters, such as Excel's own pivot labels."""
    found = []
    values = sheet.UsedRange.Value
    for row in values if isinstance(values, tuple) else ((values,),):
        for value in row:
            if isinstance(value, str) and CHINESE.search(value):
                found.append(value)
    return found


def overlaps(a, b):
    return not (a[0] + a[2] <= b[0] or b[0] + b[2] <= a[0]
                or a[1] + a[3] <= b[1] or b[1] + b[3] <= a[1])


def run_excel(rows):
    classified = rows[rows["Category"].ne("")]
    checks, failures = {}, []

    def check(name, got, want):
        checks[name] = checks.get(name, 0) + 1
        if got != want:
            failures.append(f"{name}: got {got!r}, expected {want!r}")

    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    with tempfile.TemporaryDirectory() as directory:
        copy = Path(directory) / "report_copy.xlsx"
        shutil.copy2(REPORT, copy)
        try:
            excel.Visible = False
            excel.DisplayAlerts = False
            book = excel.Workbooks.Open(str(copy))
            live = Live(book)
            pivot = live.pivot
            by_location = book.Worksheets("By Location")
            grid = by_location.PivotTables("MenuByLocation")
            everything = int(classified["qty"].sum())
            check("grand total", live.total(), everything)
            check("no page fields on PMIX", pivot.PageFields.Count, 0)
            check("no page fields on By Location", grid.PageFields.Count, 0)
            check("By Location grand total", int(grid.GetPivotData("Qty (#)").Value), everything)
            for store, qty in classified.groupby("location")["qty"].sum().items():
                check("By Location column per store",
                      int(grid.GetPivotData("Qty (#)", "Location", store).Value), int(qty))
            header = [by_location.Cells(grid.TableRange1.Row + 1, c).Text for c in range(1, 30)]
            check("All locations column exists", "All locations" in header, True)
            check("no Chinese on PMIX", chinese_cells(live.sheet), [])
            check("no Chinese on By Location", chinese_cells(by_location), [])
            categories = sorted(i.Name for i in pivot.PivotFields("Category").PivotItems())
            check("category list", categories, sorted(classified["Category"].unique()))

            for name, qty in classified.groupby("Category")["qty"].sum().items():
                live.choose("Category", {name})
                check("single category", live.total(), int(qty))
                check("By Location follows category", int(grid.GetPivotData("Qty (#)").Value), int(qty))
                check("filtered percent is 100%", round(live.percent(), 9), 1.0)
            live.clear()

            for name, qty in classified.groupby("location")["qty"].sum().items():
                live.choose("Location", {name})
                check("single store", live.total(), int(qty))
            live.clear()

            for name, qty in classified.groupby("month")["qty"].sum().items():
                live.choose("Month", {name})
                check("single month", live.total(), int(qty))
            live.clear()

            stores = sorted(classified["location"].unique())
            pick_stores, pick_cats = set(stores[:3]), {"Pancakes", "Sandwiches"}
            live.choose("Location", pick_stores)
            live.choose("Category", pick_cats)
            live.choose("Month", {"2026-09"})
            wanted = classified[classified["location"].isin(pick_stores)
                                & classified["Category"].isin(pick_cats)
                                & classified["month"].eq("2026-09")]
            check("several stores, categories and a month", live.total(), int(wanted["qty"].sum()))
            live.clear()

            weeks = {"2026-09-07", "2026-09-14"}
            live.choose("Week starting", weeks)
            wanted = classified[classified["week"].isin(weeks)]
            check("week range", live.total(), int(wanted["qty"].sum()))
            check("By Location follows week", int(grid.GetPivotData("Qty (#)").Value), int(wanted["qty"].sum()))
            for week, qty in classified.groupby("week")["qty"].sum().items():
                live.choose("Week starting", {week})
                check("single week", live.total(), int(qty))
            live.clear()
            check("all filters cleared", live.total(), int(classified["qty"].sum()))

            item = pivot.PivotFields("Category").PivotItems("Sandwiches")
            collapsed_rows = pivot.TableRange1.Rows.Count
            item.ShowDetail = True
            expanded_rows = pivot.TableRange1.Rows.Count
            check("expand adds rows", expanded_rows > collapsed_rows, True)
            check("expand keeps total", live.total(), int(classified["qty"].sum()))
            item.ShowDetail = False
            check("collapse restores rows", pivot.TableRange1.Rows.Count, collapsed_rows)

            table = pivot.TableRange2
            box = (table.Left, table.Top, table.Width, table.Height)
            shapes = [(s.Left, s.Top, s.Width, s.Height) for s in live.sheet.Shapes]
            check("slicer count", len(shapes), 4)
            clashes = sum(overlaps(a, b) for n, a in enumerate(shapes) for b in shapes[n + 1:])
            check("slicers overlap each other", clashes, 0)
            check("slicers overlap the table", sum(overlaps(a, box) for a in shapes), 0)
            table2 = grid.TableRange2
            box2 = (table2.Left, table2.Top, table2.Width, table2.Height)
            shapes2 = [(s.Left, s.Top, s.Width, s.Height) for s in by_location.Shapes]
            check("By Location slicer count", len(shapes2), 3)
            check("By Location slicers overlap each other",
                  sum(overlaps(a, b) for n, a in enumerate(shapes2) for b in shapes2[n + 1:]), 0)
            check("By Location slicers overlap the table", sum(overlaps(a, box2) for a in shapes2), 0)
            check("exception notice intact",
                  str(live.sheet.Range("A5").Value).startswith("EXCEPTIONS"), True)
        finally:
            if book is not None:
                book.Close(SaveChanges=False)
            excel.Quit()
    return checks, failures


def main():
    rows, mapping = expected_rows()
    checks, failures = run_excel(rows)
    classified = rows[rows["Category"].ne("")]
    blank = mapping[mapping["Category"].eq("")]
    result = {
        "status": "DRAFT - all entries still need business approval",
        "source_rows": len(rows),
        "source_net_qty": int(rows["qty"].sum()),
        "pivot_rows": len(classified),
        "pivot_net_qty": int(classified["qty"].sum()),
        "exception_rows": len(rows) - len(classified),
        "exception_net_qty": int(rows["qty"].sum() - classified["qty"].sum()),
        "mapping_entries": len(mapping),
        "entries_pending_review": int(mapping["Review Status"].eq("Pending").sum()),
        "entries_excluded": int(mapping["Review Status"].eq("Excluded").sum()),
        "entries_with_no_category": len(blank),
        "report_locations": int(rows["location"].nunique()),
        "categories_in_report": int(classified["Category"].nunique()),
        "date_window": [str(rows["day"].min().date()), str(rows["day"].max().date())],
        "live_excel_checks": checks,
        "live_excel_checks_run": sum(checks.values()),
        "live_excel_failures": failures,
        "overall": "PASS" if not failures else "FAIL",
    }
    (RESULT / "VALIDATION_RESULTS.json").write_text(
        json.dumps(result, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2, ensure_ascii=True))
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
