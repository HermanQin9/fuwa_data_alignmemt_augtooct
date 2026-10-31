"""Build an editable menu mapping, then a native, filterable Excel PMIX report.

Run with the Anaconda interpreter documented in README.md. Excel for Windows
and pywin32 are required for the report command, not the mapping command.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd


ROOT = Path(__file__).resolve().parent.parent
CATEGORIES = [
    "Pancakes", "Fluffy All Day", "Drinks", "Japanese Eats",
    "Cheesecakes", "Pastries", "Promotions", "Add-Ons", "Sandwiches",
]  # Sandwiches added at Cheena's request (2026-10-31); the other eight are unchanged
# Cheena (2026-10-31): the Square POS Tucker label is the same store as the Toast one.
LOCATION_ALIASES = {"Fluffy Fluffy Tucker GA": "Fluffy Fluffy - Tucker"}
SOURCE_COLUMNS = {
    "TOAST": ("Menu Group", "Menu Item", "Sent Date"),
    "SQUARE": ("Category", "Item", "Date"),
}
KEY_COLUMNS = ["POS Sheet", "Source Group", "Source Item"]
MAPPING_COLUMNS = [
    "Mapping ID", *KEY_COLUMNS, "Category", "Menu", "Menu group", "Item",
    "Review Status", "Evidence", "Candidate Paths", "Rows", "Net Qty",
    "Locations", "Review Notes",
]
TOP_CATEGORIES = {
    "All Pancakes": "Pancakes",
    "All Fluffy All Day": "Fluffy All Day",
    "All Beverages": "Drinks",
    "All Promotions": "Promotions",
    "All Add Ons": "Add-Ons",
    "Japanese Eats": "Japanese Eats",
}


def text(value: object) -> str:
    return "" if pd.isna(value) else str(value)


def normalized(value: object) -> str:
    result = text(value).strip().lower().lstrip("#").replace("&", " and ")
    result = unicodedata.normalize("NFKD", result)
    result = "".join(c for c in result if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", result).strip()


def mapping_id(sheet: str, group: object, item: object) -> str:
    key = json.dumps([sheet, text(group), text(item)], ensure_ascii=True)
    return hashlib.sha256(key.encode("utf-8")).hexdigest()


def load_sources(source: Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    book = openpyxl.load_workbook(source, read_only=True, data_only=True)
    sheets = {}
    try:
        for name in SOURCE_COLUMNS:
            rows = book[name].iter_rows(values_only=True)
            sheets[name] = pd.DataFrame(rows, columns=next(rows), dtype=object)
    finally:
        book.close()
    parts = []
    for name, df in sheets.items():
        group, item, date = SOURCE_COLUMNS[name]
        required = {group, item, date, "Qty", "Location"}
        missing = required - set(df.columns)
        if missing:
            raise ValueError(f"{name}: missing columns {sorted(missing)}")
        for column in (item, "Location"):
            if df[column].map(lambda v: not text(v).strip()).any():
                raise ValueError(f"{name}: blank {column}; source needs correction")
        qty = pd.to_numeric(df["Qty"], errors="raise")
        if qty.isna().any() or not qty.map(lambda v: float(v).is_integer()).all():
            raise ValueError(f"{name}: Qty must contain nonblank integer quantities")
        dates = pd.to_datetime(df[date], errors="raise")
        if dates.isna().any():
            raise ValueError(f"{name}: missing sale dates")
        part = pd.DataFrame({
            "POS Sheet": name, "Source Row": df.index + 2,
            "Source Group": df[group].map(text), "Source Item": df[item],
            "Location": df["Location"].map(
                lambda v: LOCATION_ALIASES.get(text(v).strip(), text(v).strip())),
            "Date": dates.dt.normalize(), "Qty": qty,
        })
        part["Mapping ID"] = [
            mapping_id(name, g, i) for g, i in zip(df[group], df[item])
        ]
        parts.append(part)
    return sheets, pd.concat(parts, ignore_index=True)


def read_reference(source: Path) -> pd.DataFrame:
    records = []
    top = menu = group = ""
    book = openpyxl.load_workbook(source, read_only=True, data_only=True)
    try:
        for row in book["Menu Category"].iter_rows():
            if len(row) < 5 or row[4].value is None:
                continue
            cell = row[4]
            name = text(cell.value).strip()
            level = int(cell.alignment.indent or 0)
            if level == 0:
                top, menu, group = name, "", ""
            elif level in (1, 2):
                menu, group = name, ""
            elif level == 3:
                group = name
            elif level == 4:
                category = TOP_CATEGORIES.get(top, "")
                if top == "All Dessert/Croffles":
                    category = ("Cheesecakes" if group == "Cheesecake and Ice Cream"
                                else "Pastries")
                if normalized(group) in ("add ons", "all add ons"):
                    category = "Add-Ons"
                records.append({
                    "Reference Row": cell.row, "Reference Top": top,
                    "Category": category, "Menu": menu,
                    "Menu group": group, "Item": name,
                })
    finally:
        book.close()
    if not records:
        raise ValueError("No indented item records found in Menu Category")
    return pd.DataFrame(records)


def item_hint(group: str, item: str) -> str:
    """Conservative draft suggestions; none of these are business approvals."""
    i, g = normalized(item), normalized(group)
    if re.search(r"\b(promo|promotion|combo|deal)\b", i + " " + g):
        return "Promotions"
    if re.search(r"\bbuy\s+(?:\d+|any|one|two)\b.*\b(get|free)\b", i):
        return "Promotions"
    if ("+" in text(item)
            and re.search(r"\b(pancake|pancakes|croffle|croffles)\b", i)
            and re.search(r"\b(latte|frappe|coffee|smoothie|refresher)\b", i)):
        return "Promotions"
    if re.search(r"\b(add|extra|side|tots|hash|syrup|sauce)\b", i):
        if not re.search(r"\b(sandwich|benedict|bowl)\b", i):
            return "Add-Ons"
    if re.search(r"\b(cheesecake|cheesecakes|basque)\b", i) or "cheese cake" in i:
        return "Cheesecakes"
    if re.search(r"\b(latte|frappe|coffee|espresso|tea|refresher|water|milk|juice|boba)\b", i):
        return "Drinks"
    if re.search(r"\b(benedict|sandwich|breakfast|savory|omelette)\b", i):
        return "Fluffy All Day"
    if "croissant" in i and re.search(r"\b(egg|bacon|ham|sausage|turkey|salmon|avocado)\b", i):
        return "Fluffy All Day"
    if "chicken" in i and re.search(r"\b(pancake|pancakes|croffle|croffles)\b", i):
        return "Fluffy All Day"
    if re.search(r"\b(pancake|pancakes)\b", i) or "french toast" in i:
        return "Pancakes"
    if "croffle" in i:
        return ("Fluffy All Day" if re.search(r"\b(egg|bacon|ham|salmon|avocado|chicken)\b", i)
                else "Pastries")
    if re.search(r"\b(bowl|donburi|karaage|katsu|gyoza|takoyaki)\b", i):
        return "Japanese Eats"
    if re.search(r"\b(macaron|cookie|cookies|mousse|flan|brownie|croissant)\b", i):
        return "Pastries"
    if "roll cake" in i:
        return "Pastries"
    aliases = {
        "fluffy pancakes": "Pancakes", "sweet pancakes": "Pancakes",
        "souffle pancakes": "Pancakes", "pancakes": "Pancakes",
        "fluffy all day": "Fluffy All Day", "fluffy all day long": "Fluffy All Day",
        "savory": "Fluffy All Day", "brunch": "Fluffy All Day",
        "coffee": "Drinks", "drinks": "Drinks", "boba": "Drinks",
        "iced tea": "Drinks", "refreshers": "Drinks", "refresher": "Drinks",
        "latte": "Drinks", "espresso lattes": "Drinks",
        "side orders": "Add-Ons", "side order": "Add-Ons", "sides": "Add-Ons",
        "sweet croffles": "Pastries", "pastries": "Pastries",
        "basque cheesecake": "Cheesecakes", "cheesecakes": "Cheesecakes",
    }
    return aliases.get(g, "")


def propose_mapping(source: pd.DataFrame, reference: pd.DataFrame) -> pd.DataFrame:
    lookup = {}
    for record in reference.to_dict("records"):
        lookup.setdefault(normalized(record["Item"]), []).append(record)
    rows = []
    for key, frame in source.groupby(["Mapping ID", *KEY_COLUMNS], dropna=False, sort=False):
        ident, sheet, group, item = key
        candidates = lookup.get(normalized(item), [])
        contextual = [r for r in candidates
                      if normalized(r["Menu group"]) == normalized(group)]
        if contextual:
            candidates = contextual
        else:
            regular = [r for r in candidates
                       if r["Category"] and r["Category"] != "Promotions"
                       and not r["Menu"].endswith("(GH)")]
            if regular:
                candidates = regular
        categories = sorted({r["Category"] for r in candidates if r["Category"]})
        hint = item_hint(group, item)
        category = categories[0] if len(categories) == 1 else hint
        if len(categories) > 1:
            category = ""
        usable = [r for r in candidates if r["Category"] == category]
        selected = usable[0] if usable else None
        evidence = ("Reference item + group" if contextual else "Reference item") if candidates else "Draft rule"
        if not category:
            evidence = "Ambiguous / needs review"
        paths = sorted({" > ".join([r["Category"] or "Unresolved", r["Menu"],
                                    r["Menu group"], r["Item"]]) for r in candidates})
        rows.append({
            "Mapping ID": ident, "POS Sheet": sheet, "Source Group": group,
            "Source Item": item, "Category": category,
            "Menu": selected["Menu"] if selected else category,
            "Menu group": selected["Menu group"] if selected else (text(group).strip() or category),
            "Item": selected["Item"] if selected else item,
            "Review Status": "Pending", "Evidence": evidence,
            "Candidate Paths": "\n".join(paths),
            "Rows": len(frame), "Net Qty": frame["Qty"].sum(),
            "Locations": "; ".join(sorted(frame["Location"].unique())),
            "Review Notes": ("Suggested mapping, not approved. Review category and menu path."
                             if category else "Insufficient evidence; do not force a category."),
        })
    return pd.DataFrame(rows, columns=MAPPING_COLUMNS)


def write_table(writer, name: str, frame: pd.DataFrame, autofilter: bool = True) -> None:
    frame.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    ws.freeze_panes(1, 0)
    if autofilter:
        ws.autofilter(0, 0, len(frame), len(frame.columns) - 1)
    ws.set_column(0, len(frame.columns) - 1, 22)
    for col in ("Source Item", "Item", "Candidate Paths", "Review Notes", "Locations"):
        if col in frame:
            index = frame.columns.get_loc(col)
            ws.set_column(index, index, 42)
    time_format = writer.book.add_format({"num_format": "hh:mm:ss"})
    for column, values in enumerate(frame.items()):
        _, series = values
        if series.dtype == object:
            for row, value in enumerate(series, 1):
                if isinstance(value, dt.time):
                    ws.write_datetime(row, column, value, time_format)
                elif isinstance(value, str) and value == "":
                    ws.write_string(row, column, "")


def make_mapping(source_path: Path, output: Path) -> dict:
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite an editable mapping: {output}")
    _, source = load_sources(source_path)
    reference = read_reference(source_path)
    mapping = propose_mapping(source, reference)
    coverage = source.groupby(["POS Sheet", "Location"], dropna=False).agg(
        Rows=("Qty", "size"), Net_Qty=("Qty", "sum"),
        First_Date=("Date", "min"), Last_Date=("Date", "max"),
    ).reset_index()
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="xlsxwriter",
                        engine_kwargs={"options": {"strings_to_formulas": False,
                                                   "strings_to_urls": False}}) as writer:
        write_table(writer, "Menu Category", mapping)
        write_table(writer, "Reference Outline", reference)
        write_table(writer, "Source Coverage", coverage)
        write_table(writer, "Allowed Categories", pd.DataFrame({"Category": CATEGORIES}))
        ws = writer.sheets["Menu Category"]
        ws.data_validation(1, 4, len(mapping), 4, {
            "validate": "list", "source": "='Allowed Categories'!$A$2:$A$%d" % (len(CATEGORIES) + 1),
        })
        ws.data_validation(1, 8, len(mapping), 8, {
            "validate": "list", "source": ["Pending", "Approved", "Excluded"],
        })
    return {"mapping_rows": len(mapping), "source_rows": len(source),
            "net_qty": int(source["Qty"].sum()),
            "suggested_rows": int(mapping["Category"].ne("").sum()),
            "unresolved_rows": int(mapping["Category"].eq("").sum())}


def apply_mapping(source: pd.DataFrame, mapping: pd.DataFrame, draft: bool) -> pd.DataFrame:
    required = set(MAPPING_COLUMNS)
    if required - set(mapping):
        raise ValueError(f"Mapping missing columns: {sorted(required - set(mapping))}")
    if mapping["Mapping ID"].duplicated().any():
        raise ValueError("Duplicate Mapping ID values in Menu Category")
    mapping = mapping.copy()
    for column in mapping:
        mapping[column] = mapping[column].map(lambda v: "" if pd.isna(v) else v)
    for row in mapping.to_dict("records"):
        if row["Mapping ID"] != mapping_id(*(row[c] for c in KEY_COLUMNS)):
            raise ValueError("Mapping source identity was edited; rebuild the mapping instead")
        if row["Category"] and row["Category"] not in CATEGORIES:
            raise ValueError(f"Invalid Category {row['Category']!r}")
        if row["Review Status"] not in ("Pending", "Approved", "Excluded"):
            raise ValueError(f"Invalid Review Status {row['Review Status']!r}")
        if row["Review Status"] == "Approved" and not row["Category"]:
            raise ValueError("Approved mappings require one of the allowed categories")
        if row["Review Status"] == "Excluded" and not text(row["Review Notes"]).strip():
            raise ValueError("Excluded mappings require a reason in Review Notes")
        if row["Category"] and not all(text(row[c]).strip() for c in ("Menu", "Menu group", "Item")):
            raise ValueError("Categorized mappings require Menu, Menu group and Item")
    fields = ["Mapping ID", "Category", "Menu", "Menu group", "Item",
              "Review Status", "Evidence", "Review Notes"]
    merged = source.merge(mapping[fields], on="Mapping ID", how="left", validate="many_to_one",
                          indicator=True, sort=False)
    if merged["_merge"].ne("both").any():
        raise ValueError("Source items are missing from the saved Menu Category mapping")
    if not draft and merged["Review Status"].eq("Pending").any():
        raise ValueError("Final report blocked: selected data still has Pending mappings")
    merged["Exception Reason"] = ""
    merged.loc[merged["Category"].eq(""), "Exception Reason"] = "Unresolved mapping"
    excluded = merged["Review Status"].eq("Excluded")
    merged.loc[excluded, "Exception Reason"] = merged.loc[excluded, "Review Notes"]
    return merged.drop(columns="_merge")


def select_window(source: pd.DataFrame, start: str, end: str) -> pd.DataFrame:
    first, last = pd.Timestamp(start), pd.Timestamp(end)
    if first != first.normalize() or last != last.normalize():
        raise ValueError("Start and end must be dates without a time component")
    if first > last:
        raise ValueError("Start date must not be later than end date")
    selected = source[source["Date"].between(first, last)].copy()
    if selected.empty:
        raise ValueError("No source rows in the selected inclusive date window")
    return selected


def place(slicer, top, left, width, height) -> None:
    """Set slicer geometry after creation; keyword arguments to Slicers.Add are ignored."""
    slicer.Top, slicer.Left, slicer.Width, slicer.Height = top, left, width, height


def create_pivot(path: Path, start: str, end: str) -> None:
    import pythoncom
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(path.resolve()))
        sheet = book.Worksheets("PMIX")
        cache = book.PivotCaches().Create(SourceType=1, SourceData="PMIXData", Version=5)
        pivot = cache.CreatePivotTable(
            TableDestination=sheet.Range("A16"), TableName="MenuPMIX", DefaultVersion=5,
        )
        pivot.ManualUpdate = True
        for position, field in enumerate(("Location", "Year", "Month", "Date"), 1):
            pf = pivot.PivotFields(field)
            pf.Orientation, pf.Position = 3, position
            pf.EnableMultiplePageItems = True
        for position, field in enumerate(("Category", "Menu", "Menu group", "Item"), 1):
            pf = pivot.PivotFields(field)
            pf.Orientation, pf.Position = 1, position
            pf.Subtotals = (position < 4,) + (False,) * 11
        pf = pivot.PivotFields("Period")
        pf.Orientation = 2
        pf.Position = 1
        pivot.AddDataField(pivot.PivotFields("Qty"), "Qty (#)", -4157).NumberFormat = "#,##0"
        percentage = pivot.AddDataField(pivot.PivotFields("Qty"), "Qty (%)", -4157)
        percentage.Calculation = 8
        percentage.NumberFormat = "0.00%"
        pivot.RowAxisLayout(0)
        pivot.RowGrand, pivot.ColumnGrand = True, True
        pivot.DisplayErrorString = True
        pivot.ErrorString = "N/A (zero net Qty)"
        pivot.ManualUpdate = False
        pivot.RefreshTable()
        category = pivot.PivotFields("Category")
        for index in range(1, category.PivotItems().Count + 1):
            category.PivotItems(index).ShowDetail = False
        for name, left, width in (("Category", 650, 200), ("Location", 870, 240),
                                  ("Month", 1130, 160)):
            slicer_cache = book.SlicerCaches.Add2(pivot, name)
            place(slicer_cache.Slicers.Add(sheet, pythoncom.Missing, "Select" + name, name),
                  155, left, width, 260)
        timeline = book.SlicerCaches.Add2(pivot, "Date", "SelectDateWindow", 2)
        place(timeline.Slicers.Add(sheet, pythoncom.Missing, "SaleDateTimeline", "Sale date window"),
              35, 650, 640, 105)
        timeline.TimelineState.SetFilterDateRange(start, end)
        sheet.Columns("A:A").ColumnWidth = 48
        sheet.Columns("B:I").ColumnWidth = 14
        book.Save()
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()


def make_report(source_path: Path, mapping_path: Path, output: Path,
                start: str | None, end: str | None, draft: bool, all_dates: bool,
                native: bool = True) -> dict:
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite an existing report: {output}")
    sheets, source = load_sources(source_path)
    if all_dates:
        if not draft or start or end:
            raise ValueError("--all-dates is only valid for a draft without --start/--end")
        start = source["Date"].min().date().isoformat()
        end = source["Date"].max().date().isoformat()
    if not start or not end:
        raise ValueError("Provide both --start and --end, or --draft --all-dates")
    selected = select_window(source, start, end)
    mapping = pd.read_excel(mapping_path, sheet_name="Menu Category",
                            dtype=object, keep_default_na=False)
    mapped = apply_mapping(selected, mapping, draft)
    exceptions = mapped[mapped["Exception Reason"].ne("")].copy()
    data = mapped[mapped["Exception Reason"].eq("")].copy()
    if data.empty:
        raise ValueError("No categorized rows available for the PMIX PivotTable")
    data["Year"] = data["Date"].dt.year
    data["Period"] = data["Date"].map(lambda d: f"{d.year} Q{d.quarter}")
    data["Month"] = data["Date"].dt.strftime("%Y-%m")
    data["Quarter"] = data["Date"].map(
        lambda d: f"{d.quarter}{('st', 'nd', 'rd', 'th')[d.quarter - 1]} Quarter")
    coverage = mapped.groupby(["POS Sheet", "Location"], dropna=False).agg(
        Rows=("Qty", "size"), Net_Qty=("Qty", "sum"),
        First_Date=("Date", "min"), Last_Date=("Date", "max"),
    ).reset_index()
    by_store = mapped.assign(_exc=mapped["Qty"].where(mapped["Exception Reason"].ne(""), 0)
                             ).groupby(["POS Sheet", "Location"], dropna=False)["_exc"].sum()
    coverage["Exception Qty"] = [int(by_store[(a, b)]) for a, b in zip(coverage["POS Sheet"], coverage["Location"])]
    coverage["Exception %"] = (coverage["Exception Qty"] / coverage["Net_Qty"].where(coverage["Net_Qty"].ne(0))).round(4)
    exc_pct = float(exceptions["Qty"].sum()) / float(mapped["Qty"].sum()) if mapped["Qty"].sum() else 0.0
    report_data = data.groupby(
        ["Location", "Year", "Quarter", "Category", "Menu", "Menu group", "Item"],
        dropna=False, sort=False,
    )["Qty"].sum().rename("Qty sold").reset_index()
    audit = {
        "Status": "DRAFT - mappings pending review" if draft else "Approved mapping report",
        "Inclusive start date": start, "Inclusive end date": end,
        "Source item-sales rows": len(source), "Selected item-sales rows": len(mapped),
        "Selected net Qty": int(mapped["Qty"].sum()),
        "Pivot item-sales rows": len(data), "Pivot net Qty": int(data["Qty"].sum()),
        "Exception rows (not in pivot)": len(exceptions),
        "Exception net Qty": int(exceptions["Qty"].sum()),
        "Source location labels": source["Location"].nunique(),
        "Selected location labels": mapped["Location"].nunique(),
        "Pct denominator": "Filtered classified net Qty, across visible periods/categories/locations",
        "Measure": "Sum of item Qty, not unique customer orders; returns are retained",
        "Location policy": "Source labels retained, except Square 'Fluffy Fluffy Tucker GA' renamed to the Toast label 'Fluffy Fluffy - Tucker' (confirmed by Cheena)",
        "Category spelling": "Promotions is the corrected spelling of source Pormotions",
        "Output type": "Native PivotTable" if native else "Prepared data only - NOT an interactive report",
    }
    if audit["Selected net Qty"] != audit["Pivot net Qty"] + audit["Exception net Qty"]:
        raise AssertionError("Selected quantity does not reconcile to pivot plus exceptions")
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.stem + ".building.xlsx")
    if temporary.exists():
        raise FileExistsError(f"Previous unfinished build exists: {temporary}")
    with pd.ExcelWriter(temporary, engine="xlsxwriter",
                        datetime_format="yyyy-mm-dd hh:mm:ss",
                        engine_kwargs={"options": {"strings_to_formulas": False,
                                                   "strings_to_urls": False}}) as writer:
        ws = writer.book.add_worksheet("PMIX")
        ws.write(0, 0, audit["Status"] if native else "PREPARED DATA ONLY - NOT AN INTERACTIVE REPORT")
        ws.write(1, 0, f"Window: {start} through {end}, inclusive")
        ws.write(2, 0, "Expand + beside Category, Menu and Menu group for item detail.")
        ws.write(3, 0, "Qty (#) = net item quantity; Qty (%) = share of filtered classified quantity.")
        ws.write(4, 0, "EXCEPTIONS (not in the PivotTable, not filtered by slicers): "
                       f"{audit['Exception net Qty']:,} Qty in {audit['Exception rows (not in pivot)']:,} rows "
                       f"= {exc_pct:.2%} of selected Qty. Per-store figures: Location Coverage sheet.")
        ws.write(5, 0, "Unresolved/excluded rows are listed in Exceptions; see Validation for reconciliation.")
        ws.freeze_panes(0, 1)
        fields = ["Location", "Date", "Year", "Period", "Month", "Category", "Menu",
                  "Menu group", "Item", "Qty", "POS Sheet", "Source Row",
                  "Mapping ID", "Review Status"]
        write_table(writer, "PMIXData", data[fields], autofilter=False)
        writer.sheets["PMIXData"].add_table(0, 0, len(data), len(fields) - 1, {
            "name": "PMIXData", "columns": [{"header": c} for c in fields],
        })
        write_table(writer, "Menu Category", mapping)
        write_table(writer, "Report Data", report_data)
        write_table(writer, "Exceptions", exceptions)
        write_table(writer, "Location Coverage", coverage)
        write_table(writer, "Validation", pd.DataFrame(audit.items(), columns=["Check", "Value"]))
        for name, frame in sheets.items():
            copied = frame.copy()
            subset = mapped[mapped["POS Sheet"].eq(name)].set_index("Source Row")
            for col in ("Category", "Review Status", "Exception Reason"):
                copied["Mapped " + col] = [subset.at[r, col] if r in subset.index else ""
                                          for r in range(2, len(copied) + 2)]
            copied["In Report Window"] = [r in subset.index for r in range(2, len(copied) + 2)]
            write_table(writer, name, copied)
    if native:
        create_pivot(temporary, start, end)
    temporary.rename(output)
    return audit


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=ROOT / "USA PMix.xlsx")
    commands = parser.add_subparsers(dest="command", required=True)
    build = commands.add_parser("map", help="Build a new editable Menu Category master")
    build.add_argument("--output", type=Path, required=True)
    for command, help_text in (
        ("report", "Read saved mapping and build native Excel pivot"),
        ("prepare", "Prepare auditable pivot source data WITHOUT an interactive report"),
    ):
        report = commands.add_parser(command, help=help_text)
        report.add_argument("--mapping", type=Path, required=True)
        report.add_argument("--output", type=Path, required=True)
        report.add_argument("--start", help="Inclusive date YYYY-MM-DD")
        report.add_argument("--end", help="Inclusive date YYYY-MM-DD")
        report.add_argument("--draft", action="store_true", help="Use pending suggestions, explicitly marked DRAFT")
        report.add_argument("--all-dates", action="store_true", help="Explicitly use all supplied dates for a draft")
    args = parser.parse_args()
    if args.command == "map":
        result = make_mapping(args.source, args.output)
    else:
        result = make_report(args.source, args.mapping, args.output, args.start,
                             args.end, args.draft, args.all_dates, native=args.command == "report")
    print(json.dumps(result, indent=2, ensure_ascii=True))


if __name__ == "__main__":
    main()
