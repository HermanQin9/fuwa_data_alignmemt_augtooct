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
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd


ROOT = Path(__file__).resolve().parent.parent
CATEGORIES = [
    "Pancakes", "Fluffy All Day", "Drinks", "Japanese Eats",
    "Cheesecakes", "Pastries", "Promotions", "Add-Ons",
    # Added by business decision, 2026-10-31.
    "Sandwiches", "Appetizers", "Miscellaneous",
]
# Business decision, 2026-10-31: the Square Tucker label is the same store as the Toast one.
LOCATION_ALIASES = {"Fluffy Fluffy Tucker GA": "Fluffy Fluffy - Tucker"}
OVERRIDES_FILE = Path(__file__).resolve().parent / "mapping_overrides.json"
NAME_OVERRIDES_FILE = Path(__file__).resolve().parent / "menu_name_overrides.json"
MENU_LIST_FILE = ROOT / "Menu Items List.xlsx"
MENU_LIST_SHEET = "USA Menu List"
# Cheena's own pile of names she has not placed or cleaned yet; never used as a clean spelling.
UNSORTED_SECTION = "UNCATEGORIZED"
DECISION_EVIDENCE = "Business decision 2026-10-31"
AUDIT_EVIDENCE = "Audit proposal"
SOURCE_COLUMNS = {
    "TOAST": ("Menu Group", "Menu Item", "Sent Date"),
    "SQUARE": ("Category", "Item", "Date"),
}
KEY_COLUMNS = ["POS Sheet", "Source Group", "Source Item"]
MAPPING_COLUMNS = [
    "Mapping ID", *KEY_COLUMNS, "Category", "Menu", "Menu group", "Item",
    "Review Status", "Evidence", "Candidate Paths", "Rows", "Net Qty",
    "Locations", "Review Notes", "Menu List Status", "Menu List Match",
]
# Menu List Status values; Cheena filters on these.
LISTED = "Listed"
LISTED_CLEANED = "Listed - name cleaned"
CLOSE_MATCH = "Close match - review"
UNSORTED_ONLY = "Listed in UNCATEGORIZED only"
NOT_LISTED = "Not in Menu Items List"
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


def tidy(value: object) -> str:
    """Trim a name and collapse runs of whitespace, so 'X ' and 'X' are one pivot item."""
    return re.sub(r"\s+", " ", text(value)).strip()


def read_menu_list(path: Path) -> list[tuple[str, str]]:
    """(section, name) pairs from the Menu Items List. Bold cells are section headings."""
    book = openpyxl.load_workbook(path, data_only=True)
    try:
        section, entries = "", []
        for (cell,) in book[MENU_LIST_SHEET].iter_rows(min_col=1, max_col=1):
            name = tidy(cell.value)
            if not name:
                continue
            if cell.font.b:
                section = name
            else:
                entries.append((section, name))
    finally:
        book.close()
    if not entries:
        raise ValueError(f"No menu names found in {path}")
    return entries


class MenuList:
    """The menu names, indexed by normalized spelling.

    Names under UNCATEGORIZED are kept apart: they are unsorted and often misspelled, so they
    are evidence that an item is known, not a spelling to copy.
    """

    def __init__(self, entries: list[tuple[str, str]]) -> None:
        self.entries = entries
        self.names: dict[str, tuple[str, list[str]]] = {}
        self.unsorted: dict[str, str] = {}
        for section, name in entries:
            key = normalized(name)
            if not key:
                continue
            if section == UNSORTED_SECTION:
                self.unsorted.setdefault(key, name)
            elif key not in self.names:
                self.names[key] = (name, [section])
            elif section not in self.names[key][1]:
                self.names[key][1].append(section)

    def find(self, *spellings: object) -> tuple[str, list[str]] | None:
        for spelling in spellings:
            hit = self.names.get(normalized(spelling))
            if hit:
                return hit
        return None

    def find_unsorted(self, *spellings: object) -> str | None:
        for spelling in spellings:
            hit = self.unsorted.get(normalized(spelling))
            if hit:
                return hit
        return None

    def label(self, name: str) -> str:
        """'Name [SECTION / SECTION]' for a name that is on the list."""
        listed, sections = self.names[normalized(name)]
        return f"{listed} [{' / '.join(sections)}]"


def load_name_overrides() -> dict:
    """Judged close matches keyed by Mapping ID; holds menu names only, no sales quantities."""
    with NAME_OVERRIDES_FILE.open(encoding="utf-8") as handle:
        return json.load(handle)


def check_menu_name(menu_list: MenuList, name_overrides: dict, ident: str,
                    source_item: object, name: object) -> tuple[str, str, str]:
    """Compare one entry with the Menu Items List. Returns (clean name, status, match).

    Only a name that is on the list after ignoring case, accents, punctuation and '&' is
    respelled automatically. A judged close match is respelled too but flagged for review.
    Anything else keeps its own name and is noted as not listed.
    """
    name = tidy(name)
    hit = menu_list.find(source_item, name)
    if hit:
        listed = hit[0]
        # "Listed" means the source text already was the list spelling, extra spaces included.
        status = LISTED if listed == text(source_item) else LISTED_CLEANED
        return listed, status, menu_list.label(listed)
    pick = name_overrides.get(ident)
    if pick:
        if normalized(pick["list_name"]) not in menu_list.names:
            raise ValueError(f"menu_name_overrides.json: {pick['list_name']!r} is no longer in the "
                             f"Menu Items List (entry {text(source_item)!r}); update the file")
        label = menu_list.label(pick["list_name"])
        if pick["action"] == "rename":
            return menu_list.names[normalized(pick["list_name"])][0], CLOSE_MATCH, label
        return name, NOT_LISTED, f"Possible match, not applied: {label}"
    unsorted = menu_list.find_unsorted(source_item, name)
    if unsorted:
        return name, UNSORTED_ONLY, unsorted
    return name, NOT_LISTED, ""


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
            "Source Location": df["Location"].map(lambda v: text(v).strip()),
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


PROMO_GROUP = re.compile(
    r"\bfree\b|\bpromo|\bcombos?\b|\bbogo\b|\bbundle\b|\bduo\b|anniversary|celebration"
    r"|meal box|\bbuy\b|\+|\$\s*\d|tasty tuesday|thursday", re.I)
NON_MENU = re.compile(r"\b(gift card|utensils?|condiments?|testing printer|donation|merchandise)\b")
ICE_CREAM_GROUPS = {"ice cream", "soft serve", "double scoops ice cream"}
ICE_CREAM_ITEM = re.compile(r"\bice ?cr|\bicecream|\bicre cream|\bsoft serve\b|^(single|double) scoop$")
NOT_ICE_CREAM = re.compile(r"brownie|pancake|combo|cheesecake")
SINGLE_PANCAKE = re.compile(r"\b(single|just one)\b.*\b(pancake|souffle|souflle)\b|\bsingle stack\b"
                            r"|^1( plain)? pancake\b")
SIDE_GROUPS = {"small bites", "sides", "side order", "side orders", "a la cart", "add ons"}
CHICKEN_SIDE = re.compile(r"\bwings?\b|\btenders?\b|\bchicken strips?\b|\bstrips chicken\b"
                          r"|\bchicken bites\b")
CAKE = re.compile(r"\bcakes?\b|\bcake ?pops?\b|\bcakepops?\b|\brollcakes?\b")
NOT_CAKE = re.compile(r"cheese ?cake|basque|pancake|latte|frappe|\btea\b|shake|coffee")
SANDWICH_GROUPS = {"savory croffle", "savory croffles", "crunchy croissant"}
CUSTOM_AMOUNT_NOTE = "Custom Amount intentionally left blank by business decision, 2026-10-31."


def business_rule(group: object, item: object, category: str) -> tuple[str, str]:
    """Apply decisions confirmed on 2026-10-31. Returns (category, reason); no match has no reason."""
    i, g = normalized(item), normalized(group)
    if i == "custom amount":
        return "", CUSTOM_AMOUNT_NOTE
    if NON_MENU.search(f"{g} {i}"):
        return "Miscellaneous", "Non-menu items go in Miscellaneous."
    if PROMO_GROUP.search(text(group)):
        return "Promotions", "Products sold through a promotion group go in Promotions."
    if re.search(r"\bmango pud", i):
        return "Pastries", "Mango Pudding goes in Pastries."
    if category == "Promotions":
        return category, ""
    if g in ("appetizer", "appetizers"):
        return "Appetizers", "Items in an appetizer source group go in Appetizers."
    if g in ICE_CREAM_GROUPS or (ICE_CREAM_ITEM.search(i) and not NOT_ICE_CREAM.search(i)):
        return "Add-Ons", "Ice cream goes in Add-Ons."
    if SINGLE_PANCAKE.search(i):
        return "Add-Ons", "A single pancake goes in Add-Ons."
    if CHICKEN_SIDE.search(i) or (g in SIDE_GROUPS and "chicken" in i and "gyoza" not in i):
        return "Add-Ons", "Chicken sides go in Add-Ons."
    if CAKE.search(i) and not NOT_CAKE.search(i):
        return "Pastries", "Cakes other than cheesecakes go in Pastries."
    if category == "Fluffy All Day":
        if "sandwich" in i:
            return "Sandwiches", "Sandwiches go in Sandwiches."
        if "croffle" in i or g in ("savory croffle", "savory croffles"):
            return "Sandwiches", "Savory croffles go in Sandwiches."
        if g == "crunchy croissant":
            return "Sandwiches", "Filled croissants are treated as sandwiches."
    return category, ""


def load_overrides() -> dict:
    """Per-entry audit decisions keyed by Mapping ID; holds no sales quantities."""
    with OVERRIDES_FILE.open(encoding="utf-8") as handle:
        return json.load(handle)


def propose_mapping(source: pd.DataFrame, reference: pd.DataFrame,
                    menu_list: MenuList | None = None,
                    name_overrides: dict | None = None) -> pd.DataFrame:
    overrides = load_overrides()
    name_overrides = name_overrides or {}
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
        notes = ("Suggested mapping, not approved. Review category and menu path."
                 if category else "Insufficient evidence; do not force a category.")
        draft_category, final = category, category
        override = overrides.get(ident)
        if override:
            final, evidence = override["category"], AUDIT_EVIDENCE
            notes = f"{override['reason']} Draft was {draft_category or 'blank'}. Not approved."
        ruled, reason = business_rule(group, item, final)
        status = "Pending"
        if reason == CUSTOM_AMOUNT_NOTE:
            status, evidence, notes = "Excluded", DECISION_EVIDENCE, reason
        elif reason and ruled != draft_category:
            evidence = DECISION_EVIDENCE
            notes = f"{reason} Draft was {draft_category or 'blank'}. Not approved."
        final = ruled
        if final == draft_category or not final:
            menu = selected["Menu"] if selected else final
            menu_group = selected["Menu group"] if selected else (text(group).strip() or final)
            name = selected["Item"] if selected else item
        else:
            # A changed category must not keep the old category's reference path.
            own_group = text(group).strip()
            own_group = final if own_group in ("", "None") else own_group
            menu = final
            menu_group = selected["Menu group"] if selected and final == "Sandwiches" else own_group
            name = selected["Item"] if selected else item
        list_status = list_match = ""
        if menu_list is not None:
            name, list_status, list_match = check_menu_name(
                menu_list, name_overrides, ident, item, name)
        rows.append({
            "Mapping ID": ident, "POS Sheet": sheet, "Source Group": group,
            "Source Item": item, "Category": final, "Menu": menu,
            "Menu group": menu_group, "Item": name,
            "Review Status": status, "Evidence": evidence,
            "Candidate Paths": "\n".join(paths),
            "Rows": len(frame), "Net Qty": frame["Qty"].sum(),
            "Locations": "; ".join(sorted(frame["Location"].unique())),
            "Review Notes": notes,
            "Menu List Status": list_status, "Menu List Match": list_match,
        })
    return pd.DataFrame(rows, columns=MAPPING_COLUMNS)


def write_table(writer, name: str, frame: pd.DataFrame, autofilter: bool = True) -> None:
    frame.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    ws.freeze_panes(1, 0)
    if autofilter:
        ws.autofilter(0, 0, len(frame), len(frame.columns) - 1)
    ws.set_column(0, len(frame.columns) - 1, 22)
    for col in ("Source Item", "Item", "Candidate Paths", "Review Notes", "Locations",
                "Menu List Status", "Menu List Match", "Menu Item"):
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


def make_mapping(source_path: Path, output: Path, menu_list_path: Path = MENU_LIST_FILE) -> dict:
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite an editable mapping: {output}")
    if not menu_list_path.exists():
        raise FileNotFoundError(f"Menu Items List not found: {menu_list_path}")
    _, source = load_sources(source_path)
    reference = read_reference(source_path)
    menu_entries = read_menu_list(menu_list_path)
    mapping = propose_mapping(source, reference, MenuList(menu_entries), load_name_overrides())
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
        write_table(writer, "Menu Items List",
                    pd.DataFrame(menu_entries, columns=["Section", "Menu Item"]))
        ws = writer.sheets["Menu Category"]
        ws.data_validation(1, 4, len(mapping), 4, {
            "validate": "list", "source": f"='Allowed Categories'!$A$2:$A${len(CATEGORIES) + 1}",
        })
        ws.data_validation(1, 8, len(mapping), 8, {
            "validate": "list", "source": ["Pending", "Approved", "Excluded"],
        })
    return {"mapping_rows": len(mapping), "source_rows": len(source),
            "net_qty": int(source["Qty"].sum()),
            "menu_list_status": mapping["Menu List Status"].value_counts().to_dict(),
            "suggested_rows": int(mapping["Category"].ne("").sum()),
            "excluded_rows": int(mapping["Review Status"].eq("Excluded").sum()),
            "unresolved_rows": int((mapping["Category"].eq("")
                                    & mapping["Review Status"].ne("Excluded")).sum())}


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


def place(slicer, caption, top, left, width, height) -> None:
    """Set a slicer's caption and geometry, and stop it moving when columns resize."""
    slicer.Caption = caption
    slicer.Top, slicer.Left, slicer.Width, slicer.Height = top, left, width, height
    slicer.Shape.Placement = 3  # xlFreeFloating


ROW_HEADER = "Category > Menu > Group > Item"
BY_LOCATION = "By Location"


def add_hierarchy_rows(pivot) -> None:
    for position, field in enumerate(("Category", "Menu", "Menu group", "Item"), 1):
        pf = pivot.PivotFields(field)
        pf.Orientation, pf.Position = 1, position
        pf.Subtotals = (position < 4,) + (False,) * 11


def sort_largest_first(pivot) -> None:
    for field in ("Category", "Menu", "Menu group", "Item"):
        pivot.PivotFields(field).AutoSort(2, "Qty (#)")  # largest first, as in the reference
    category = pivot.PivotFields("Category")
    for index in range(1, category.PivotItems().Count + 1):
        category.PivotItems(index).ShowDetail = False


def create_pivot(path: Path) -> None:
    """Build the two PivotTables and the slicers that drive both.

    Slicers filter fields that are not in the table, so the PivotTables carry no page fields.
    Page fields would show Excel's own "(All)" label, which Excel writes in its interface language.
    For the same reason the quarter totals column is switched off: with two measures Excel names
    it "Qty (#)" plus a word in its interface language, and that name cannot be set.
    """
    import pythoncom
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(path.resolve()))
        sheet, by_location = book.Worksheets("PMIX"), book.Worksheets(BY_LOCATION)
        cache = book.PivotCaches().Create(SourceType=1, SourceData="PMIXData", Version=5)

        pivot = cache.CreatePivotTable(
            TableDestination=sheet.Range("A9"), TableName="MenuPMIX", DefaultVersion=5,
        )
        pivot.ManualUpdate = True
        pivot.HasAutoFormat = False  # keep column widths, so slicers never cover the table
        pivot.CompactLayoutRowHeader = ROW_HEADER
        pivot.CompactLayoutColumnHeader = "Period"
        pivot.GrandTotalName = "Total"
        add_hierarchy_rows(pivot)
        pf = pivot.PivotFields("Period")
        pf.Orientation = 2
        pf.Position = 1
        pivot.AddDataField(pivot.PivotFields("Qty"), "Qty (#)", -4157).NumberFormat = "#,##0"
        percentage = pivot.AddDataField(pivot.PivotFields("Qty"), "Qty (%)", -4157)
        percentage.Calculation = 8
        percentage.NumberFormat = "0.00%"
        pivot.RowAxisLayout(0)
        pivot.RowGrand, pivot.ColumnGrand = False, True
        pivot.DisplayErrorString = True
        pivot.ErrorString = "N/A (zero net Qty)"
        pivot.ManualUpdate = False
        pivot.RefreshTable()
        sort_largest_first(pivot)

        # Every item in every location, with an All locations total.
        grid = cache.CreatePivotTable(
            TableDestination=by_location.Range("A17"), TableName="MenuByLocation", DefaultVersion=5,
        )
        grid.ManualUpdate = True
        grid.HasAutoFormat = False
        grid.CompactLayoutRowHeader = ROW_HEADER
        grid.CompactLayoutColumnHeader = "Location"
        grid.GrandTotalName = "All locations"
        add_hierarchy_rows(grid)
        pf = grid.PivotFields("Location")
        pf.Orientation = 2
        pf.Position = 1
        grid.AddDataField(grid.PivotFields("Qty"), "Qty (#)", -4157).NumberFormat = "#,##0"
        grid.RowAxisLayout(0)
        grid.RowGrand, grid.ColumnGrand = True, True
        grid.DisplayErrorString = True
        grid.ErrorString = "N/A (zero net Qty)"
        grid.ManualUpdate = False
        grid.RefreshTable()
        sort_largest_first(grid)

        sheet.Columns("A:A").ColumnWidth = 48
        sheet.Columns("B:H").ColumnWidth = 14
        by_location.Columns("A:A").ColumnWidth = 48
        by_location.Columns("B:W").ColumnWidth = 16
        by_location.Rows(grid.TableRange1.Row + 1).WrapText = True
        origin = sheet.Range("I1").Left

        # Every control is a plain slicer whose buttons carry only digits and English names.
        controls = {}
        for name, caption, left, width in (
            ("Category", "Category", 0, 200), ("Location", "Location", 220, 240),
            ("Month", "Month", 480, 120), ("Week starting", "Week starting", 620, 150),
        ):
            slicer_cache = book.SlicerCaches.Add2(pivot, name)
            place(slicer_cache.Slicers.Add(sheet, pythoncom.Missing, "Select" + name, caption),
                  caption, 35, origin + left, width, 420)
            controls[name] = slicer_cache

        # The same controls also drive the By Location table, and appear on its sheet.
        for slicer_cache in controls.values():
            slicer_cache.PivotTables.AddPivotTable(grid)
        for name, left, width in (("Category", 10, 200), ("Month", 220, 120), ("Week starting", 350, 150)):
            place(controls[name].Slicers.Add(by_location, pythoncom.Missing, name + " here", name),
                  name, 75, left, width, 140)
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
    # Monday of each sale week, written as plain digits so every viewer sees the same labels.
    data["Week starting"] = (data["Date"] - pd.to_timedelta(data["Date"].dt.weekday, unit="D")
                             ).dt.strftime("%Y-%m-%d")
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
        "Source location labels": source["Source Location"].nunique(),
        "Report locations (after Tucker rename)": source["Location"].nunique(),
        "Selected report locations": mapped["Location"].nunique(),
        "Pct denominator": "Filtered classified net Qty, across visible periods/categories/locations",
        "Measure": "Sum of item Qty, not unique customer orders; returns are retained",
        "Location policy": "Source labels retained, except Square 'Fluffy Fluffy Tucker GA' renamed to the Toast label 'Fluffy Fluffy - Tucker' (business decision)",
        "Category spelling": "Promotions is the corrected spelling of source Pormotions",
        "Output type": "Native PivotTable" if native else "Prepared data only - NOT an interactive report",
    }
    if audit["Selected net Qty"] != audit["Pivot net Qty"] + audit["Exception net Qty"]:
        raise AssertionError("Selected quantity does not reconcile to pivot plus exceptions")
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.stem + ".building.xlsx")
    if temporary.exists():
        raise FileExistsError(f"Previous unfinished build exists: {temporary}")
    try:
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
            ws.write(6, 0, "Location: nothing selected means ALL locations. Use the funnel icon on a slicer "
                           "to clear it. The By Location sheet shows every item in every location.")
            ws.freeze_panes(0, 1)
            if native:
                grid = writer.book.add_worksheet(BY_LOCATION)
                grid.write(0, 0, audit["Status"])
                grid.write(1, 0, f"Window: {start} through {end}, inclusive")
                grid.write(2, 0, "Every item in every location. The All locations column and row add up every store shown.")
                grid.write(3, 0, "Filter with the Category, Month and Week starting slicers here, or the Location slicer on PMIX.")
                grid.freeze_panes(0, 1)
                grid.write(4, 0, "Qty (#) = net item quantity. Rows listed in Exceptions are not included.")
            fields = ["Location", "Date", "Year", "Period", "Month", "Week starting", "Category", "Menu",
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
            create_pivot(temporary)
        temporary.rename(output)
    except BaseException:
        temporary.unlink(missing_ok=True)  # a failed build must not block the next run
        raise
    return audit


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=ROOT / "USA PMix.xlsx")
    commands = parser.add_subparsers(dest="command", required=True)
    build = commands.add_parser("map", help="Build a new editable Menu Category master")
    build.add_argument("--output", type=Path, required=True)
    build.add_argument("--menu-list", type=Path, default=MENU_LIST_FILE,
                       help="Menu Items List workbook used to clean item names")
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
        result = make_mapping(args.source, args.output, args.menu_list)
    else:
        result = make_report(args.source, args.mapping, args.output, args.start,
                             args.end, args.draft, args.all_dates, native=args.command == "report")
    print(json.dumps(result, indent=2, ensure_ascii=True))


if __name__ == "__main__":
    main()
