"""Record a reviewer's recategorizing, renaming and approvals so every rebuild keeps them.

Decisions are saved by Mapping ID in mapping_first/review_decisions.json, which the mapping
applies last, after every rule. Three ways in:

  template     Write a fill-in workbook: one row per item, with New Section, New Item Name and
               New Status columns to fill.
  apply        Read that workbook back.
  from-report  Read a reviewed report the way Cheena marked one on 2026-11-06: a status or a
               target section typed in PMIXData's Review Status column (and, since 2026-11-09,
               in the Exceptions sheet's Review Status column), and CATEGORY / REMARKS notes
               typed beside the items on the PMIX sheet.
  list-copy    Write a copy of the Menu Items List with the names in menu_list_additions.json
               added. The original workbook is never changed.

Each run writes a check workbook listing every decision, and anything it could not apply. Then
rebuild with run_pipeline.py, or pass --rebuild.
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import difflib
import json
import re
import subprocess
import sys
from pathlib import Path

import html
import zipfile

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pmix_workflow as wf  # noqa: E402

HERE = Path(__file__).resolve().parent
RESULT = wf.ROOT / "Result"
STATUSES = ("Approved", "Pending")
NEW_COLUMNS = ["New Section", "New Item Name", "New Status", "Note"]
# A pivot note that defers to the row-level status column.
SEE_DATA = re.compile(r"^see pmixdata", re.I)
REMOVE = re.compile(r'^remove\s+"(.+?)"', re.I)
REMOVE_CHAR = re.compile(r"^remove\s+(\S)\s+character", re.I)
RENAME = re.compile(r"^ra?e?name\s+(?:to\s+)?(.+?)(?:,\s*if item is under (.+?) category)?\.?\s*$", re.I)


def source_label(row) -> str:
    return " | ".join(wf.text(row[c]) for c in wf.KEY_COLUMNS)


def sections_list(menu_list_path: Path) -> list[str]:
    return wf.report_sections(wf.MenuList(wf.read_full_menu_list(menu_list_path)))


def instruction(remark: object) -> tuple[str, str, str]:
    """(kind, value, condition) from a REMARKS note. kind is '', 'remove' or 'rename'."""
    note = wf.tidy(remark)
    if not note:
        return "", "", ""
    if match := REMOVE.match(note):
        return "remove", match.group(1), ""
    if match := REMOVE_CHAR.match(note):
        return "remove", match.group(1), ""
    if match := RENAME.match(note):
        return "rename", match.group(1).strip(), match.group(2) or ""
    return "rename", note, ""  # a bare name, e.g. "Ube Pancakes"


def apply_instruction(name: str, kind: str, value: str) -> str | None:
    """The new name, or None when a 'remove' text is not in the name."""
    if kind == "rename":
        return wf.tidy(value)
    if kind == "remove":
        # Spacing-insensitive: 'Remove " -Cold"' also removes " - Cold".
        pattern = r"\s*".join(re.escape(c) for c in value if not c.isspace())
        found = re.search(pattern, name, re.I)
        if not found:
            return None
        # A removal must not leave a doubled word where it cut ('Thai Tea Iced Tea' less ' Iced'
        # is 'Thai Tea'); doubles elsewhere, like 'Fluffy Fluffy Latte', are left alone.
        before, after = wf.tidy(name[:found.start()]).split(), wf.tidy(name[found.end():]).split()
        if before and after and before[-1].casefold() == after[0].casefold():
            after = after[1:]
        return wf.tidy(" ".join(before + after))
    return name


def snap(name: str, menu_list: wf.MenuList) -> str:
    """The list spelling of a name a typo away from it ('Mocha Expresso' -> 'Mocha Espresso')."""
    if menu_list.find(name):
        return menu_list.find(name)[0]
    # Same number of words: a typo, never a dropped word ('Oreo Basque ...' is not 'Basque ...').
    listed = [n for n, _ in menu_list.names.values() if len(n.split()) == len(name.split())]
    close = difflib.get_close_matches(name, listed, n=2, cutoff=0.9)
    if len(close) == 1 or (close and difflib.SequenceMatcher(None, name, close[0]).ratio() >= 0.95):
        return close[0]
    return name


def decision(entry: dict, by: str, section: str, item: str, status: str, note: str,
             original: str = "") -> dict:
    category = wf.section_category(section, wf.text(entry["Category"]))
    made = {"source": source_label(entry), "status": status, "category": category,
            "section": section, "item": item, "by": by, "note": note}
    if original and item != original:
        made["renamed_from"] = original
    return made


def save(decisions: dict, check: dict[str, list], label: str, menu_list: wf.MenuList,
         known: set | None = None) -> Path:
    stored = wf.load_review_decisions()
    replaced = sum(1 for k in decisions if k in stored)
    stored.update(decisions)
    wf.REVIEW_FILE.write_text(json.dumps(dict(sorted(stored.items(), key=lambda kv: kv[1]["source"])),
                                         indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    # Renamed to a name that is neither on the list nor already an item name in the data.
    known = {wf.normalized(n) for n in known or ()}
    unlisted = sorted({d["item"] for d in decisions.values()
                       if d.get("renamed_from") and not menu_list.find(d["item"])
                       and wf.normalized(d["item"]) not in known})
    check["Names not on the list"] = [{"Item": n} for n in unlisted]
    path = RESULT / f"Recategorize Check - {label}.xlsx"
    path = path.with_name(f"{path.stem} {dt.datetime.now():%Y%m%d-%H%M%S}.xlsx")
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        rows = [{"Mapping ID": k, **v} for k, v in decisions.items()]
        wf.write_table(writer, "Decisions", pd.DataFrame(rows))
        for name, items in check.items():
            wf.write_table(writer, name[:31], pd.DataFrame(items or [{"Result": "none"}]))
    counts = pd.Series([d["status"] for d in decisions.values()]).value_counts().to_dict()
    print(json.dumps({"decisions": len(decisions), "replacing_earlier": replaced, "by_status": counts,
                      **{k: len(v) for k, v in check.items()},
                      "saved_to": str(wf.REVIEW_FILE.relative_to(wf.ROOT)),
                      "check_workbook": str(path.relative_to(wf.ROOT))}, indent=2))
    return path


def read_pmix_notes(report: Path, report_sections: set[str]) -> tuple[list[dict], list[str]]:
    """CATEGORY / REMARKS notes typed beside pivot rows: (notes, general remarks)."""
    book = openpyxl.load_workbook(report, read_only=True, data_only=True)
    try:
        rows = [list(r) for r in book["PMIX"].iter_rows(values_only=True)]
    finally:
        book.close()
    header = next(((i, r) for i, r in enumerate(rows)
                   if "CATEGORY" in r and "REMARKS" in r), None)
    if header is None:
        return [], []
    start, cells = header
    cat_col, rem_col = cells.index("CATEGORY"), cells.index("REMARKS")
    notes, general, parent = [], [], None
    for row in rows[start + 1:]:
        label = wf.tidy(row[0])
        if not label or label == "Total":
            continue
        target, remark = wf.tidy(row[cat_col]), wf.tidy(row[rem_col])
        if label in report_sections:
            parent = label
            if remark:
                general.append(remark)
            continue
        if target or remark:
            notes.append({"section": parent, "label": label, "target": target, "remark": remark})
    return notes, general


def read_cell_comments(book_path: Path, sheet: str) -> list[tuple[int, str]]:
    """(Excel row, text) of every comment on a sheet: threaded comments, else legacy notes."""
    with zipfile.ZipFile(book_path) as book:
        names = set(book.namelist())
        workbook = book.read("xl/workbook.xml").decode("utf-8")
        rels = book.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        found = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(html.escape(sheet)), workbook)
        if not found:
            return []
        target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % found.group(1), rels) or \
            re.search(r'Target="([^"]+)"[^>]*Id="%s"' % found.group(1), rels)
        part = target.group(1).split("/")[-1]
        sheet_rels = f"xl/worksheets/_rels/{part}.rels"
        if sheet_rels not in names:
            return []
        links = re.findall(r'Target="([^"]+)"', book.read(sheet_rels).decode("utf-8"))
        threaded = [t for t in links if "threadedComment" in t]
        legacy = [t for t in links if re.search(r"comments?\d*\.xml$", t) and "threaded" not in t]
        comments = []
        for link in threaded or legacy:
            part = link.lstrip("/") if link.startswith("/") else "xl/" + link.replace("../", "")
            xml = book.read(part).decode("utf-8")
            if threaded:
                found = re.findall(r'<threadedComment [^>]*ref="[A-Z]+(\d+)"[^>]*>.*?<text>(.*?)</text>', xml, re.S)
            else:
                found = [(r, re.sub(r"<[^>]+>", "", t).split("Comment:")[-1])
                         for r, t in re.findall(r'<comment ref="[A-Z]+(\d+)"[^>]*>(.*?)</comment>', xml, re.S)]
            comments += [(int(r), html.unescape(t).strip()) for r, t in found]
    return comments


def read_exception_statuses(report: Path) -> dict[str, list[str]]:
    """Review Status typed on the Exceptions sheet, by Mapping ID (sorted distinct values)."""
    try:
        frame = pd.read_excel(report, sheet_name="Exceptions", dtype=object, keep_default_na=False,
                              usecols=["Mapping ID", "Review Status"])
    except ValueError:  # an older report has no such sheet or column
        return {}
    return frame.groupby("Mapping ID")["Review Status"].agg(
        lambda s: sorted({wf.tidy(v) for v in s})).to_dict()


def drop_unchanged(decisions: dict, check: dict[str, list]) -> dict:
    """Keep only decisions that differ from the ones on file, so earlier reviews keep their own
    date and wording when a later report repeats them."""
    stored = wf.load_review_decisions()
    changed, same = {}, []
    for ident, made in decisions.items():
        before = stored.get(ident)
        if before and all(before.get(k, "") == made.get(k, "")
                          for k in ("status", "category", "section", "item")):
            same.append({"Source": made["source"], "Item": made["item"], "Status": made["status"]})
            continue
        if before and before.get("renamed_from") and not made.get("renamed_from"):
            made["renamed_from"] = before["renamed_from"]
        changed[ident] = made
    check["Already on file"] = same
    return changed


def from_report(report: Path, by: str, menu_list_path: Path) -> dict:
    sections = sections_list(menu_list_path)
    menu_list = wf.MenuList(wf.read_full_menu_list(menu_list_path))
    data = pd.read_excel(report, sheet_name="PMIXData", dtype=object, keep_default_na=False,
                         usecols=["Mapping ID", "Category", wf.LIST_SECTION, "Item", "Review Status"])
    mapping = pd.read_excel(report, sheet_name="Menu Category", dtype=object, keep_default_na=False)
    entries = {r["Mapping ID"]: r for r in mapping.to_dict("records")}
    per_entry = data.groupby("Mapping ID").agg(
        status=("Review Status", lambda s: sorted(set(s))), category=("Category", "first"),
        section=(wf.LIST_SECTION, "first"), item=("Item", "first"))
    check = {"Could not apply": [], "Still Pending": [], "General remarks": [], "Inferred": [],
             "Cell comments": [], "Left as is": []}
    decisions, targets = {}, {}
    for ident, row in per_entry.iterrows():
        entry = {**entries[ident], "Category": row["category"]}
        if len(row["status"]) > 1:
            check["Could not apply"].append({"Source": source_label(entry), "Problem":
                                             f"Rows disagree: {row['status']}"})
            continue
        status = row["status"][0]
        if status == "Pending":
            continue
        if status == "Approved":
            decisions[ident] = {"source": source_label(entry), "status": "Approved",
                                "category": wf.section_category(row["section"], row["category"]),
                                "section": row["section"],
                                "item": row["item"], "by": by, "note": f"Approved by {by}."}
            continue
        section = wf.canonical_section(status, sections)
        if section is None:
            check["Could not apply"].append({"Source": source_label(entry),
                                             "Problem": f"Unknown section {status!r}"})
            continue
        targets[ident] = section
        decisions[ident] = decision(entry, by, section, row["item"],
                                    "Pending" if section == wf.UNSORTED_SECTION else "Approved",
                                    f"Moved to {section} by {by}.")

    # The Exceptions sheet holds entries the pivot does not: a section typed there places the entry;
    # UNCATEGORIZED means the reviewer could not place it, so it stays an unresolved exception.
    exception_ids = read_exception_statuses(report)
    for ident, statuses in exception_ids.items():
        entry = entries[ident]
        if len(statuses) > 1:
            check["Could not apply"].append({"Source": source_label(entry),
                                             "Problem": f"Exceptions rows disagree: {statuses}"})
            continue
        status = statuses[0]
        if status in ("", "Pending"):
            continue
        if wf.text(entry["Review Status"]) == "Excluded":
            check["Left as is"].append({"Source": source_label(entry), "Why":
                                        f"Excluded on purpose; {status!r} typed on Exceptions not applied. "
                                        + wf.tidy(entry["Review Notes"])})
            continue
        section = None if status == "Approved" else wf.canonical_section(status, sections)
        if section is None:
            check["Could not apply"].append({"Source": source_label(entry), "Problem":
                                             "Approved needs a section" if status == "Approved"
                                             else f"Unknown section {status!r}"})
            continue
        unplaced = section == wf.UNSORTED_SECTION
        decisions[ident] = decision(entry, by, section, wf.text(entry["Item"]),
                                    "Pending" if unplaced else "Approved",
                                    f"Left in {section} by {by}: could not be placed." if unplaced
                                    else f"Moved to {section} by {by}.")

    notes, general = read_pmix_notes(report, set(per_entry["section"]))
    check["General remarks"] = [{"Remark": g} for g in general]
    by_label = {}
    for ident, row in per_entry.iterrows():
        by_label.setdefault((row["section"], wf.tidy(row["item"]).casefold()), []).append(ident)
    for note in notes:
        idents = by_label.get((note["section"], note["label"].casefold()), [])
        if not idents:
            check["Could not apply"].append({"Source": note["label"], "Problem":
                                             f"No item under {note['section']} with this name"})
            continue
        deferred = bool(SEE_DATA.match(note["target"]))
        kind, value, condition = instruction(note["remark"])
        for ident in idents:
            row, entry = per_entry.loc[ident], {**entries[ident], "Category": per_entry.loc[ident, "category"]}
            section = targets.get(ident) if deferred else wf.canonical_section(note["target"], sections)
            if section is None and deferred and condition and                     wf.singular(condition) == wf.singular(row["category"]):
                # "Rename ..., if item is under Pancakes Category" and this entry's Category is Pancakes.
                section = wf.canonical_section(condition, sections)
            fallback = False
            if section is None and deferred and not condition:
                # "See PMIXData", but that row was left Pending: the entry's own Category names
                # the section (a Pancakes entry goes under PANCAKES), and the name follows the
                # list spelling her other renames of that item use.
                section = wf.canonical_section(row["category"], sections)
                fallback = section is not None
            if section is None:
                problem = ("PMIXData status not set" if deferred
                           else f"Unknown section {note['target']!r}") if note["target"] else "No CATEGORY"
                if not note["target"] and kind:
                    section = row["section"]  # a rename alone keeps the current section
                else:
                    check["Could not apply"].append({"Source": source_label(entry), "Problem": problem})
                    continue
            item = row["item"]
            if kind:
                wanted = wf.canonical_section(condition, sections) if condition else None
                # A deferred note's rename was written for the item's main use, not for an
                # add-on, promotion or combo of the same name.
                skip = (condition and wanted != section) or (
                    deferred and not condition and section in ("ADD-ONS", "PROMOTIONS", "COMBOS"))
                if not skip:
                    renamed = apply_instruction(item, kind, value)
                    if renamed is None:
                        check["Could not apply"].append({"Source": source_label(entry), "Problem":
                                                         f"{note['remark']!r}: text not in {item!r}"})
                    else:
                        item = renamed
            if fallback:
                item = snap(item, menu_list)
                check["Inferred"].append({"Source": source_label(entry), "Section": section, "Item": item,
                                          "Why": "Note says see PMIXData, but that row was left Pending"})
            status = "Pending" if section == wf.UNSORTED_SECTION else "Approved"
            text = f"{section} by {by}" + (f"; renamed from {row['item']!r}" if item != row["item"] else "")
            decisions[ident] = decision(entry, by, section, item, status, text + ".", row["item"])
    # Comments on PMIXData cells, e.g. "Rename to Crème Brulee Pancakes" on that item's rows.
    row_ids = dict(zip(range(2, len(data) + 2), data["Mapping ID"]))
    asked = {}
    for excel_row, comment in read_cell_comments(report, "PMIXData"):
        if excel_row in row_ids:
            asked.setdefault(row_ids[excel_row], set()).add(comment)
    for ident, comments in asked.items():
        row, entry = per_entry.loc[ident], {**entries[ident], "Category": per_entry.loc[ident, "category"]}
        if len(comments) > 1:
            check["Could not apply"].append({"Source": source_label(entry), "Problem": f"Comments differ: {sorted(comments)}"})
            continue
        comment = comments.pop()
        kind, value, _ = instruction(comment)
        made = decisions.get(ident) or decision(entry, by, row["section"], row["item"],
                                                "Pending", f"Commented by {by}.")
        renamed = apply_instruction(made["item"], kind, value) if kind else None
        if not renamed:
            check["Could not apply"].append({"Source": source_label(entry), "Problem": f"Comment {comment!r}"})
            continue
        if renamed != made["item"]:
            made["renamed_from"] = made.get("renamed_from") or row["item"]
            made["item"] = renamed
            made["note"] = made["note"].rstrip(".") + f"; renamed from {row['item']!r} (cell comment)."
        decisions[ident] = made
        check["Cell comments"].append({"Source": source_label(entry), "Comment": comment, "Item": made["item"]})
    # Follow a rename onto another note's rename ('Iced Vanilla Late' -> 'Vanilla Late' -> 'Vanilla
    # Latte'), then snap a near miss to the list's spelling.
    final = {}
    for note in notes:
        kind, value, condition = instruction(note["remark"])
        if kind and not condition:
            renamed = apply_instruction(note["label"], kind, value)
            if renamed and renamed.casefold() != note["label"].casefold():
                final[note["label"].casefold()] = renamed
    check["Renames adjusted"] = []
    for ident, made in decisions.items():
        if not made.get("renamed_from"):
            continue
        name, seen = made["item"], set()
        while name.casefold() in final and name.casefold() not in seen:
            seen.add(name.casefold())
            name = final[name.casefold()]
        name = snap(name, menu_list)
        if name != made["item"]:
            check["Renames adjusted"].append({"Source": made["source"], "Asked": made["item"], "Used": name})
            made["item"] = name
    for ident, row in per_entry.iterrows():
        if ident not in decisions:
            check["Still Pending"].append({"Source": source_label(entries[ident]),
                                           "Section": row["section"], "Item": row["item"]})
    check["Still unplaced (Exceptions)"] = [
        {"Source": source_label(entries[i]), "Item": wf.text(entries[i]["Item"]),
         "Status": wf.text(entries[i]["Review Status"])}
        for i in exception_ids if decisions.get(i, {}).get("status", "Pending") != "Approved"
        and wf.text(entries[i]["Review Status"]) != "Excluded"]
    decisions = drop_unchanged(decisions, check)
    save(decisions, check, "report", menu_list, set(per_entry["item"]))
    return decisions


def list_copy(menu_list_path: Path, output: Path) -> None:
    """Copy the Menu Items List with menu_list_additions.json added at the end of each section."""
    if output.exists():
        raise SystemExit(f"Refusing to overwrite {output}")
    additions = wf.load_menu_additions()
    book = openpyxl.load_workbook(menu_list_path)
    sheet = book[wf.MENU_LIST_SHEET]
    section, last, have = "", {}, {}
    for (cell,) in sheet.iter_rows(min_col=1, max_col=1):
        name = wf.tidy(cell.value)
        if not name:
            continue
        if cell.font.b:
            section = name
            have.setdefault(section, set())
        else:
            have[section].add(wf.normalized(name))
        last[section] = cell.row
    fresh = {}
    for target, names in additions.items():
        if target not in last:
            raise SystemExit(f"menu_list_additions.json: section {target!r} is not on the list")
        fresh[target] = [n for n in names if wf.normalized(n) not in have[target]]
    # Bottom sections first, so the row numbers of the ones above stay valid.
    for target in sorted(fresh, key=lambda s: -last[s]):
        names, after = fresh[target], last[target]
        if not names:
            continue
        sheet.insert_rows(after + 1, amount=len(names))
        model = sheet.cell(after, 1)
        for offset, name in enumerate(names, 1):
            cell = sheet.cell(after + offset, 1, name)
            cell._style = copy.copy(model._style)
            if model.font.b:  # a section with no names yet: the new name is not a heading
                cell.font = openpyxl.styles.Font(name=model.font.name, sz=model.font.sz, b=False)
    output.parent.mkdir(parents=True, exist_ok=True)
    book.save(output)
    print(json.dumps({"added": sum(len(v) for v in fresh.values()), "written_to": str(output)}, indent=2))


def template(mapping_path: Path, output: Path, menu_list_path: Path) -> None:
    if output.exists():
        raise SystemExit(f"Refusing to overwrite {output}")
    sections = sections_list(menu_list_path)
    mapping = pd.read_excel(mapping_path, sheet_name="Menu Category", dtype=object, keep_default_na=False)
    if wf.LIST_SECTION not in mapping:
        mapping[wf.LIST_SECTION] = ""
    frame = mapping[["Mapping ID", *wf.KEY_COLUMNS, "Category", wf.LIST_SECTION, "Item",
                     "Review Status", "Menu List Status"]].copy()
    for column in NEW_COLUMNS:
        frame[column] = ""
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wf.write_table(writer, "Recategorize", frame)
        wf.write_table(writer, "Sections", pd.DataFrame({"Section": sections}))
        ws = writer.sheets["Recategorize"]
        first = len(frame.columns) - len(NEW_COLUMNS)
        ws.data_validation(1, first, len(frame), first, {
            "validate": "list", "source": f"=Sections!$A$2:$A${len(sections) + 1}"})
        ws.data_validation(1, first + 2, len(frame), first + 2, {"validate": "list", "source": list(STATUSES)})
        fill = writer.book.add_format({"bg_color": "#FFF2CC"})
        ws.set_column(first, first + len(NEW_COLUMNS) - 1, 28, fill)
        ws.set_column(0, 0, 12)
    print(f"Template written: {output}\nFill New Section, New Item Name and/or New Status, then run "
          f"recategorize.py apply \"{output}\"")


def apply(path: Path, by: str, menu_list_path: Path) -> dict:
    sections = sections_list(menu_list_path)
    menu_list = wf.MenuList(wf.read_full_menu_list(menu_list_path))
    sheet = pd.read_excel(path, sheet_name="Recategorize", dtype=object, keep_default_na=False)
    check = {"Could not apply": []}
    decisions = {}
    for row in sheet.to_dict("records"):
        new_section, new_item = wf.tidy(row["New Section"]), wf.tidy(row["New Item Name"])
        new_status = wf.tidy(row["New Status"]).capitalize()
        if not (new_section or new_item or new_status):
            continue
        if wf.mapping_id(*(row[c] for c in wf.KEY_COLUMNS)) != row["Mapping ID"]:
            check["Could not apply"].append({"Source": source_label(row), "Problem": "Source columns were edited"})
            continue
        section = wf.canonical_section(new_section, sections) if new_section else wf.text(row[wf.LIST_SECTION])
        if new_section and section is None:
            check["Could not apply"].append({"Source": source_label(row),
                                             "Problem": f"Unknown section {new_section!r}"})
            continue
        if new_status and new_status not in STATUSES:
            check["Could not apply"].append({"Source": source_label(row),
                                             "Problem": f"Status must be Approved or Pending, not {new_status!r}"})
            continue
        status = new_status or wf.text(row["Review Status"])
        if status == "Approved" and not wf.text(row["Category"]) and not section:
            check["Could not apply"].append({"Source": source_label(row), "Problem": "Approved needs a section"})
            continue
        note = wf.tidy(row["Note"]) or f"Reviewed by {by}."
        decisions[row["Mapping ID"]] = decision(row, by, section, new_item or wf.text(row["Item"]), status,
                                                note, wf.text(row["Item"]))
    save(decisions, check, "template", menu_list)
    return decisions


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--by", default=f"Cheena {dt.date.today().isoformat()}",
                        help="Who decided, shown as Evidence (default: Cheena and today's date)")
    parser.add_argument("--menu-list", type=Path, default=wf.MENU_LIST_FILE)
    commands = parser.add_subparsers(dest="command", required=True)
    make = commands.add_parser("template", help="Write a fill-in workbook from a mapping")
    make.add_argument("--mapping", type=Path, default=RESULT / "USA PMix - Menu Mapping.xlsx")
    make.add_argument("--output", type=Path, default=RESULT / "Recategorize Template.xlsx")
    copy_list = commands.add_parser("list-copy", help="Write the Menu Items List with the approved additions")
    copy_list.add_argument("--output", type=Path, default=RESULT / "Menu Items List - with additions.xlsx")
    for name, help_text in (("apply", "Read a filled-in template"),
                            ("from-report", "Read a reviewed report (PMIXData status and PMIX notes)")):
        command = commands.add_parser(name, help=help_text)
        command.add_argument("workbook", type=Path)
        command.add_argument("--rebuild", action="store_true", help="Run run_pipeline.py afterwards")
    args = parser.parse_args()
    if args.command == "template":
        template(args.mapping, args.output, args.menu_list)
        return 0
    if args.command == "list-copy":
        list_copy(args.menu_list, args.output)
        return 0
    if args.command == "apply":
        apply(args.workbook, args.by, args.menu_list)
    else:
        from_report(args.workbook, args.by, args.menu_list)
    if args.rebuild:
        return subprocess.run([sys.executable, str(HERE / "run_pipeline.py"), "--label", "recategorize"],
                              cwd=wf.ROOT).returncode
    print("Next: rebuild with mapping_first\\run_pipeline.py --label recategorize")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
