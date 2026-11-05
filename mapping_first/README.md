# Mapping-first PMIX workflow

Two stages, so that a category decision is made once and reused:

1. `map` builds an editable Menu Category master covering every TOAST and SQUARE item, with item
   names cleaned against Cheena's Menu Items List.
2. `report` reads only that saved master and builds a native, filterable Excel PMIX report.

Finished output goes in [`../Result/`](../Result/). This folder holds code, docs and the two
decision files (`mapping_overrides.json`, `menu_name_overrides.json`). Earlier outputs and retired
files are kept in `../versions/`.

## Current status

- The report is a **draft**. Mapping entries stay `Pending` until the business marks them
  `Approved`, and a final build refuses Pending entries.
- Categories and tests reflect the business decisions of 2026-10-31.
- Workbooks are not in Git, because this repository is public and they hold store sales data.

## Categories

Pancakes; Fluffy All Day; Drinks; Japanese Eats; Cheesecakes; Pastries; Promotions; Add-Ons;
Sandwiches; Appetizers; Miscellaneous.

`Promotions` deliberately corrects `Pormotions` in the reference. Nothing is forced into a
category to reach zero unmapped: unknown items stay blank and appear in `Exceptions`.

## How a category is decided

`propose_mapping` applies these in order, and the last match wins:

1. The reference outline in `USA PMix.xlsx`, then the draft keyword rules in `item_hint`.
2. `mapping_overrides.json`: audit decisions per entry, for items a rule cannot recognise, such
   as misspelled free text. It is keyed by Mapping ID and holds no sales figures.
3. `business_rule`: the decisions confirmed on 2026-10-31.

| Decision | Applies to |
|---|---|
| Sandwiches | Sandwiches, savory croffles and filled croissants that would be Fluffy All Day |
| Promotions | Any product sold through a promotion source group |
| Pastries | Cakes other than cheesecakes, and Mango Pudding |
| Add-Ons | Ice cream, single pancakes, and chicken sides |
| Appetizers | Everything in an Appetizer or Appetizers source group |
| Miscellaneous | Non-menu entries such as gift cards, utensil requests and tests |
| Blank, `Excluded` | Custom Amount, on purpose, with the reason in Review Notes |

Chicken sides are Spicy Yuzu Chicken Bites, wings, tenders, strips, and chicken in a side source
group. Bowls, dons, plates, chicken feasts, shareable portions, gyoza and Karaage Chicken stay in
Japanese Eats.

When a rule changes an entry's category, the entry takes the new category as its Menu, so it does
not keep the path of the category it left. Sandwiches also keep their reference Menu group.

## Cleaning item names with the Menu Items List

`../Menu Items List.xlsx` is Cheena's cleaned menu: one column, with bold cells as section
headings (PANCAKES, DRINKS - TEA, and so on). `map` compares every entry with it and fills two
columns, so the master can be filtered instead of checked row by row.

| Menu List Status | Meaning | `Item` |
|---|---|---|
| `Listed` | The source name is already on the list exactly as written | unchanged |
| `Listed - name cleaned` | On the list apart from case, accents, punctuation, `&` or spacing | list spelling |
| `Close match - review` | A typo or variant judged to be the same product (`menu_name_overrides.json`) | list spelling |
| `Listed in UNCATEGORIZED only` | Only in the list's UNCATEGORIZED section, which is unsorted and often misspelled | unchanged |
| `Not in Menu Items List` | No counterpart on the list. `Menu List Match` may say `Possible match, not applied: ...` | unchanged |

`Menu List Match` shows the matched name and its section(s). Only the `Item` column is ever
respelled; `Source Item` and the Mapping ID stay as exported, and whitespace is always trimmed.
Nothing here changes a Category and nothing is marked Approved. The first two statuses are
exact-spelling matches. Items that are not on the list are noted and never guessed at, as Cheena
asked; most are free text, typos or POS variants such as `(Cold or Hot)`.

`menu_name_overrides.json` is the place for a typo or variant that a person has judged to be the
same product as a list name. It is empty today, so no `Close match - review` rows exist yet. An
entry is reversible by deleting its line. Each line holds the Mapping ID, the source name, the
list name and the reason, and no sales figures. A rebuild fails loudly if a listed override name
has since been removed from the list.

The list's own spellings are copied as they are, including its typos (for example `Sapporro`,
`Franch Sparkling Wine`, `Creme Burlee Sauce`). Fix the typo in `Menu Items List.xlsx` and
rebuild if the report should show the corrected name.

## Inputs and outputs

- Input: `../USA PMix.xlsx`, sheets `TOAST`, `SQUARE` and `Menu Category`. It is never modified.
- Input: `../Menu Items List.xlsx`, sheet `USA Menu List`, used by `map` to clean item names.
- `USA PMix - Menu Mapping.xlsx`: `Menu Category` is the editable master, with one row for each
  POS sheet, source group and source item. `Allowed Categories` feeds the Category dropdown, and
  `Menu Items List` is a copy of the list the names were checked against.
- `USA PMix - PMIX Report DRAFT.xlsx`:
  - `PMIX` is a native PivotTable. Rows expand as Category > Menu > Menu group > Item, largest
    first. Columns show `Qty (#)` and `Qty (%)` by year-aware quarter.
  - Category, Location, Month and Week starting slicers filter the table. Nothing selected in a
    slicer means everything, so no Location selected shows all locations. There are no page fields
    and no quarter total columns, because Excel writes those labels in its own interface
    language. The Total row at the bottom stays.
  - `By Location` is a second PivotTable with every item in every location, plus an All
    locations column and row. The same slicers drive both sheets.
  - `PMIXData` feeds the pivot. `Menu Category` is a copy of the exact mapping used.
  - `TOAST` and `SQUARE` keep every input row and column, with mapping status appended.
  - `Exceptions` lists rows outside the pivot. `Location Coverage` gives per-store exception
    Qty and %. `Validation` reconciles pivot Qty plus exception Qty to total Qty.

## Commands

Run from the project root in PowerShell with the Anaconda interpreter, which has pandas,
openpyxl, XlsxWriter and pywin32. The report and `verify_result.py` need Excel for Windows.
Outputs are never overwritten, so remove or rename an old output before rebuilding it.

```powershell
& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py map `
  --output 'Result\USA PMix - Menu Mapping.xlsx'   # --menu-list <file> uses another list

& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py report `
  --mapping 'Result\USA PMix - Menu Mapping.xlsx' `
  --output 'Result\USA PMix - PMIX Report DRAFT.xlsx' `
  --draft --all-dates
```

`prepare` takes the same arguments as `report` and writes the data sheets without a PivotTable.

After review, edit Category, Menu, Menu group, Item and Review Notes in the master. Filter
`Menu List Status` for `Not in Menu Items List` to see the items still to be cleaned. Mark
supported rows `Approved`. Mark genuine exclusions `Excluded` with a reason. Do not edit Mapping
ID or the source identity columns. Hand edits live only in that saved file, so a new `map` run
does not include them. Changing the master affects the **next report build**; Excel Refresh
alone does not rerun the mapping.

For an approved report, drop `--draft --all-dates` and give `--start` and `--end`.

## Measures and filters

- `Qty (#)` is the sum of item quantity, including returns and zero-quantity lines. It is not an
  order count, because the input has no order or check ID.
- `Qty (%)` is % of grand total, recalculated for the filtered rows in the pivot. Rows in
  `Exceptions` are outside this denominator. Miscellaneous is a category, so it is inside it.
- The exception notice on `PMIX` is a fixed total and does not follow the slicers.
- Quarters are calendar quarters labeled with the year. Weeks start on Monday and are labeled
  with that date, for example 2026-08-03.
- Language: every label the report sets is English, and `verify_result.py` scans both report
  sheets for Chinese characters. The date timeline was replaced by the Week starting slicer,
  because Excel draws a timeline's month names and date range in the viewer's Windows language,
  which a file cannot override. Excel's own menus, tooltips and filter pop-ups still follow the
  viewer's Office language.
- Source location labels are kept, except that Square `Fluffy Fluffy Tucker GA` is renamed to the
  Toast label `Fluffy Fluffy - Tucker`, confirmed as the same store on 2026-10-31. The report
  therefore has 19 locations from 20 source labels. The `TOAST` and `SQUARE` sheets keep the
  original label.
- Excel treats labels that differ only by case as one pivot item, for example `Oreo Frappe` and
  `OREO Frappe`. Totals are unaffected; one spelling is displayed.
- No sales or revenue figures are invented, because those inputs are absent.

## Tests and checks

```powershell
& 'D:\Anaconda\python.exe' mapping_first\test_pmix_workflow.py
& 'D:\Anaconda\python.exe' mapping_first\validate_prepared.py
& 'D:\Anaconda\python.exe' mapping_first\verify_result.py
```

- `test_pmix_workflow.py` covers the mapping rules, menu-name cleaning and workbook helpers.
- `validate_prepared.py` compares every original cell and reconciles rows, Qty and all report keys.
- `verify_result.py` opens a temporary copy in live Excel, drives the real slicers,
  checks each result against an independent recompute, and writes
  `Result/VALIDATION_RESULTS.json`.
- Both checks read the date window from the report's own `Validation` sheet, so they work for the
  all-dates draft and for a dated final report. Use `--report`, `--mapping` and `--source` to
  point them at other files.

These prove the arithmetic and that the rules were applied. They do not prove a category is what
the business wants; only review can do that.
