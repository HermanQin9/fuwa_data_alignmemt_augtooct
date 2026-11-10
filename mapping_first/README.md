# Mapping-first PMIX workflow

Two stages, so that a category decision is made once and reused:

1. `map` builds an editable Menu Category master covering every TOAST and SQUARE item, with item
   names cleaned against Cheena's Menu Items List.
2. `report` reads only that saved master and builds a native, filterable Excel PMIX report.

Finished output goes in [`../Result/`](../Result/). This folder holds code, docs and the four
decision files (`mapping_overrides.json`, `menu_name_overrides.json`, `review_decisions.json`,
`menu_list_additions.json`). Earlier outputs and retired files are kept in `../versions/`.

## Current status

- The report is a **draft**. Mapping entries stay `Pending` until the business marks them
  `Approved`, and a final build refuses Pending entries.
- Categories and tests reflect the business decisions of 2026-10-31, Cheena's answers of
  2026-11-05 and her reviews of 2026-11-06 and 2026-11-09 (`review_decisions.json`): 2,062
  entries Approved, with her sections and item names. On 2026-11-09 she placed the Basque
  cheesecakes, Bellini and Cream, approved T-Shirt, confirmed `Crème Brulee Pancakes` and moved
  Raspberry Caramelized Banana Pancakes (a `FREE ICE CREAM` item) to Promotions. 39 entries
  stay Pending and unplaced in `Exceptions` because she could not place them, and Custom Amount
  stays Excluded. Nothing was forced into a category.
- Workbooks are not in Git, because this repository is public and they hold store sales data.

## Categories

Pancakes; Fluffy All Day; Drinks; Japanese Eats; Cheesecakes; Pastries; Promotions; Add-Ons;
Sandwiches; Appetizers; Miscellaneous; Combos (Cheena, 2026-11-06: combos are not Promotions).

`Promotions` deliberately corrects `Pormotions` in the reference. Nothing is forced into a
category to reach zero unmapped: unknown items stay blank and appear in `Exceptions`.

## How a category is decided

`propose_mapping` applies these in order, and the last match wins:

1. The reference outline in `USA PMix.xlsx`, then the draft keyword rules in `item_hint`.
2. `mapping_overrides.json`: audit decisions per entry, for items a rule cannot recognise, such
   as misspelled free text. It is keyed by Mapping ID and holds no sales figures.
3. `business_rule`: the decisions confirmed on 2026-10-31.
4. `review_decisions.json`: a reviewer's own decision for an entry (status, Menu List section,
   category, item name). It is applied last and wins over every rule. See "Recategorizing".

| Decision | Applies to |
|---|---|
| Sandwiches | Sandwiches, savory croffles and filled croissants that would be Fluffy All Day |
| Promotions | Any product sold through a promotion source group |
| Pastries | Cakes other than cheesecakes, and Mango Pudding |
| Add-Ons | Ice cream, single pancakes, and chicken sides |
| Appetizers | Everything in an Appetizer or Appetizers source group; wings, gyoza, Karaage Chicken and Takoyaki (2026-11-05) |
| Pastries | Affogato, which is a dessert; Affogato Float stays a drink (2026-11-05) |
| Miscellaneous | Non-menu entries such as gift cards, utensil requests and tests |
| Blank, `Excluded` | Custom Amount, on purpose, with the reason in Review Notes |

Chicken sides are Spicy Yuzu Chicken Bites, tenders, strips, and chicken in a side source group.
On 2026-11-05 Cheena confirmed that wings, gyoza and Karaage Chicken follow the APPETIZER section of
her Menu Items List instead, so they go in Appetizers, as does the free-text typo `Karrage
Chicken`. Takoyaki, also under APPETIZER on her list, follows it too (Herman, 2026-11-05). Bowls, dons, plates, chicken feasts and
shareable portions stay in Japanese Eats. These rows show `Business decision 2026-11-05` as
Evidence.

When a rule changes an entry's category, the entry takes the new category as its Menu, so it does
not keep the path of the category it left. Sandwiches also keep their reference Menu group.

## Recategorizing (reviewer decisions)

`recategorize.py` records a reviewer's decisions in `review_decisions.json`, by Mapping ID, so
every rebuild keeps them. Each decision holds the status (`Approved` or `Pending`), the Menu List
section, the Category that goes with it, the item name, who decided and a note. No sales figures.

```powershell
# 1. A fill-in sheet: one row per item, with New Section, New Item Name and New Status to fill.
& 'D:\Anaconda\python.exe' mapping_first\recategorize.py template --output 'Result\Recategorize Template.xlsx'
# 2. Read it back (--by names who decided; the default is Cheena and today's date).
& 'D:\Anaconda\python.exe' mapping_first\recategorize.py --by 'Cheena 2026-11-20' apply 'Result\Recategorize Template.xlsx'
# Or read a report marked up the way Cheena did on 2026-11-06 (see below).
& 'D:\Anaconda\python.exe' mapping_first\recategorize.py from-report 'cheene_data\USA PMix - Menu List Report DRAFT.xlsx'
# 3. Rebuild (or add --rebuild to step 2).
& 'D:\Anaconda\python.exe' mapping_first\run_pipeline.py --label recategorize
```

- Sections are the headings of the Menu Items List plus `PROMOTIONS` and `MISCELLANEOUS`, which
  are not on the list. Reviewer spellings such as `Fluffy All Day - Entrée` or `Croffles` are
  matched to the heading.
- The Category follows the section: an entry keeps its own Category when it fits the section
  (for example a Fluffy All Day item under FLUFFY ALL DAY - SANDWICHES), otherwise it takes the
  section's usual one. `SECTION_CATEGORIES` in `pmix_workflow.py` holds the table, taken from
  Cheena's approvals; COMBOS is Combos, ICE CREAM and SMALL BITES are Add-Ons.
- `from-report` reads three things from a reviewed report: PMIXData's `Review Status` column, where
  each item is `Approved`, `Pending` or a section name to move it to; the same column on the
  `Exceptions` sheet (2026-11-09), where a section places an item the pivot does not hold,
  `Uncategorized` means the reviewer could not place it (it stays a Pending exception, with the
  review recorded), and an `Excluded` entry such as Custom Amount is left Excluded; and `CATEGORY` / `REMARKS`
  columns typed beside the items on the PMIX sheet. Remarks understood: `Rename to X`, `Remove
  "text" in item name` (spacing-insensitive), `Remove # character`, a bare name, and `..., if item
  is under <section> Category`. A rename that lands on another remarked name follows it, and a
  name one typo from a list name takes the list spelling (never by dropping a word).
- Every run writes `Result/Recategorize Check - <kind> <time>.xlsx`: every decision, anything it
  could not apply, renamed items whose new name is not on the list, and entries still Pending.
- A decision replaces an earlier one for the same entry. Delete an entry's block from
  `review_decisions.json` to undo it. `from-report` records only what differs from the decisions
  already on file, so earlier reviews keep their own date and wording when a later report repeats
  them; the check workbook's `Already on file` sheet lists the ones it left alone.

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
| `Renamed by review` | A reviewer renamed it in `review_decisions.json` | reviewer's name |

`Menu List Match` shows the matched name and its section(s). Only the `Item` column is ever
respelled; `Source Item` and the Mapping ID stay as exported, and whitespace is always trimmed.
Nothing here changes a Category and nothing is marked Approved. The first two statuses are
exact-spelling matches. Items that are not on the list are noted and never guessed at, as Cheena
asked; most are free text, typos or POS variants such as `(Cold or Hot)`.

`menu_name_overrides.json` is the place for a typo or variant that a person has judged to be the
same product as a list name. It holds seven today, accepted on 2026-11-05: `Calssic Veg Mayo
Sandwich`, `Veggi Sandwich`, `Chicken Wings (6)`, `Chicken Wings (10)`, and `Deep Fried Chicken
Gyoza`, `Deep Fried Pork Gyoza` and `Deep Fried Edamame Gyoza` (without their portion size),
renamed to their list names. An entry is reversible by deleting its block. Each line holds the Mapping ID, the source name, the
list name and the reason, and no sales figures. A rebuild fails loudly if a listed override name
has since been removed from the list.

`menu_list_additions.json` holds names Cheena approved adding to the list (2026-11-09: 16 names
that her review had introduced, such as `Eggs Benedict Pancake (Ham and Cheese)` and `Apple
Caramel Pancakes`), by section. `Menu Items List.xlsx` is not modified: the names are added at the
end of their section each time the list is read, so the mapping, its `Menu Items List` sheet and
every report see them. A name already on the list is skipped, so the file can stay if the workbook
is later updated by hand; a section that is not on the list stops the build. To give Cheena the
list with the names in it:

```powershell
& 'D:\Anaconda\python.exe' mapping_first\recategorize.py list-copy --output 'Result\Menu Items List - with additions.xlsx'
```

The list's own spellings are copied as they are, including its typos (for example `Sapporro`,
`Franch Sparkling Wine`, `Creme Burlee Sauce`). Fix the typo in `Menu Items List.xlsx` and
rebuild if the report should show the corrected name. Most misspellings in the POS data come from
cashiers typing free-text item names at the register (Cheena, 2026-11-05), which is why so many
entries are `Not in Menu Items List`.

## Inputs and outputs

- Input: `../USA PMix.xlsx`, sheets `TOAST`, `SQUARE` and `Menu Category`. It is never modified.
- Input: `../Menu Items List.xlsx`, sheet `USA Menu List`, used by `map` to clean item names.
- `Menu List Section` in the master is the section each entry is reported under in the Menu
  List report: the reviewer's choice if there is one, else the list section of its Item.
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
  - `PMIXData` feeds the pivot, with the sale `Time` (TOAST `Sent Date - Copy`, SQUARE `Time`)
    beside the Date. `Menu Category` is a copy of the exact mapping used.
  - `TOAST` and `SQUARE` keep every input row and column, with mapping status appended.
  - `Exceptions` lists rows outside the pivot. `Location Coverage` gives per-store exception
    Qty and %. `Validation` reconciles pivot Qty plus exception Qty to total Qty.
- `USA PMix - Menu List Report DRAFT.xlsx` is the same report built with `--layout menu-list`.
  Rows expand as Menu List Section > Item, following the sections of the Menu Items List
  (PANCAKES, DRINKS - TEA, and so on). The first slicer is `Menu List Section` instead of
  `Category`; Location, Month and Week starting are unchanged, as are the `By Location` sheet,
  the rows in the pivot and every total.
  - Sections come from the master's `Menu List Section` column, so a reviewer can place an item
    that is not on the list. Masters without that column fall back to the list copy saved in the
    mapping workbook, matched on `Item`. Items only in UNCATEGORIZED are under `UNCATEGORIZED`;
    unplaced items not on the list are under `Not in Menu Items List`.
  - A name listed in two sections goes to the one that matches its own Category, otherwise the
    first. Today that is only `Tater Tots (1 serving)` and `(2 serving)`, under ADD-ONS and
    APPETIZER.

## Running on new data

`run_pipeline.py` does the whole run in one command: unit tests, the mapping, both reports,
`validate_prepared.py` and `verify_result.py` on each report.

```powershell
& 'D:\Anaconda\python.exe' mapping_first\run_pipeline.py `
  --source 'D:\data\USA PMix Nov.xlsx' --label nov-data
```

- `--source` is the new POS export, a workbook with `TOAST` and `SQUARE` sheets in the same
  columns as today. It does not need a `Menu Category` sheet; the outline is read from
  `--reference` (default `USA PMix.xlsx`).
- `--previous` is the reviewed mapping to learn from. It defaults to the one in `Result/`. When
  Cheena sends back a reviewed copy, pass that file.
- Everything is built in `Result/_building/<run>/` first. Only when every step passes are the
  current `Result/` files moved to `versions/<date>-before-<label>/` and the new files moved in.
  If a step fails, `Result/` is not touched and the run's log is in that folder.
- `--dry-run` builds and checks but leaves `Result/` as it is.
- If a `Result/` workbook is open in Excel when the checks finish, the checked output stays in
  `Result/_building/<run>/`; close Excel and run `--publish-staged <that folder>`.
- Default is a draft over all dates. For the approved report, add `--final --start YYYY-MM-DD
  --end YYYY-MM-DD`; it stops if any entry in that window is still Pending.
- `Result/PIPELINE_RUN.json` and `PIPELINE_RUN.log` record the run: inputs, every step and its
  time, and the mapping counts below.

What the new mapping keeps from `--previous`, matched by Mapping ID:

- `Approved` and `Excluded` entries keep Category, Menu, Menu group, Item, Review Status,
  Evidence and Review Notes. Counts (Rows, Net Qty, Locations) come from the new data.
- `Pending` entries are rebuilt by the rules, so rule changes reach them. A hand edit on a
  Pending row is not kept; mark the row Approved to keep it.
- Reviewed entries for items not in the new data go to the `Earlier Decisions` sheet and are
  used again when the item sells again.

The `Changes` sheet in the mapping lists what needs a look: `New item` (not in the previous
mapping), `Reviewed decision kept` (the rules now disagree with an approved row; the approval
wins), and `Pending entry changed` (with the old values, so nothing is lost silently).

## Commands

Run from the project root in PowerShell with the Anaconda interpreter, which has pandas,
openpyxl, XlsxWriter and pywin32. The report and `verify_result.py` need Excel for Windows.
Outputs are never overwritten, so remove or rename an old output before rebuilding it.

```powershell
& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py map `
  --output 'Result\USA PMix - Menu Mapping.xlsx'   # --menu-list <file> uses another list
# --previous <mapping> keeps its reviewed decisions; --reference <file> reads the outline elsewhere

& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py report `
  --mapping 'Result\USA PMix - Menu Mapping.xlsx' `
  --output 'Result\USA PMix - PMIX Report DRAFT.xlsx' `
  --draft --all-dates
```

`prepare` takes the same arguments as `report` and writes the data sheets without a PivotTable.

The Menu List report uses the same mapping:

```powershell
& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py report --layout menu-list `
  --mapping 'Result\USA PMix - Menu Mapping.xlsx' `
  --output 'Result\USA PMix - Menu List Report DRAFT.xlsx' `
  --draft --all-dates
```

After review, edit Category, Menu, Menu group, Item and Review Notes in the master. Filter
`Menu List Status` for `Not in Menu Items List` to see the items still to be cleaned. Mark
supported rows `Approved`. Mark genuine exclusions `Excluded` with a reason. Do not edit Mapping
ID or the source identity columns. A new `map` run keeps Approved and Excluded rows when given
`--previous`; other hand edits live only in that saved file. Changing the master affects the **next report build**; Excel Refresh
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

## Adding sales amounts (Gross Sales, Discount, Net Sales)

Cheena asked on 2026-11-09 whether extra columns in PMIXData would cause a problem.

- **The PivotTable can take them.** `PMIXData` is an Excel Table and the PivotTable's source is
  the table name, so a column inside the table becomes a field after Refresh. Existing rows,
  slicers and totals do not change. `validate_prepared.py` reads PMIXData by column name, so an
  extra column does not break it either.
- **The data has no amounts.** `USA PMix.xlsx` holds only Qty in `TOAST` and `SQUARE` (eight
  columns each). The amounts must come from a new POS export with one row per item line, and
  the workflow never invents them.
- **Do not type them into the delivered file.** PMIXData is rewritten from the source on every
  build, so a hand-added column disappears at the next rebuild. To keep them, the build has to
  carry the columns: read them in `load_sources`, add them to the report's field list, add the
  value fields to both PivotTables, and add checks that `Net Sales = Gross Sales - Discount`
  and that the pivot totals match the source totals.
- **Questions to settle first.** Whether returns and voids carry negative amounts, whether a
  discount applies to the item line or to the whole check (a check-level discount has to be
  split across items to be summed by item), and whether `Qty (%)` should stay as the share of
  quantity or become a share of Net Sales.
- Amounts are store sales data, so they stay out of Git like the other workbooks.

## Tests and checks

```powershell
& 'D:\Anaconda\python.exe' mapping_first\test_pmix_workflow.py
& 'D:\Anaconda\python.exe' mapping_first\validate_prepared.py
& 'D:\Anaconda\python.exe' mapping_first\verify_result.py
```

- `test_pmix_workflow.py` covers the mapping rules, menu-name cleaning and workbook helpers.
- `validate_prepared.py` compares every original cell and reconciles rows, Qty and all report keys,
  and checks each sale's Time, the saved sections, and that no report name has edge spaces.
- `verify_result.py` opens a temporary copy in live Excel, drives the real slicers,
  checks each result against an independent recompute, and writes
  `Result/VALIDATION_RESULTS.json`.
- Both checks read the date window from the report's own `Validation` sheet, so they work for the
  all-dates draft and for a dated final report. Use `--report`, `--mapping` and `--source` to
  point them at other files.
- For the Menu List report, run `validate_prepared.py --output 'Result\USA PMix - Menu List
  Report DRAFT.xlsx'` and `verify_result.py --layout menu-list --report` with the same file. The
  live check works out each sale's section on its own from the saved list copy, requires any
  saved section that differs to be a recorded review decision, and writes
  `Result/VALIDATION_RESULTS - Menu List.json`.

These prove the arithmetic and that the rules were applied. They do not prove a category is what
the business wants; only review can do that.
