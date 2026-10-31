# Mapping-first PMIX workflow

## What changed

This is the corrected implementation, isolated from the earlier script and static
report. The earlier deliverables are preserved in
[`../versions/2026-10-30-before-mapping-first/`](../versions/2026-10-30-before-mapping-first/).
Git branch creation is **blocked**, not complete: the session sandbox denied
execution of the installed Git binary. The intended branch is
`feature/menu-category-pivot`. A separate folder is not a Git branch.

The user approved a **draft using all supplied dates**, with filters to narrow
the view. The original source and earlier report are not overwritten.

### Current build status

- The editable master `USA PMix - Menu Mapping.xlsx` was generated.
- `USA PMix - Pivot Source.xlsx` is **prepared data only**, not a working pivot.
- The full report build was attempted, but Excel COM activation was denied by
  the sandbox. The final interactive `USA PMix - Recategorized.xlsx` **has not
  been generated or interactively tested** in this folder.
- The earlier root output is still the earlier static report, not this new
  implementation's output.
- The mapping contains 2,102 distinct source identities: 1,664 have draft category
  suggestions and 438 need resolution. All 2,102 are Pending business review.
  These are mapping-entry counts, not item-sales row counts.
- Thirteen regression tests passed. An independent audit compared all **1,542,456
  original cells**, including value types, in source row order: PASS.
- All 192,807 item-sales lines and 202,865 net Qty are preserved. The prepared
  pivot data has 185,121 rows / 194,679 net Qty. The remaining 7,686 rows /
  8,186 net Qty are explicitly retained in Exceptions, not silently counted
  as approved categories. These counts do not constitute semantic approval.
- The output audit corrected nine draft mapping entries affecting **560 source
  rows / 594 net Qty**. Filled savory croissants now suggest Fluffy All Day;
  tea-flavored basque cheesecake suggests Cheesecakes; explicit pancake/drink
  bundles and a buy/get-free offer suggest Promotions. All remain Pending.
- The expanded verifier also checks all **3,786 detailed Report Data keys**,
  mapping counts/Qty per identity, year/quarter/month fields, each exception's
  mapped fields and quantity, and Validation totals: PASS.
- [`OUTPUT_AUDIT.xlsx`](OUTPUT_AUDIT.xlsx) lists corrected mappings, unresolved
  entries and context-dependent cases needing review. The previous draft is
  preserved in
  [`../versions/2026-10-30-before-output-audit/`](../versions/2026-10-30-before-output-audit/).

The sandbox needs to permit Git under `D:\Software\Github\Git\` and Excel
automation at `C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE` before
the native creation/filter tests and branch creation can be completed here.

## Inputs and outputs

1. Input: `../USA PMix.xlsx`, sheets `TOAST`, `SQUARE`, and the original
   `Menu Category` reference outline.
2. Mapping stage: `USA PMix - Menu Mapping.xlsx`.
   - `Menu Category` is the editable master lookup, built from **both** POS sheets.
   - One row represents a distinct POS sheet / source group / source item.
   - All locations contribute to this lookup; `Locations`, `Rows`, and `Net Qty`
     describe coverage.
   - Original indented reference paths are available in `Reference Outline`.
   - Category proposals are **Pending**, not automatically approved.
3. Reporting stage target: `USA PMix - Recategorized.xlsx` (not built yet).
   - `PMIX` is a **native Excel PivotTable**, not a static summary.
   - Rows expand as **Category > Menu > Menu group > Item**.
   - Columns show year-aware quarters, `Qty (#)` and `Qty (%)`, and totals.
   - Location, category and month slicers, a date timeline, and year/date
     report filters recalculate quantities and percentages.
   - `PMIXData` is the Excel table feeding the pivot.
   - `Menu Category` is a copy of the exact saved lookup used for that build.
   - `Report Data` contains Location, Year, Quarter, Category, Menu, Menu group,
     Item and Qty sold. Category and Location extend the example Data schema.
   - `TOAST` and `SQUARE` preserve all input rows and original columns.
     Appended columns identify mapping status and inclusion in the report window.
   - `Exceptions` preserves unresolved/excluded selected rows outside the pivot.
   - `Validation` reconciles selected Qty to pivot Qty plus exception Qty.
   - `Location Coverage` lists all source location labels in the selected window,
     including labels whose rows may be unresolved.

## Nine categories

Pancakes; Fluffy All Day; Drinks; Japanese Eats; Cheesecakes; Pastries;
Promotions; Add-Ons; Sandwiches (added at Cheena's request, 2026-10-31).

`Promotions` deliberately corrects `Pormotions` in the screenshot/reference.
There is **no tenth category in the pivot**. Unresolved data is not forced into
one of the nine categories, and is not silently discarded.

Reference outline entries are evidence, not universal business rules. In
particular, regular products versus promotion sales, add-ons versus full dishes,
ice cream versus cheesecake, and ambiguous open-item descriptions need review.
Draft suggestions are not a claim of semantic correctness.

## Commands

Run from the project root in PowerShell, with the existing Anaconda environment.
It has pandas, openpyxl, XlsxWriter and pywin32 installed. The VS Code-selected
interpreter did not have these packages during this session.

```powershell
& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py map `
  --output 'mapping_first\USA PMix - Menu Mapping.xlsx'

& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py report `
  --mapping 'mapping_first\USA PMix - Menu Mapping.xlsx' `
  --output 'mapping_first\USA PMix - Recategorized.xlsx' `
  --draft --all-dates
```

Outputs are not overwritten automatically, particularly the editable master.
Use a new output filename when rebuilding a report.

The explicit noninteractive preparation command is:

```powershell
& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py prepare `
  --mapping 'mapping_first\USA PMix - Menu Mapping.xlsx' `
  --output 'mapping_first\USA PMix - Pivot Source.xlsx' `
  --draft --all-dates
```

`prepare` does not activate Excel and never claims to build a PivotTable.
It is a workpaper for auditing the mapped source, not a substitute for `report`.

After review, edit Category, Menu, Menu group, Item and Review Notes in the
mapping master. Mark supported rows `Approved`. Mark genuine exclusions
`Excluded` and give a reason; retain them for reconciliation. Do not edit the
Mapping ID or source identity columns.

For an approved report in an explicitly supplied inclusive window:

```powershell
& 'D:\Anaconda\python.exe' mapping_first\pmix_workflow.py report `
  --mapping 'mapping_first\USA PMix - Menu Mapping.xlsx' `
  --output 'mapping_first\USA PMix - Approved Window.xlsx' `
  --start 2026-08-03 --end 2026-10-25
```

The dates above illustrate the available full window, not a separately agreed
final reporting period. A final build refuses Pending mappings within its window.
Changing the mapping affects the **next report build**; Excel Refresh alone does
not rerun the mapping stage.

The timeline can narrow the source window in a built workbook. To include dates
outside that workbook's data window, rebuild with a wider window and appropriate
source data. No macro is required for filtering or expand/collapse.

## Measures and filters

- `Qty (#)` is **sum of item quantity**, including negative returns and zero
  quantity lines. It is not a distinct customer-order count; the input lacks
  order/check IDs.
- `Qty (%)` uses Excel's **% of Grand Total**, recomputed for the filtered,
  classified data. Expanding one category without filtering leaves its share
  relative to the whole selection. Selecting only one category makes that
  category's visible total 100%.
- Unresolved/excluded quantities are outside this denominator and explicitly
  reconciled on `Validation`. A draft is not final until those cases are reviewed.
- Quarters are calendar quarters, labeled with the year. October is Q4.
- Source location labels are retained, except that Square "Fluffy Fluffy Tucker GA"
  is renamed to the Toast label "Fluffy Fluffy - Tucker" (Cheena confirmed they are
  the same store, 2026-10-31). The actual input places Miami and Carrollton in SQUARE.
- No gross sales, net sales, or revenue percentages are invented; those inputs
  are absent from this workbook.

## Tests

```powershell
& 'D:\Anaconda\python.exe' -m unittest discover `
  -s mapping_first -p test_pmix_workflow.py -v

& 'D:\Anaconda\python.exe' mapping_first\validate_prepared.py
```

The `report` command attempts to create its native PivotTable through a separate
Excel COM instance and save it with Excel itself. The user's existing Excel
workbooks are never intentionally closed by the pipeline.
