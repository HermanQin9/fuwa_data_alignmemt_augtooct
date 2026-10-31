# Files to upload

1. `../USA PMix.xlsx` — original POS sales and original Menu Category outline.
2. `../Quarterly PMIX Report.xlsx` — reference quarterly layout for Schaumburg.
3. `pmix_workflow.py` — corrected implementation.
4. `USA PMix - Menu Mapping.xlsx` — editable mapping master, Stage 1 output.
5. `USA PMix - Pivot Source.xlsx` **from this folder** — prepared Stage 2 data,
   NOT an interactive final report.
6. `test_pmix_workflow.py`, `validate_prepared.py`, `README.md`, and both
   screenshots, if possible.
7. Once successfully generated, `USA PMix - Recategorized.xlsx` **from this
   folder**, not the earlier root output — the intended native PivotTable draft.
8. `OUTPUT_AUDIT.xlsx` and `VALIDATION_RESULTS.json` — the latest audit evidence.
   The audit corrected nine draft mappings affecting 560 source rows / 594 net
   Qty, but did not approve them or complete the native workbook.

Only upload business information you are authorized to share.

# Prompt to paste into Claude

Act as an independent Excel/data-quality auditor and Python reviewer. Inspect
the actual uploaded files, not prior claims of correctness. First judge whether
the implementation meets the objective, then identify the corrections required.

Current limitation: native Excel generation was attempted but the execution
sandbox denied Excel COM activation. The prepared source workbook exists, but
the new interactive final report has NOT been generated or verified. Git branch
creation is also blocked by that environment. Do not confuse the older
root-level static report with a newly generated interactive report, or treat
the prepared source as meeting the interactive requirement.

## Correct business objective

This is a two-stage pipeline:

1. Map items from BOTH TOAST and SQUARE into an editable master sheet called
   Menu Category, covering all source location labels. Preserve original
   identities and the reference menu hierarchy.
2. Use that saved mapping, not a second independent classifier, to calculate
   quantities and percentages for a user-selected time window. Deliver one
   interactive Excel report in the style of Quarterly PMIX Report.xlsx and
   screenshot 1.

The final table must have exactly the eight top-level categories in screenshot 2:
Pancakes; Fluffy All Day; Drinks; Japanese Eats; Cheesecakes; Pastries; Promotions;
Add-Ons. The screenshot spells Promotions "Pormotions"; the implementation
deliberately corrects that spelling.

Users must be able to:
- Expand/collapse Category > Menu > Menu group > Item.
- Select one or multiple locations.
- Choose dates/months and year.
- Filter one or multiple categories.
- See recalculated Qty (#), Qty (%), quarterly columns, and grand totals.

Native Excel PivotTables/slicers/timeline are expected. Static formatted totals,
worksheet row outlines, and ordinary autofilters alone do not satisfy this UX.
The pivot uses classified net item Qty and % of filtered grand total.
Expanding a category is not the same as filtering to it.

The user approved all available dates for the INITIAL DRAFT. That is not approval
of a final date window or of all category assignments.

## File roles

- USA PMix.xlsx: original input.
- Quarterly PMIX Report.xlsx: format/percentage reference, not the same period.
- USA PMix - Menu Mapping.xlsx: Stage 1 master; mapping suggestions are Pending.
- USA PMix - Pivot Source.xlsx: prepared Stage 2 workpaper, not a finished report.
- USA PMix - Recategorized.xlsx (new mapping_first version): the intended native
  Stage 2 draft; only audit its actual behavior if it has been built and supplied.
- pmix_workflow.py: implementation being judged.
- Earlier root-level recategorize_menu.py/output, if supplied: historical only.

## Required audit

1. Read the input schemas, all location labels, reference outline indentation,
   report Data schema, and pivot grouping/percentage conventions.
2. Verify every distinct POS/group/item mapping, including blank groups and
   ambiguous descriptions. Distinguish evidence-supported mappings, likely
   errors and decisions requiring Cheena's confirmation. Zero unmapped labels
   would not prove semantic correctness.
3. Check that the saved Menu Category master is the actual reporting input.
   Change a mapping in a temporary copy, rebuild, and verify the corresponding
   report rows change without recategorizing them using hidden rules.
4. Check originals row by row and cell by cell, including date/time value types,
   booleans, blanks, locations, quantities, original order, zero quantities,
   refunds and identical-looking sales lines. Include Excel row references.
5. Independently reconcile row counts and net quantity:
   expected benchmarks to verify are TOAST 185,823 rows, SQUARE 6,984 rows,
   combined 192,807 item-sales lines and 202,865 net item Qty.
6. If a native report is supplied, inspect its actual pivot parts/cache, row hierarchy, eight category labels,
   slicers, and timeline. If Excel automation is available, test filtering,
   multi-selection, expansion and reopening. Otherwise explicitly say that
   interactive Excel behavior was not verified; do not treat cached values as
   proof that controls work.
7. Test date boundaries inclusively, partial months/quarters, multi-year data,
   empty windows, and reversed dates. No hardcoded Q3/Q4 assumptions or
   combining identical quarter names from different years.
8. Recalculate quantities and percentages after filtering location, category
   and dates individually and in combination. Match Excel's denominator to
   the declared % of filtered classified grand total policy.
9. Audit Exceptions and Validation. Unresolved/excluded selected rows must be
   preserved and reconcile: selected Qty = pivot Qty + exceptions Qty.
   They must not be forced into invented categories to make totals look final.
   Final builds must refuse Pending mappings. Drafts must visibly disclose
   their pending status and exception quantities.
10. Compare Report Data to the reference's Year, Quarter, Menu, Menu group,
    Item, Qty sold schema. Clearly distinguish intentional Location/Category
    extensions from incompatibilities; do not claim identical layouts.
11. Verify all source location labels are covered. Use the actual workbook to
    resolve reversed TOAST/SQUARE descriptions. Do not merge Tucker and
    Tucker GA without evidence, or claim source labels prove distinct stores.
12. Do not equate Qty sold with unique orders: no order/check ID exists.
    No sales/revenue measures should be fabricated from absent inputs.
13. Run the tests and reproduce prepared/native outputs to different temporary filenames.
    Do not overwrite uploaded originals or silently repair the output.

## Deliverable

- Verdict: PASS / PASS WITH CONDITIONS / FAIL.
- Confirmed scope and two-stage input/output flow.
- Findings table with severity, code line or sheet/cell, actual result, expected
  result or uncertainty, evidence, and affected rows/quantity.
- Reconciliation table and interactive-control test results.
- Proposed corrected mappings with reasons, leaving business uncertainties
  flagged rather than guessed.
- Specific decisions Cheena must approve.
- Exact review coverage and limitations: distinguish automated full-row checks
  from individual semantic review. Never claim a manual review you did not do.
- Prioritized fixes before a final approved report can be delivered.
