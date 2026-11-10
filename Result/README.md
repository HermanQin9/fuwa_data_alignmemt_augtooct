# Result

The current output of the PMIX workflow. Everything here is a draft until the business approves
the mapping.

| File | What it is | In Git |
|---|---|---|
| `USA PMix - PMIX Report DRAFT.xlsx` | The interactive report, rows by Category > Menu > Menu group > Item | No |
| `USA PMix - Menu List Report DRAFT.xlsx` | The same report with rows by Menu Items List section > Item | No |
| `USA PMix - Menu Mapping.xlsx` | The editable category mapping behind both reports | No |
| `Menu Items List - with additions.xlsx` | Cheena's Menu Items List plus the 16 names she approved on 2026-11-09, for her to finalize. `../Menu Items List.xlsx` itself is unchanged | No |
| `VALIDATION_RESULTS.json` | Overall totals and pass/fail results for the PMIX report | Yes |
| `VALIDATION_RESULTS - Menu List.json` | The same for the Menu List report | Yes |
| `PIPELINE_RUN.json`, `PIPELINE_RUN.log` | What the last `run_pipeline.py` run used and did (after the first pipeline run) | No |
| `README.md` | This file | Yes |

Each report opens on `PMIX`, the main table. The `By Location` sheet shows every item in every
location, with an All locations total. The same slicers drive both sheets: Category (or Menu
List Section in the Menu List report), Location, Month and Week starting. Both reports have the
same totals; only the row grouping differs.

In the mapping workbook, `Menu List Status` and `Menu List Match` show which items are on Cheena's
`Menu Items List.xlsx`. Filter `Menu List Status` for `Not in Menu Items List` to see the items
that still need a clean name. Item names were cleaned against that list. Cheena's review of
2026-11-06 and 2026-11-09 then set sections, names and approvals; `Menu List Section` shows where
each item is reported, and Evidence `Reviewed by Cheena <date>` marks her decisions.

The repository is public, so the workbooks stay on the local machine. They hold store sales data
and are ignored by Git.

`Recategorize Check - ... .xlsx` lists what `recategorize.py` recorded from a review: every
decision, anything it could not apply, and what is still Pending.

PMIXData has a `Time` column with each sale's time. Names in the report are trimmed; the
`Source Item` and `Source Group` columns keep the POS spelling, spaces included, because they
identify each entry.

A mapping built by the pipeline also has a `Changes` sheet (new items, and rows whose rules
changed) and an `Earlier Decisions` sheet (approved items not in the current data).

To run everything on new data, or rebuild or re-check these files, see [`../mapping_first/README.md`](../mapping_first/README.md).
