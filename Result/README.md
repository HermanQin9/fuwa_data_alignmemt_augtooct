# Result

The current output of the PMIX workflow. Everything here is a draft until the business approves
the mapping.

| File | What it is | In Git |
|---|---|---|
| `USA PMix - PMIX Report DRAFT.xlsx` | The interactive report | No |
| `USA PMix - Menu Mapping.xlsx` | The editable category mapping behind the report | No |
| `USA PMix - Result Check.xlsx` | Check results, changes from the first draft, review findings and open questions | No |
| `VALIDATION_RESULTS.json` | Overall totals and pass/fail results | Yes |
| `README.md` | This file | Yes |

The report opens on `PMIX`, the main table. The `By Location` sheet shows every item in every
location, with an All locations total. The same slicers drive both sheets.

In the mapping workbook, `Menu List Status` and `Menu List Match` show which items are on Cheena's
`Menu Items List.xlsx`. Filter `Menu List Status` for `Not in Menu Items List` to see the items
that still need a clean name. Item names were cleaned against that list; categories were not changed.

The repository is public, so the workbooks stay on the local machine. They hold store sales data
and are ignored by Git.

To rebuild or re-check these files, see [`../mapping_first/README.md`](../mapping_first/README.md).
