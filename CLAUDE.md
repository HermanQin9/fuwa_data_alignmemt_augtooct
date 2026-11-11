# Project notes for Claude

## Auto upload to GitHub
- Remote: `origin` = https://github.com/HermanQin9/fuwa_data_alignmemt_augtooct.git (public repo).
- `feature/data-alignment` is the base branch. Never commit or push to `main` directly.
- Each piece of work goes on its own `feature/*` branch cut from `feature/data-alignment`
  (for example `feature/menu-list-cleanup`). Small fixes may go straight on `feature/data-alignment`.
- After finishing any change to code, tests, docs or this file, and once tests pass
  (`python mapping_first/test_pmix_workflow.py`), commit and push the current `feature/*` branch to
  `origin` without asking again. Use a short message describing the change.
- When a `feature/*` branch is finished, open a pull request into `feature/data-alignment`.
  Merge it only when the user says so.
- End commit messages with the attribution lines given in the session's system reminder.
- Never use `--force` or `--no-verify`, and never push to `main` or a branch that is not
  `feature/*` without being asked.

## Never upload (repo is public)
- Excel/CSV files (`*.xlsx`, `*.xls`, `*.csv`) contain store sales data. They are in `.gitignore`.
  Do not remove those rules, and do not use `git add -f` on them, unless the user explicitly says
  the data may be published.
- Do not commit lock files (`~$*`), `__pycache__`, `.pkl` files, or scratch output.
- Before every push, run `git status` and `git diff --cached --stat` and confirm only code/docs are staged.

## Where things go
- `Result/` holds the current output only. `mapping_first/` holds code and docs. `versions/` keeps
  earlier output; move old files there rather than deleting them.
- Rebuild and check steps are in `mapping_first/README.md`. After changing a mapping rule, rebuild
  the output, then run `validate_prepared.py` and `verify_result.py` as well as the unit tests.

## Data rules
- Do not modify the original source workbooks (`USA PMix.xlsx`, `Quarterly PMIX Report.xlsx`,
  `Menu Items List.xlsx`). The last is Cheena's cleaned menu list and the reference for item names.
- Business decisions go in `business_rule` or `mapping_overrides.json`, so a rebuild keeps them.
  Item-name decisions go in `menu_name_overrides.json`. Names Cheena approved adding to the menu
  list go in `menu_list_additions.json`; `Menu Items List.xlsx` itself stays unmodified.
- Never mark a mapping Approved without the user saying so.
- Do not force unknown items into a category to reach zero unmapped.
- Do not put category-level or store-level Qty in tracked files. The repo is public.
