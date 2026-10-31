# Project notes for Claude

## Auto upload to GitHub
- Remote: `origin` = https://github.com/HermanQin9/fuwa_data_alignmemt_augtooct.git (public repo).
- Work on the branch `feature/data-alignment`. Never commit or push to `main` directly.
- After finishing any change to code, tests, docs or this file, and once tests pass
  (`python mapping_first/test_pmix_workflow.py`), commit and push to `origin feature/data-alignment`
  without asking again. Use a short message describing the change.
- End commit messages with the attribution lines given in the session's system reminder.
- Never use `--force`, `--no-verify`, or push to another branch without being asked.

## Never upload (repo is public)
- Excel/CSV files (`*.xlsx`, `*.xls`, `*.csv`) contain store sales data. They are in `.gitignore`.
  Do not remove those rules, and do not use `git add -f` on them, unless the user explicitly says
  the data may be published.
- Do not commit lock files (`~$*`), `__pycache__`, `.pkl` files, or scratch output.
- Before every push, run `git status` and `git diff --cached --stat` and confirm only code/docs are staged.

## Data rules
- Do not modify the original source workbooks (`USA PMix.xlsx`, `Quarterly PMIX Report.xlsx`).
- Mapping changes go in new files; never mark a mapping Approved without the user saying so.
- Do not force unknown items into a category to reach zero unmapped.
