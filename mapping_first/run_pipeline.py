"""Run the whole PMIX pipeline on a data export: tests, mapping, the report and every check.

Everything is built in Result/_building/<run> first. Only when every step passes are the
current files in Result/ moved to versions/<date>-before-<label>/ and the new ones put in their
place, so a failed run never leaves Result/ half replaced.

    python mapping_first/run_pipeline.py --source "D:\\data\\USA PMix Nov.xlsx" --label nov-data

The previous mapping (default: the one in Result/, or a reviewed copy Cheena sends back) passes
its Approved and Excluded decisions to the new mapping. See README.md, "Running on new data".
"""
from __future__ import annotations

import argparse
import contextlib
import datetime as dt
import importlib
import io
import json
import os
import shutil
import subprocess
import sys
import time
import traceback
from pathlib import Path

# The packaged Windows app (pmix_app.py) runs from one folder: the .exe, the reference workbooks,
# decisions/, Result/ and versions/ all sit together, and every step runs inside the .exe.
FROZEN = getattr(sys, "frozen", False)
HERE = Path(__file__).resolve().parent
ROOT = Path(sys.executable).resolve().parent if FROZEN else HERE.parent
# Steps run in this process when packaged (there is no separate Python to start); set
# PMIX_INPROCESS=1 to try that path from source.
INPROCESS = FROZEN or bool(os.environ.get("PMIX_INPROCESS"))
RESULT = ROOT / "Result"
VERSIONS = ROOT / "versions"
STAGING = RESULT / "_building"
MAPPING_NAME = "USA PMix - Menu Mapping.xlsx"
REPORTS = {"category": "USA PMix - PMIX Report", "menu-list": "USA PMix - Menu List Report"}
JSON_NAME = {"category": "VALIDATION_RESULTS.json", "menu-list": "VALIDATION_RESULTS - Menu List.json"}
RUN_SUMMARY = "PIPELINE_RUN.json"
KEEP_IN_RESULT = {"README.md"}


class StepFailed(RuntimeError):
    pass


def call_in_process(script: Path, args: list) -> tuple[int, str, str]:
    """Run a step's main() here, with its own argv and its output captured."""
    out, err = io.StringIO(), io.StringIO()
    old_argv, code = sys.argv, 0
    sys.argv = [str(script), *map(str, args)]
    try:
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                sys.path.insert(0, str(HERE))
                importlib.import_module(Path(script).stem).main()
            except SystemExit as stop:
                code = stop.code if isinstance(stop.code, int) else (0 if stop.code is None else 1)
                if stop.code and not isinstance(stop.code, int):
                    print(stop.code, file=err)
            except Exception:
                traceback.print_exc()
                code = 1
    finally:
        sys.argv = old_argv
    return code, out.getvalue(), err.getvalue()


def run_step(name: str, args: list, log: Path, steps: list) -> str:
    """Run one Python step, log its output, and stop the pipeline if it fails."""
    print(f"=== {name}", flush=True)
    started = time.time()
    if INPROCESS and Path(args[0]).stem != "test_pmix_workflow":
        code, stdout, stderr = call_in_process(Path(args[0]), args[1:])
    else:
        done = subprocess.run([sys.executable, *map(str, args)], cwd=ROOT, capture_output=True,
                              text=True, encoding="utf-8", errors="replace")
        code, stdout, stderr = done.returncode, done.stdout, done.stderr
    output = stdout + stderr
    with log.open("a", encoding="utf-8") as handle:
        handle.write(f"=== {name}\n{output}\n")
    seconds = round(time.time() - started)
    steps.append({"step": name, "result": "PASS" if code == 0 else "FAIL", "seconds": seconds})
    if code != 0:
        print(output[-3000:], flush=True)
        raise StepFailed(f"{name} failed (exit {code}); full output in {log}")
    print(f"    PASS in {seconds}s", flush=True)
    return stdout


def pending_entries(mapping: Path) -> int:
    """How many mapping entries are still Pending (the new items that need a decision)."""
    import openpyxl
    book = openpyxl.load_workbook(mapping, read_only=True, data_only=True)
    try:
        rows = book["Menu Category"].iter_rows(values_only=True)
        column = list(next(rows)).index("Review Status")
        return sum(1 for row in rows if row[column] == "Pending")
    finally:
        book.close()


def last_json(output: str) -> dict:
    """The JSON object a step prints at the end of its output."""
    start = output.rfind("\n{")
    try:
        return json.loads(output[start + 1 if start >= 0 else 0:])
    except ValueError:
        return {}


def preflight(args) -> None:
    for label, path in (("Source", args.source), ("Reference", args.reference),
                        ("Menu Items List", args.menu_list)):
        if not path.exists():
            raise SystemExit(f"{label} not found: {path}")
    if args.previous is not None and not args.previous.exists():
        raise SystemExit(f"Previous mapping not found: {args.previous}. "
                         "Use --no-previous only for a first run with nothing reviewed.")
    if (args.final or args.auto) and not (args.start and args.end):
        raise SystemExit("--final and --auto need --start and --end")
    if args.final and args.auto:
        raise SystemExit("Use --final or --auto, not both")
    if bool(args.start) != bool(args.end):
        raise SystemExit("Give both --start and --end, or neither")


def open_workbooks() -> list[str]:
    """Result workbooks open in Excel; they cannot be moved until closed."""
    return sorted(p.name[2:] for p in RESULT.glob("~$*"))


def archive_folder(label: str) -> Path:
    base = VERSIONS / f"{dt.date.today().isoformat()}-before-{label}"
    folder, n = base, 2
    while folder.exists():
        folder, n = base.with_name(f"{base.name}-{n}"), n + 1
    return folder


def publish(staged: list[Path], label: str) -> Path | None:
    """Move the current Result files to versions/, then the new ones into Result/."""
    if open_workbooks():
        raise RuntimeError(f"Close these workbooks in Excel first: {open_workbooks()}")
    current = [p for p in RESULT.iterdir() if p.is_file() and p.name not in KEEP_IN_RESULT]
    folder = None
    if current:
        folder = archive_folder(label)
        folder.mkdir(parents=True)
        for path in current:
            shutil.move(str(path), folder / path.name)
    for path in staged:
        shutil.move(str(path), RESULT / path.name)
    return folder


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--source", type=Path, default=ROOT / "USA PMix.xlsx",
                        help="POS export workbook with TOAST and SQUARE sheets")
    parser.add_argument("--reference", type=Path, default=ROOT / "USA PMix.xlsx",
                        help="Workbook with the 'Menu Category' outline")
    parser.add_argument("--menu-list", type=Path, default=ROOT / "Menu Items List.xlsx")
    parser.add_argument("--previous", type=Path, default=RESULT / MAPPING_NAME,
                        help="Reviewed mapping whose Approved/Excluded decisions are kept")
    parser.add_argument("--no-previous", action="store_true",
                        help="Start the mapping from the rules alone")
    parser.add_argument("--label", default="pipeline-run",
                        help="Name for the versions/ folder the current Result files move to")
    parser.add_argument("--start", help="Inclusive date YYYY-MM-DD (default: all dates)")
    parser.add_argument("--end", help="Inclusive date YYYY-MM-DD")
    parser.add_argument("--layout", choices=["menu-list", "category", "both"], default="menu-list",
                        help="Report to build. Default menu-list, the one Cheena chose on "
                             "2026-11-10; category is the older PMIX layout")
    parser.add_argument("--final", action="store_true",
                        help="Approved report: refuses Pending entries in the window")
    parser.add_argument("--auto", action="store_true",
                        help="Final report when the new mapping has no Pending entries; otherwise "
                             "a DRAFT that says how many entries still need a decision")
    parser.add_argument("--skip-tests", action="store_true",
                        help="Do not run the unit tests (the packaged app does not carry them)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Build and check in Result/_building only; leave Result/ unchanged")
    parser.add_argument("--publish-staged", type=Path, metavar="FOLDER",
                        help="Publish a checked run left in Result/_building (for example while a "
                             "Result workbook was open in Excel), without rebuilding")
    args = parser.parse_args()
    if args.publish_staged:
        return publish_staged(args.publish_staged, args.label)
    if args.no_previous:
        args.previous = None
    preflight(args)

    run = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    stage = STAGING / run
    stage.mkdir(parents=True)
    log = stage / "PIPELINE_RUN.log"
    mapping = stage / MAPPING_NAME
    layouts = list(REPORTS) if args.layout == "both" else [args.layout]
    window = ["--start", args.start, "--end", args.end] if args.start else ["--all-dates"]
    steps: list = []
    summary = {"run": run, "source": args.source.name, "reference": args.reference.name,
               "menu_list": args.menu_list.name,
               "previous_mapping": str(args.previous) if args.previous else None,
               "mode": "final" if args.final else "draft", "steps": steps}
    reports: dict = {}
    mode: list = []
    workflow = HERE / "pmix_workflow.py"
    try:
        if not args.skip_tests:
            run_step("unit tests", [HERE / "test_pmix_workflow.py"], log, steps)
        map_args = [workflow, "--source", args.source, "map", "--output", mapping,
                    "--menu-list", args.menu_list, "--reference", args.reference]
        if args.previous is not None:
            map_args += ["--previous", args.previous]
        summary["mapping"] = last_json(run_step("build mapping", map_args, log, steps))
        if args.auto:
            summary["pending_entries"] = pending_entries(mapping)
            args.final = summary["pending_entries"] == 0
            summary["mode"] = "final" if args.final else "draft"
            print(f"    {summary['pending_entries']} Pending entries: building the "
                  f"{summary['mode']} report", flush=True)
        suffix = "" if args.final else " DRAFT"
        mode = [] if args.final else ["--draft"]
        reports.update({layout: stage / f"{REPORTS[layout]}{suffix}.xlsx" for layout in layouts})
        for layout, report in reports.items():
            built = run_step(f"build {layout} report", [
                workflow, "--source", args.source, "report", "--layout", layout,
                "--mapping", mapping, "--output", report, *mode, *window], log, steps)
            summary.setdefault("reports", {})[report.name] = {
                k: v for k, v in last_json(built).items()
                if k in ("Inclusive start date", "Inclusive end date", "Selected net Qty",
                         "Pivot net Qty", "Exception net Qty", "Status")}
        for layout, report in reports.items():
            run_step(f"validate {layout} report", [
                HERE / "validate_prepared.py", "--source", args.source,
                "--mapping", mapping, "--output", report], log, steps)
            run_step(f"live Excel check, {layout} report", [
                HERE / "verify_result.py", "--source", args.source, "--mapping", mapping,
                "--report", report, "--layout", layout, "--json-dir", stage], log, steps)
    except StepFailed as error:
        summary["overall"] = "FAIL"
        (stage / RUN_SUMMARY).write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
        print(f"\nFAILED: {error}\nResult/ was not changed. Partial output: {stage}")
        return 1

    summary["overall"] = "PASS"
    (stage / RUN_SUMMARY).write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    staged = [mapping, *reports.values(), *(stage / JSON_NAME[layout] for layout in layouts),
              stage / RUN_SUMMARY, log]
    print("\nAll steps passed.")
    print(json.dumps(summary.get("mapping", {}), indent=2))
    if args.dry_run:
        print(f"Dry run: output left in {stage}; Result/ unchanged.")
        return 0
    if open_workbooks():
        print(f"Checked output is staged, but these are open in Excel: {open_workbooks()}.\n"
              f"Close them, then run: run_pipeline.py --publish-staged \"{stage}\" --label {args.label}")
        return 0
    return finish(stage, staged, args.label)


def finish(stage: Path, staged: list[Path], label: str) -> int:
    folder = publish(staged, label)
    stage.rmdir()
    if STAGING.exists() and not any(STAGING.iterdir()):
        STAGING.rmdir()
    if folder:
        print(f"Previous Result files moved to {folder.relative_to(ROOT)}")
    print("New output is in Result/.")
    return 0


def publish_staged(stage: Path, label: str) -> int:
    """Publish a run that passed every check but was left staged."""
    summary_path = stage / RUN_SUMMARY
    if not summary_path.exists() or json.loads(summary_path.read_text(encoding="utf-8")).get("overall") != "PASS":
        raise SystemExit(f"{stage} is not a run that passed every check")
    if open_workbooks():
        raise SystemExit(f"Close these workbooks in Excel first: {open_workbooks()}")
    return finish(stage, sorted(p for p in stage.iterdir() if p.is_file()), label)


if __name__ == "__main__":
    raise SystemExit(main())
