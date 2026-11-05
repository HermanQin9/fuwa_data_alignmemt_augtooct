"""Run the whole PMIX pipeline on a data export: tests, mapping, both reports and every check.

Everything is built in Result/_building/<run> first. Only when every step passes are the
current files in Result/ moved to versions/<date>-before-<label>/ and the new ones put in their
place, so a failed run never leaves Result/ half replaced.

    python mapping_first/run_pipeline.py --source "D:\\data\\USA PMix Nov.xlsx" --label nov-data

The previous mapping (default: the one in Result/, or a reviewed copy Cheena sends back) passes
its Approved and Excluded decisions to the new mapping. See README.md, "Running on new data".
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
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


def run_step(name: str, args: list, log: Path, steps: list) -> str:
    """Run one Python step, log its output, and stop the pipeline if it fails."""
    print(f"=== {name}", flush=True)
    started = time.time()
    done = subprocess.run([sys.executable, *map(str, args)], cwd=ROOT, capture_output=True,
                          text=True, encoding="utf-8", errors="replace")
    output = done.stdout + done.stderr
    with log.open("a", encoding="utf-8") as handle:
        handle.write(f"=== {name}\n{output}\n")
    seconds = round(time.time() - started)
    steps.append({"step": name, "result": "PASS" if done.returncode == 0 else "FAIL",
                  "seconds": seconds})
    if done.returncode != 0:
        print(output[-3000:], flush=True)
        raise StepFailed(f"{name} failed (exit {done.returncode}); full output in {log}")
    print(f"    PASS in {seconds}s", flush=True)
    return done.stdout


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
    if args.final and not (args.start and args.end):
        raise SystemExit("--final needs --start and --end")
    if bool(args.start) != bool(args.end):
        raise SystemExit("Give both --start and --end, or neither")
    locks = sorted(p.name for p in RESULT.glob("~$*"))
    if locks and not args.dry_run:
        raise SystemExit(f"Close these workbooks in Excel first: {locks}")


def archive_folder(label: str) -> Path:
    base = VERSIONS / f"{dt.date.today().isoformat()}-before-{label}"
    folder, n = base, 2
    while folder.exists():
        folder, n = base.with_name(f"{base.name}-{n}"), n + 1
    return folder


def publish(staged: list[Path], label: str) -> Path | None:
    """Move the current Result files to versions/, then the new ones into Result/."""
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
    parser.add_argument("--final", action="store_true",
                        help="Approved report: refuses Pending entries in the window")
    parser.add_argument("--dry-run", action="store_true",
                        help="Build and check in Result/_building only; leave Result/ unchanged")
    args = parser.parse_args()
    if args.no_previous:
        args.previous = None
    preflight(args)

    run = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    stage = STAGING / run
    stage.mkdir(parents=True)
    log = stage / "PIPELINE_RUN.log"
    suffix = "" if args.final else " DRAFT"
    mapping = stage / MAPPING_NAME
    reports = {layout: stage / f"{name}{suffix}.xlsx" for layout, name in REPORTS.items()}
    window = ["--start", args.start, "--end", args.end] if args.start else ["--all-dates"]
    mode = [] if args.final else ["--draft"]
    steps: list = []
    summary = {"run": run, "source": args.source.name, "reference": args.reference.name,
               "menu_list": args.menu_list.name,
               "previous_mapping": str(args.previous) if args.previous else None,
               "mode": "final" if args.final else "draft", "steps": steps}
    workflow = HERE / "pmix_workflow.py"
    try:
        run_step("unit tests", [HERE / "test_pmix_workflow.py"], log, steps)
        map_args = [workflow, "--source", args.source, "map", "--output", mapping,
                    "--menu-list", args.menu_list, "--reference", args.reference]
        if args.previous is not None:
            map_args += ["--previous", args.previous]
        summary["mapping"] = last_json(run_step("build mapping", map_args, log, steps))
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
    staged = [mapping, *reports.values(), *(stage / n for n in JSON_NAME.values()),
              stage / RUN_SUMMARY, log]
    print("\nAll steps passed.")
    print(json.dumps(summary.get("mapping", {}), indent=2))
    if args.dry_run:
        print(f"Dry run: output left in {stage}; Result/ unchanged.")
        return 0
    folder = publish(staged, args.label)
    stage.rmdir()
    if not any(STAGING.iterdir()):
        STAGING.rmdir()
    if folder:
        print(f"Previous Result files moved to {folder.relative_to(ROOT)}")
    print("New output is in Result/.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
