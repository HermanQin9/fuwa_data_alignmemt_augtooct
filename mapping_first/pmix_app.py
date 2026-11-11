"""Windows app for the PMIX Menu List report: pick a Raw data file and a date range, press Run.

It wraps run_pipeline.py (mapping, report, and every check) in a small tkinter window. Packaged
with build_app.py it becomes one .exe that runs from a single folder, see APP_GUIDE.txt.

From source:  python mapping_first/pmix_app.py
"""
from __future__ import annotations

import contextlib
import datetime as dt
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
from pathlib import Path

import run_pipeline

ROOT = run_pipeline.ROOT
RESULT = run_pipeline.RESULT
REFERENCE = "Menu Category reference.xlsx"
# Sheet -> the column that holds the sale date; used to suggest the date range.
DATE_COLUMNS = {"TOAST": "Sent Date", "SQUARE": "Date"}


def find_default_source() -> Path | None:
    """The newest 'Raw data*.xlsx' next to the app, if there is one."""
    found = sorted(ROOT.glob("Raw data*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return found[0] if found else None


def date_range(source: Path) -> tuple[dt.date, dt.date]:
    """First and last sale date in a Raw data workbook (read-only scan of the date columns)."""
    import openpyxl
    book = openpyxl.load_workbook(source, read_only=True, data_only=True)
    first = last = None
    try:
        for sheet, column in DATE_COLUMNS.items():
            if sheet not in book.sheetnames:
                continue
            rows = book[sheet].iter_rows(values_only=True)
            headers = list(next(rows, ()))
            if column not in headers:
                continue
            index = headers.index(column)
            for row in rows:
                value = row[index]
                if isinstance(value, dt.datetime):
                    value = value.date()
                if isinstance(value, dt.date):
                    first = value if first is None or value < first else first
                    last = value if last is None or value > last else last
    finally:
        book.close()
    if first is None:
        raise ValueError("No sale dates found. Check that the file has TOAST and SQUARE sheets.")
    return first, last


def pipeline_args(source: Path, start: str, end: str) -> list[str]:
    args = ["--source", str(source), "--reference", str(ROOT / REFERENCE), "--auto",
            "--skip-tests", "--start", start, "--end", end,
            "--label", f"run-{dt.date.today().isoformat()}"]
    if not (RESULT / run_pipeline.MAPPING_NAME).exists():
        args.append("--no-previous")
    return args


def run_report(argv: list[str]) -> int:
    """Run the pipeline with the given arguments; returns its exit code."""
    old = sys.argv
    sys.argv = ["run_pipeline", *argv]
    try:
        return run_pipeline.main()
    except SystemExit as stop:
        if stop.code not in (None, 0):
            print(stop.code)
        return 0 if stop.code is None else (stop.code if isinstance(stop.code, int) else 1)
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        sys.argv = old


def summary_text() -> str:
    """Plain-English result of the last run, from Result/PIPELINE_RUN.json."""
    try:
        run = json.loads((RESULT / run_pipeline.RUN_SUMMARY).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return "Your report is in the Result folder."
    pending = run.get("pending_entries", 0)
    if pending:
        return (
            f"The report is a DRAFT, because {pending} new item(s) do not have a category yet.\n\n"
            "To finish it:\n"
            "1. Open 'USA PMix - Menu Mapping.xlsx' in the Result folder.\n"
            "2. Go to the 'Menu Category' sheet and filter the 'Review Status' column for 'Pending'.\n"
            "3. For each new item, fill in the category and change 'Review Status' to 'Approved'.\n"
            "4. Save, close Excel, and press 'Build report' again.")
    return ("Your report is ready: 'USA PMix - Menu List Report.xlsx' in the Result folder.\n\n"
            "Every item has a category and every check passed.")


# The pipeline's step names, in plain English (all four run in the packaged app).
STEPS = {
    "build mapping": (1, "Matching every item to its menu list section"),
    "build menu-list report": (2, "Building the Excel report (this is the longest step)"),
    "validate menu-list report": (3, "Checking the report against the raw data"),
    "live Excel check, menu-list report": (4, "Testing the report in Excel"),
}


def excel_installed() -> bool:
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
        return True
    except OSError:
        return False


def main() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    window = tk.Tk()
    window.title("PMIX Menu List Report")
    window.geometry("760x600")
    window.minsize(640, 520)
    messages: queue.Queue = queue.Queue()
    state = {"running": False, "started": 0.0}

    source_var, start_var, end_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
    default = find_default_source()
    if default:
        source_var.set(str(default))

    class Writer:
        """Receives the pipeline's output: every line goes to the details box, and the step
        headers ('=== name') also move the plain-English progress line."""

        def __init__(self):
            self.pending = ""

        def write(self, text):
            if text:
                messages.put(("log", text))
                self.pending += text
                *lines, self.pending = self.pending.split("\n")
                for line in lines:
                    if line.startswith("=== ") and line[4:].strip() in STEPS:
                        messages.put(("step", line[4:].strip()))
            return len(text)

        def flush(self):
            pass

    frame = ttk.Frame(window, padding=16)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    frame.rowconfigure(10, weight=1)

    ttk.Label(frame, text="PMIX Menu List Report", font=("Segoe UI", 16, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w")
    ttk.Label(frame, justify="left", wraplength=700, text=(
        "1.  Choose your Raw data file (it is filled in if it is in this folder).\n"
        "2.  Check the start and end dates.\n"
        "3.  Press 'Build report' and keep this window open. A full quarter takes about "
        "15-20 minutes.")).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 12))

    ttk.Label(frame, text="Raw data file").grid(row=2, column=0, sticky="w")
    ttk.Entry(frame, textvariable=source_var).grid(row=2, column=1, sticky="ew", padx=8)
    browse = ttk.Button(frame, text="Browse...")
    browse.grid(row=2, column=2)

    ttk.Label(frame, text="Start date").grid(row=3, column=0, sticky="w", pady=(10, 0))
    ttk.Entry(frame, textvariable=start_var, width=14).grid(
        row=3, column=1, sticky="w", padx=8, pady=(10, 0))
    ttk.Label(frame, text="End date").grid(row=4, column=0, sticky="w", pady=(6, 0))
    ttk.Entry(frame, textvariable=end_var, width=14).grid(
        row=4, column=1, sticky="w", padx=8, pady=(6, 0))
    ttk.Label(frame, foreground="#555555", text="Type dates as year-month-day, for example "
              "2026-08-03. They are filled in from the file.").grid(
        row=5, column=1, columnspan=2, sticky="w", padx=8)

    run_button = ttk.Button(frame, text="Build report")
    run_button.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(14, 8), ipady=8)

    status = ttk.Label(frame, text="Ready.", wraplength=700)
    status.grid(row=7, column=0, columnspan=3, sticky="w")
    progress = ttk.Progressbar(frame, mode="indeterminate")
    progress.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(6, 0))

    tools = ttk.Frame(frame)
    tools.grid(row=9, column=0, columnspan=3, sticky="w", pady=(10, 0))
    open_button = ttk.Button(tools, text="Open Result folder",
                             command=lambda: os.startfile(RESULT) if RESULT.exists() else None)
    open_button.pack(side="left")
    details_button = ttk.Button(tools, text="Show details")
    details_button.pack(side="left", padx=8)
    copy_button = ttk.Button(tools, text="Copy details")
    copy_button.pack(side="left")

    details = tk.Text(frame, height=10, wrap="word", state="disabled", font=("Consolas", 9))
    scroll = ttk.Scrollbar(frame, command=details.yview)
    details.configure(yscrollcommand=scroll.set)
    shown = {"on": False}

    def toggle_details() -> None:
        if shown["on"]:
            details.grid_remove()
            scroll.grid_remove()
            details_button.configure(text="Show details")
        else:
            details.grid(row=10, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
            scroll.grid(row=10, column=2, sticky="nsw", pady=(10, 0))
            details_button.configure(text="Hide details")
        shown["on"] = not shown["on"]

    def copy_details() -> None:
        window.clipboard_clear()
        window.clipboard_append(details.get("1.0", "end"))
        status.configure(text="The details were copied. Paste them into a message to Herman.")

    details_button.configure(command=toggle_details)
    copy_button.configure(command=copy_details)

    def say(text: str) -> None:
        details.configure(state="normal")
        details.insert("end", text)
        details.see("end")
        details.configure(state="disabled")

    def tick() -> None:
        if state["running"]:
            seconds = int(time.time() - state["started"])
            current = status.cget("text").split("   (")[0]
            status.configure(text=f"{current}   ({seconds // 60} min {seconds % 60:02d} s)")
            window.after(1000, tick)

    def detect_dates() -> None:
        path = Path(source_var.get())
        if not path.is_file():
            return
        status.configure(text="Reading the dates from the file...")

        def work():
            try:
                messages.put(("dates", date_range(path)))
            except Exception as error:
                messages.put(("note", f"Could not read the dates ({error}). Type them in."))

        threading.Thread(target=work, daemon=True).start()

    def choose() -> None:
        picked = filedialog.askopenfilename(
            title="Choose the Raw data file", initialdir=str(ROOT),
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])
        if picked:
            source_var.set(picked)
            detect_dates()

    browse.configure(command=choose)

    def finish(code: int) -> None:
        state["running"] = False
        progress.stop()
        run_button.configure(state="normal")
        if code == 0:
            status.configure(text="Finished. All checks passed.")
            if messagebox.askyesno("PMIX report", summary_text() + "\n\nOpen the Result folder now?"):
                os.startfile(RESULT)
        else:
            status.configure(text="Something went wrong. No files were changed.")
            if not shown["on"]:
                toggle_details()
            lines = [x.strip() for x in details.get("1.0", "end").splitlines() if x.strip()]
            errors = [x for x in lines if re.match(r"^[A-Za-z]+(Error|Exception): ", x)]
            reason = errors[-1].split(": ", 1)[1] if errors else ""
            messagebox.showerror(
                "PMIX report",
                "Something went wrong, and none of your existing files were changed.\n\n"
                + (f"Reason: {reason}\n\n" if reason else "")
                + "Please press 'Copy details' and send the text to Herman.")

    def start() -> None:
        if state["running"]:
            return
        source = Path(source_var.get().strip().strip('"'))
        if not source.is_file():
            messagebox.showerror("PMIX report", "Please choose your Raw data file first "
                                 "(press 'Browse...').")
            return
        try:
            first, last = (dt.date.fromisoformat(v.get().strip()) for v in (start_var, end_var))
        except ValueError:
            messagebox.showerror("PMIX report", "Please type the dates as year-month-day, "
                                 "for example 2026-08-03.")
            return
        if first > last:
            messagebox.showerror("PMIX report", "The start date is after the end date.")
            return
        if not (ROOT / REFERENCE).is_file():
            messagebox.showerror("PMIX report", f"The file '{REFERENCE}' is missing from this "
                                 "folder. Please keep all the files together.")
            return
        if not excel_installed():
            messagebox.showerror("PMIX report", "Microsoft Excel was not found on this "
                                 "computer. The report needs Excel to be built and checked.")
            return
        if run_pipeline.open_workbooks():
            messagebox.showerror("PMIX report", "Please close these files in Excel first, then "
                                 "try again:\n\n" + "\n".join(run_pipeline.open_workbooks()))
            return
        state.update(running=True, started=time.time())
        run_button.configure(state="disabled")
        progress.start(12)
        status.configure(text="Starting...")
        say(f"\nRun started {dt.datetime.now():%Y-%m-%d %H:%M}\n")
        window.after(1000, tick)

        def work():
            writer = Writer()
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                code = run_report(pipeline_args(source, first.isoformat(), last.isoformat()))
            messages.put(("done", code))

        threading.Thread(target=work, daemon=True).start()

    run_button.configure(command=start)

    def poll() -> None:
        try:
            while True:
                kind, value = messages.get_nowait()
                if kind == "log":
                    say(value)
                elif kind == "step":
                    number, label = STEPS[value]
                    status.configure(text=f"Step {number} of {len(STEPS)}: {label}")
                elif kind == "dates":
                    start_var.set(value[0].isoformat())
                    end_var.set(value[1].isoformat())
                    status.configure(text="Ready. The dates come from the file; change them "
                                          "if you want a shorter period.")
                elif kind == "note":
                    status.configure(text=value)
                elif kind == "done":
                    finish(value)
        except queue.Empty:
            pass
        window.after(150, poll)

    def closing() -> None:
        if state["running"] and not messagebox.askyesno(
                "PMIX report", "A report is still being built. Close anyway?\n"
                "Your existing files will stay as they are."):
            return
        window.destroy()
        os._exit(0)

    window.protocol("WM_DELETE_WINDOW", closing)
    if default:
        detect_dates()
    poll()
    window.mainloop()


def headless(source: str, start: str, end: str) -> int:
    """No window: run once and keep the output in headless.log next to the app (for tests and scripting)."""
    with (ROOT / "headless.log").open("w", encoding="utf-8") as handle:
        with contextlib.redirect_stdout(handle), contextlib.redirect_stderr(handle):
            code = run_report(pipeline_args(Path(source), start, end))
            print(summary_text() if code == 0 else "FAILED")
    return code


if __name__ == "__main__":
    if len(sys.argv) == 5 and sys.argv[1] == "--headless":
        raise SystemExit(headless(*sys.argv[2:]))
    main()
