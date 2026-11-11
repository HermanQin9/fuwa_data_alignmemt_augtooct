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
import sys
import threading
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
    """What the last run did, from Result/PIPELINE_RUN.json."""
    try:
        run = json.loads((RESULT / run_pipeline.RUN_SUMMARY).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return "The run finished. The reports are in the Result folder."
    pending = run.get("pending_entries", 0)
    lines = [f"Report type: {run.get('mode', '?').upper()}"]
    if pending:
        lines += [
            f"{pending} new item(s) have no decision yet (Review Status = Pending).",
            "Open Result\\USA PMix - Menu Mapping.xlsx, sheet Menu Category, filter Review Status "
            "for Pending, set the category and mark each row Approved, then run again for the "
            "final report.",
        ]
    else:
        lines.append("Every item has a decision. This is the final report.")
    return "\n".join(lines)


def main() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    window = tk.Tk()
    window.title("PMIX Menu List Report")
    window.geometry("760x520")
    messages: queue.Queue = queue.Queue()
    state = {"running": False}

    source_var, start_var, end_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
    default = find_default_source()
    if default:
        source_var.set(str(default))

    class Writer:
        def write(self, text):
            if text:
                messages.put(("log", text))
            return len(text)

        def flush(self):
            pass

    frame = ttk.Frame(window, padding=12)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    frame.rowconfigure(4, weight=1)

    ttk.Label(frame, text="Raw data file").grid(row=0, column=0, sticky="w")
    ttk.Entry(frame, textvariable=source_var).grid(row=0, column=1, sticky="ew", padx=6)
    browse = ttk.Button(frame, text="Browse...")
    browse.grid(row=0, column=2)

    ttk.Label(frame, text="Start date (YYYY-MM-DD)").grid(row=1, column=0, sticky="w", pady=6)
    ttk.Entry(frame, textvariable=start_var, width=14).grid(row=1, column=1, sticky="w", padx=6)
    ttk.Label(frame, text="End date (YYYY-MM-DD)").grid(row=2, column=0, sticky="w")
    ttk.Entry(frame, textvariable=end_var, width=14).grid(row=2, column=1, sticky="w", padx=6)

    buttons = ttk.Frame(frame)
    buttons.grid(row=3, column=0, columnspan=3, sticky="w", pady=8)
    run_button = ttk.Button(buttons, text="Run")
    run_button.pack(side="left")
    open_button = ttk.Button(buttons, text="Open Result folder",
                             command=lambda: os.startfile(RESULT) if RESULT.exists() else None)
    open_button.pack(side="left", padx=8)
    status = ttk.Label(buttons, text="Choose a file and dates, then press Run.")
    status.pack(side="left", padx=8)

    log = tk.Text(frame, height=16, wrap="word", state="disabled", font=("Consolas", 9))
    log.grid(row=4, column=0, columnspan=3, sticky="nsew")
    scroll = ttk.Scrollbar(frame, command=log.yview)
    scroll.grid(row=4, column=3, sticky="ns")
    log.configure(yscrollcommand=scroll.set)

    def say(text: str) -> None:
        log.configure(state="normal")
        log.insert("end", text)
        log.see("end")
        log.configure(state="disabled")

    def detect_dates() -> None:
        path = Path(source_var.get())
        if not path.is_file():
            return
        status.configure(text="Reading the date range...")

        def work():
            try:
                messages.put(("dates", date_range(path)))
            except Exception as error:
                messages.put(("note", f"Could not read dates: {error}"))

        threading.Thread(target=work, daemon=True).start()

    def choose() -> None:
        picked = filedialog.askopenfilename(
            title="Raw data file", initialdir=str(ROOT),
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])
        if picked:
            source_var.set(picked)
            detect_dates()

    browse.configure(command=choose)

    def finish(code: int) -> None:
        state["running"] = False
        run_button.configure(state="normal")
        if code == 0:
            status.configure(text="Done.")
            messagebox.showinfo("PMIX report", "All checks passed.\n\n" + summary_text())
        else:
            status.configure(text="Failed. Result folder not changed.")
            messagebox.showerror(
                "PMIX report",
                "The run stopped. The Result folder was not changed.\n"
                "The details are in the window; send them to Herman if you need help.")

    def start() -> None:
        if state["running"]:
            return
        source = Path(source_var.get())
        try:
            first, last = (dt.date.fromisoformat(v.get().strip()) for v in (start_var, end_var))
        except ValueError:
            messagebox.showerror("PMIX report", "Enter the dates as YYYY-MM-DD, for example 2026-08-03.")
            return
        if not source.is_file():
            messagebox.showerror("PMIX report", "Choose the Raw data file first.")
            return
        if first > last:
            messagebox.showerror("PMIX report", "The start date is after the end date.")
            return
        if not (ROOT / REFERENCE).is_file():
            messagebox.showerror("PMIX report", f"{REFERENCE} is missing from the app folder.")
            return
        if run_pipeline.open_workbooks():
            messagebox.showerror("PMIX report", "Close these files in Excel first:\n"
                                 + "\n".join(run_pipeline.open_workbooks()))
            return
        state["running"] = True
        run_button.configure(state="disabled")
        status.configure(text="Running. This takes a while; keep the window open.")
        say(f"\nRun started {dt.datetime.now():%Y-%m-%d %H:%M}\n")

        def work():
            with contextlib.redirect_stdout(Writer()), contextlib.redirect_stderr(Writer()):
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
                elif kind == "dates":
                    start_var.set(value[0].isoformat())
                    end_var.set(value[1].isoformat())
                    status.configure(text="Dates filled from the file; change them if needed.")
                elif kind == "note":
                    status.configure(text=value)
                elif kind == "done":
                    finish(value)
        except queue.Empty:
            pass
        window.after(150, poll)

    def closing() -> None:
        if state["running"] and not messagebox.askyesno(
                "PMIX report", "A run is in progress. Close anyway? The Result folder stays unchanged."):
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
