#!/usr/bin/env python3
"""
FILE MERGER - FULLY WORKING VERSION
Built from scratch with tested merge functionality
"""

import os
import sys
import csv
import threading
import time
import random
import string
from datetime import datetime
from collections import OrderedDict
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import List, Optional

# ═══════════════════════════════════════════════════════════════
# Check pandas FIRST
# ═══════════════════════════════════════════════════════════════
try:
    import pandas as pd
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing Package",
        "pandas is required!\n\nOpen terminal and run:\npip install pandas openpyxl xlrd"
    )
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def get_downloads():
    """Get Downloads folder path."""
    home = os.path.expanduser("~")
    dl = os.path.join(home, "Downloads")
    if os.path.isdir(dl):
        return dl
    return home


def make_filename():
    """Generate: combined_file_20240115_143052_847291.csv"""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    rnd = ''.join(random.choices(string.digits, k=6))
    return f"combined_file_{ts}_{rnd}.csv"


def make_output_path():
    """Full path in Downloads."""
    return os.path.join(get_downloads(), make_filename())


def fmt_size(b):
    """Format bytes."""
    if b < 1024:
        return f"{b} B"
    elif b < 1024 * 1024:
        return f"{b / 1024:.1f} KB"
    elif b < 1024 * 1024 * 1024:
        return f"{b / (1024 * 1024):.1f} MB"
    else:
        return f"{b / (1024 * 1024 * 1024):.1f} GB"


def fmt_num(n):
    """Format number with commas."""
    return f"{n:,}"


def fmt_time(s):
    """Format seconds."""
    if s < 60:
        return f"{s:.1f} sec"
    m = int(s) // 60
    sc = int(s) % 60
    return f"{m}m {sc}s"


CSV_EXT = {".csv", ".tsv", ".txt"}
XL_EXT = {".xlsx", ".xls", ".xlsm", ".xlsb"}
ALL_EXT = CSV_EXT | XL_EXT
ENCODINGS = ["utf-8", "utf-8-sig", "cp1252", "latin1"]
DELIMS = [",", ";", "\t", "|"]
CHUNK = 200_000


def is_ok_file(path):
    """Check if file should be processed."""
    name = os.path.basename(path)
    if name.startswith(("~$", "._")):
        return False
    ext = os.path.splitext(path)[1].lower()
    return ext in ALL_EXT


def xl_engine(ext):
    """Get Excel engine."""
    if ext in {".xlsx", ".xlsm", ".xlsb"}:
        return "openpyxl"
    if ext == ".xls":
        return "xlrd"
    return None


# ═══════════════════════════════════════════════════════════════
# MERGE FUNCTION - The actual working merge logic
# ═══════════════════════════════════════════════════════════════

def do_merge(
    files: List[str],
    output: str,
    add_source: bool,
    rm_empty: bool,
    rm_dups: bool,
    progress_fn=None,
    status_fn=None,
    log_fn=None,
    cancel_check=None
) -> dict:
    """
    Merge files into one CSV. Returns stats dict.
    This is the CORE function that does all the work.
    """
    stats = {
        "ok": False, "total": len(files), "done": 0, "fail": 0,
        "rows": 0, "cols": 0, "sheets": 0,
        "empty_rm": 0, "dup_rm": 0,
        "in_size": 0, "out_size": 0,
        "time": 0, "output": output, "errors": []
    }

    def log(msg, lvl="info"):
        if log_fn:
            log_fn(msg, lvl)
        print(f"[{lvl.upper()}] {msg}")

    def status(msg):
        if status_fn:
            status_fn(msg)

    def progress(val, msg=""):
        if progress_fn:
            progress_fn(val, msg)

    def is_cancelled():
        if cancel_check:
            return cancel_check()
        return False

    t0 = time.time()

    try:
        # Input sizes
        for f in files:
            if os.path.exists(f):
                stats["in_size"] += os.path.getsize(f)

        log(f"Starting merge: {len(files)} files")
        log(f"Output: {output}")

        # ─── PHASE 1: Discover all columns ───
        status("Phase 1: Analyzing file columns...")
        col_map = OrderedDict()
        meta = ["_source_file", "_source_sheet"] if add_source else []

        for i, fpath in enumerate(files):
            if is_cancelled():
                log("Cancelled!", "warning")
                return stats

            pct = (i + 1) / (len(files) * 2) * 100
            fn = os.path.basename(fpath)
            progress(pct, f"Analyzing: {fn}")

            ext = os.path.splitext(fpath)[1].lower()
            found_cols = []

            try:
                if ext in CSV_EXT:
                    found_cols = _read_csv_header(fpath)
                    log(f"  CSV header: {fn} -> {len(found_cols)} cols")

                elif ext in XL_EXT:
                    eng = xl_engine(ext)
                    if eng:
                        try:
                            xf = pd.ExcelFile(fpath, engine=eng)
                            for sh in xf.sheet_names:
                                try:
                                    df = pd.read_excel(xf, sheet_name=sh, nrows=0, dtype=str)
                                    found_cols.extend(df.columns.tolist())
                                except Exception:
                                    pass
                            log(f"  Excel header: {fn} -> {len(found_cols)} cols, {len(xf.sheet_names)} sheets")
                        except Exception as e:
                            log(f"  Cannot open {fn}: {e}", "warning")

            except Exception as e:
                log(f"  Error analyzing {fn}: {e}", "warning")

            for c in found_cols:
                if c not in col_map:
                    col_map[c] = True

        all_cols = meta + list(col_map.keys())
        stats["cols"] = len(all_cols)
        log(f"Total unique columns: {len(all_cols)}")

        if not all_cols:
            log("ERROR: No columns found in any file!", "error")
            stats["errors"].append("No columns found")
            return stats

        # ─── PHASE 2: Write merged data ───
        status("Phase 2: Merging data...")

        # Delete existing output
        if os.path.exists(output):
            os.remove(output)

        header_written = False
        total_rows = 0

        with open(output, "w", encoding="utf-8", newline="") as outf:
            for i, fpath in enumerate(files):
                if is_cancelled():
                    log("Cancelled!", "warning")
                    return stats

                fn = os.path.basename(fpath)
                ext = os.path.splitext(fpath)[1].lower()
                pct = 50 + ((i + 1) / len(files)) * 45
                progress(pct, f"Merging: {fn}")
                status(f"Merging file {i + 1}/{len(files)}: {fn}")

                file_rows = 0

                try:
                    if ext in CSV_EXT:
                        file_rows = _write_csv(
                            fpath, fn, all_cols, outf,
                            header_written, add_source, rm_empty, stats
                        )

                    elif ext in XL_EXT:
                        file_rows = _write_excel(
                            fpath, fn, ext, all_cols, outf,
                            header_written, add_source, rm_empty, stats
                        )

                    if file_rows > 0:
                        header_written = True
                        total_rows += file_rows
                        stats["done"] += 1
                        log(f"✓ {fn}: {fmt_num(file_rows)} rows written", "success")
                    else:
                        log(f"⚠ {fn}: 0 rows (empty or unreadable)", "warning")

                except Exception as e:
                    stats["fail"] += 1
                    err = f"{fn}: {str(e)}"
                    stats["errors"].append(err)
                    log(f"✗ {err}", "error")

        stats["rows"] = total_rows

        # ─── PHASE 3: Deduplicate (optional) ───
        if rm_dups and total_rows > 0 and not is_cancelled():
            status("Phase 3: Removing duplicates...")
            progress(97, "Removing duplicates...")
            log("Removing duplicate rows...")

            try:
                df = pd.read_csv(output, dtype=str)
                before = len(df)
                df = df.drop_duplicates()
                after = len(df)
                stats["dup_rm"] = before - after
                stats["rows"] = after
                df.to_csv(output, index=False)
                log(f"Removed {fmt_num(before - after)} duplicates", "info")
            except Exception as e:
                log(f"Dedup warning: {e}", "warning")

        # ─── DONE ───
        stats["time"] = time.time() - t0
        if os.path.exists(output):
            stats["out_size"] = os.path.getsize(output)
        stats["ok"] = stats["done"] > 0

        progress(100, "Done!")
        status("✓ Merge completed!")
        log("─" * 50)
        log(f"✓ MERGE COMPLETE: {fmt_num(stats['rows'])} rows, "
            f"{stats['done']}/{stats['total']} files, "
            f"{fmt_time(stats['time'])}", "success")
        log("─" * 50)

    except Exception as e:
        stats["errors"].append(str(e))
        log(f"FATAL ERROR: {e}", "error")
        log(traceback.format_exc(), "error")

    return stats


def _read_csv_header(path):
    """Try to read CSV header."""
    for enc in ENCODINGS:
        for dlm in DELIMS:
            try:
                df = pd.read_csv(
                    path, nrows=0, dtype=str,
                    encoding=enc, sep=dlm,
                    engine="python", on_bad_lines="skip"
                )
                cols = [c for c in df.columns if not str(c).startswith("Unnamed")]
                if len(cols) > 0:
                    return cols
            except Exception:
                continue
    return []


def _write_csv(path, filename, all_cols, outf, hdr_done, add_src, rm_empty, stats):
    """Read CSV and write to output."""
    written = 0

    for enc in ENCODINGS:
        for dlm in DELIMS:
            try:
                reader = pd.read_csv(
                    path, dtype=str, encoding=enc, sep=dlm,
                    engine="python", on_bad_lines="skip",
                    chunksize=CHUNK, quoting=csv.QUOTE_MINIMAL
                )

                for chunk in reader:
                    orig = len(chunk)

                    if rm_empty:
                        chunk = chunk.dropna(how="all")
                        stats["empty_rm"] += orig - len(chunk)

                    if chunk.empty:
                        continue

                    if add_src:
                        chunk["_source_file"] = filename
                        chunk["_source_sheet"] = ""

                    # Align columns
                    for c in all_cols:
                        if c not in chunk.columns:
                            chunk[c] = ""
                    chunk = chunk[all_cols]

                    write_header = (not hdr_done and written == 0)
                    chunk.to_csv(outf, index=False, header=write_header,
                                 lineterminator="\n")
                    written += len(chunk)

                return written

            except Exception:
                continue

    return written


def _write_excel(path, filename, ext, all_cols, outf, hdr_done, add_src, rm_empty, stats):
    """Read Excel and write to output."""
    written = 0
    eng = xl_engine(ext)
    if not eng:
        return 0

    try:
        xf = pd.ExcelFile(path, engine=eng)

        for sheet in xf.sheet_names:
            try:
                df = pd.read_excel(xf, sheet_name=sheet, dtype=str)
                stats["sheets"] += 1

                orig = len(df)
                if rm_empty:
                    df = df.dropna(how="all")
                    stats["empty_rm"] += orig - len(df)

                if df.empty:
                    continue

                if add_src:
                    df["_source_file"] = filename
                    df["_source_sheet"] = sheet

                # Align columns
                for c in all_cols:
                    if c not in df.columns:
                        df[c] = ""
                df = df[all_cols]

                write_header = (not hdr_done and written == 0)
                df.to_csv(outf, index=False, header=write_header,
                          lineterminator="\n")
                written += len(df)

            except Exception as e:
                print(f"  Sheet error '{sheet}': {e}")

    except Exception as e:
        raise e

    return written


# ═══════════════════════════════════════════════════════════════
# GUI APPLICATION
# ═══════════════════════════════════════════════════════════════

import traceback


class MergerApp:
    """GUI Application."""

    def __init__(self):
        self.root = tk.Tk()
        self.files = []
        self.is_running = False
        self.cancel_flag = False
        self.thread = None

        self._build_gui()

    def start(self):
        """Launch the app."""
        self.root.mainloop()

    def _build_gui(self):
        """Build the entire GUI."""
        self.root.title("📊 File Merger Pro")
        self.root.configure(bg="#111827")

        # Window size
        w, h = 1050, 780
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")
        self.root.minsize(900, 650)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # ═══════════════════════════════════════════════════════
        # TOP BAR
        # ═══════════════════════════════════════════════════════

        top = tk.Frame(self.root, bg="#1f2937", height=55)
        top.pack(fill="x")
        top.pack_propagate(False)

        tk.Label(
            top, text="📊  File Merger Pro",
            font=("Arial", 18, "bold"),
            bg="#1f2937", fg="white"
        ).pack(side="left", padx=20, pady=10)

        tk.Button(
            top, text="  ✕  EXIT  ",
            font=("Arial", 11, "bold"),
            bg="#dc2626", fg="white",
            activebackground="#ef4444",
            activeforeground="white",
            relief="flat", bd=0,
            cursor="hand2",
            command=self._on_close
        ).pack(side="right", padx=20, pady=12)

        # ═══════════════════════════════════════════════════════
        # MAIN AREA - Two columns
        # ═══════════════════════════════════════════════════════

        main = tk.Frame(self.root, bg="#111827")
        main.pack(fill="both", expand=True, padx=15, pady=10)

        # Use grid for reliable layout
        main.grid_columnconfigure(0, weight=3)
        main.grid_columnconfigure(1, weight=2)
        main.grid_rowconfigure(0, weight=1)

        # ═══════════════════════════════════════════════════════
        # LEFT COLUMN - File Selection
        # ═══════════════════════════════════════════════════════

        left_col = tk.Frame(main, bg="#1e293b")
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        # Section header
        lhdr = tk.Frame(left_col, bg="#2563eb", height=45)
        lhdr.pack(fill="x")
        lhdr.pack_propagate(False)

        tk.Label(
            lhdr, text="  STEP 1 ─ Select Your Files",
            font=("Arial", 13, "bold"),
            bg="#2563eb", fg="white"
        ).pack(side="left", padx=12, pady=8)

        self.file_badge = tk.Label(
            lhdr, text=" 0 ",
            font=("Arial", 11, "bold"),
            bg="#111827", fg="white", padx=8
        )
        self.file_badge.pack(side="right", padx=12, pady=10)

        # Buttons
        btnrow = tk.Frame(left_col, bg="#1e293b")
        btnrow.pack(fill="x", padx=10, pady=10)

        tk.Button(
            btnrow, text="📄  ADD FILES",
            font=("Arial", 12, "bold"),
            bg="#2563eb", fg="white",
            activebackground="#3b82f6",
            activeforeground="white",
            relief="flat", bd=0,
            cursor="hand2",
            padx=18, pady=10,
            command=self._add_files
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btnrow, text="📂  ADD FOLDER",
            font=("Arial", 12, "bold"),
            bg="#7c3aed", fg="white",
            activebackground="#8b5cf6",
            activeforeground="white",
            relief="flat", bd=0,
            cursor="hand2",
            padx=18, pady=10,
            command=self._add_folder
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btnrow, text="🗑  CLEAR ALL",
            font=("Arial", 10, "bold"),
            bg="#4b5563", fg="white",
            activebackground="#dc2626",
            activeforeground="white",
            relief="flat", bd=0,
            cursor="hand2",
            padx=12, pady=10,
            command=self._clear_files
        ).pack(side="right")

        tk.Button(
            btnrow, text="➖  REMOVE",
            font=("Arial", 10, "bold"),
            bg="#4b5563", fg="white",
            activebackground="#dc2626",
            activeforeground="white",
            relief="flat", bd=0,
            cursor="hand2",
            padx=12, pady=10,
            command=self._remove_selected
        ).pack(side="right", padx=(0, 8))

        # Subfolder toggle
        self.use_subfolders = tk.BooleanVar(value=True)
        tk.Checkbutton(
            left_col, text="  Include subfolders when adding folder",
            variable=self.use_subfolders,
            font=("Arial", 10),
            bg="#1e293b", fg="#9ca3af",
            selectcolor="#111827",
            activebackground="#1e293b",
            activeforeground="white",
            cursor="hand2"
        ).pack(anchor="w", padx=10, pady=(0, 6))

        # File listbox
        list_container = tk.Frame(left_col, bg="#0f172a")
        list_container.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        scroll = tk.Scrollbar(list_container)
        scroll.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_container,
            font=("Consolas", 10),
            bg="#0f172a", fg="#d1d5db",
            selectbackground="#2563eb",
            selectforeground="white",
            highlightthickness=0,
            borderwidth=0,
            yscrollcommand=scroll.set,
            selectmode="extended",
            activestyle="none"
        )
        self.listbox.pack(fill="both", expand=True)
        scroll.config(command=self.listbox.yview)

        # Empty state
        self.empty_label = tk.Label(
            self.listbox,
            text="\n📂\n\nNo files added yet\n\n"
                 "Use the buttons above to\n"
                 "add files or a folder",
            font=("Arial", 12),
            bg="#0f172a", fg="#4b5563",
            justify="center"
        )
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

        # File summary
        self.info_label = tk.Label(
            left_col, text="Ready",
            font=("Arial", 10),
            bg="#1e293b", fg="#6b7280", anchor="w"
        )
        self.info_label.pack(fill="x", padx=10, pady=(0, 10))

        # ═══════════════════════════════════════════════════════
        # RIGHT COLUMN - Settings + Output + START
        # ═══════════════════════════════════════════════════════

        right_col = tk.Frame(main, bg="#1e293b")
        right_col.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        # Settings header
        rhdr = tk.Frame(right_col, bg="#7c3aed", height=45)
        rhdr.pack(fill="x")
        rhdr.pack_propagate(False)

        tk.Label(
            rhdr, text="  STEP 2 ─ Options",
            font=("Arial", 13, "bold"),
            bg="#7c3aed", fg="white"
        ).pack(side="left", padx=12, pady=8)

        # Options
        optf = tk.Frame(right_col, bg="#1e293b")
        optf.pack(fill="x", padx=15, pady=15)

        self.opt_source = tk.BooleanVar(value=True)
        tk.Checkbutton(
            optf, text="  Add source file/sheet columns",
            variable=self.opt_source,
            font=("Arial", 11),
            bg="#1e293b", fg="white",
            selectcolor="#111827",
            activebackground="#1e293b",
            activeforeground="white",
            cursor="hand2"
        ).pack(anchor="w", pady=6)

        self.opt_empty = tk.BooleanVar(value=True)
        tk.Checkbutton(
            optf, text="  Remove empty rows",
            variable=self.opt_empty,
            font=("Arial", 11),
            bg="#1e293b", fg="white",
            selectcolor="#111827",
            activebackground="#1e293b",
            activeforeground="white",
            cursor="hand2"
        ).pack(anchor="w", pady=6)

        self.opt_dups = tk.BooleanVar(value=False)
        tk.Checkbutton(
            optf, text="  Remove duplicate rows",
            variable=self.opt_dups,
            font=("Arial", 11),
            bg="#1e293b", fg="white",
            selectcolor="#111827",
            activebackground="#1e293b",
            activeforeground="white",
            cursor="hand2"
        ).pack(anchor="w", pady=6)

        # ─── Separator ───
        tk.Frame(right_col, bg="#374151", height=2).pack(fill="x", padx=15, pady=10)

        # ─── OUTPUT ───
        tk.Label(
            right_col, text="  STEP 3 ─ Output File",
            font=("Arial", 13, "bold"),
            bg="#1e293b", fg="white"
        ).pack(anchor="w", padx=15, pady=(5, 3))

        tk.Label(
            right_col, text="  Saves to Downloads by default:",
            font=("Arial", 10),
            bg="#1e293b", fg="#6b7280"
        ).pack(anchor="w", padx=15, pady=(0, 6))

        outrow = tk.Frame(right_col, bg="#1e293b")
        outrow.pack(fill="x", padx=15)

        self.out_var = tk.StringVar(value=make_output_path())
        self.out_entry = tk.Entry(
            outrow,
            textvariable=self.out_var,
            font=("Consolas", 9),
            bg="#0f172a", fg="#d1d5db",
            insertbackground="white",
            relief="flat",
            highlightthickness=2,
            highlightbackground="#374151",
            highlightcolor="#2563eb"
        )
        self.out_entry.pack(side="left", fill="x", expand=True, ipady=8)

        tk.Button(
            outrow, text="Browse",
            font=("Arial", 10),
            bg="#374151", fg="white",
            activebackground="#2563eb",
            activeforeground="white",
            relief="flat", bd=0,
            cursor="hand2",
            padx=12, pady=6,
            command=self._browse_output
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            right_col, text="🔄  New Filename",
            font=("Arial", 9),
            bg="#374151", fg="#9ca3af",
            activebackground="#2563eb",
            activeforeground="white",
            relief="flat", bd=0,
            cursor="hand2",
            padx=10, pady=3,
            command=lambda: self.out_var.set(make_output_path())
        ).pack(anchor="w", padx=15, pady=(6, 0))

        # ═══════════════════════════════════════════════════════
        # ███████████████████████████████████████████████████████
        #      STEP 4 - THE START BUTTON (ALWAYS VISIBLE)
        # ███████████████████████████████████████████████████████
        # ═══════════════════════════════════════════════════════

        tk.Frame(right_col, bg="#10b981", height=3).pack(fill="x", padx=15, pady=15)

        tk.Label(
            right_col,
            text="  🚀 STEP 4 ─ START MERGE",
            font=("Arial", 14, "bold"),
            bg="#1e293b", fg="#10b981"
        ).pack(anchor="w", padx=15, pady=(0, 10))

        # START BUTTON
        self.btn_start = tk.Button(
            right_col,
            text="\n▶▶  START MERGING FILES  ◀◀\n",
            font=("Arial", 17, "bold"),
            bg="#10b981",
            fg="white",
            activebackground="#34d399",
            activeforeground="white",
            relief="raised",
            bd=2,
            cursor="hand2",
            command=self._on_start_clicked
        )
        self.btn_start.pack(fill="x", padx=15, pady=(0, 5))
        self.btn_start.bind("<Enter>", lambda e: self.btn_start.config(bg="#34d399"))
        self.btn_start.bind("<Leave>", lambda e: self.btn_start.config(bg="#10b981"))

        # CANCEL BUTTON (same location, swapped during processing)
        self.btn_cancel = tk.Button(
            right_col,
            text="\n⏹  CANCEL PROCESSING\n",
            font=("Arial", 14, "bold"),
            bg="#dc2626",
            fg="white",
            activebackground="#ef4444",
            activeforeground="white",
            relief="raised",
            bd=2,
            cursor="hand2",
            command=self._on_cancel_clicked
        )
        # NOT packed yet - only shown during processing

        # ─── Progress ───
        progf = tk.Frame(right_col, bg="#1e293b")
        progf.pack(fill="x", padx=15, pady=(10, 10))

        self.status_text = tk.Label(
            progf,
            text="● Ready — Add files, then click START",
            font=("Arial", 10),
            bg="#1e293b", fg="#9ca3af",
            anchor="w"
        )
        self.status_text.pack(fill="x", pady=(0, 6))

        # Progress bar
        self.prog_outer = tk.Frame(progf, bg="#1f2937", height=24)
        self.prog_outer.pack(fill="x")
        self.prog_outer.pack_propagate(False)

        self.prog_inner = tk.Frame(self.prog_outer, bg="#10b981", width=0)
        self.prog_inner.place(x=0, y=0, relheight=1.0)

        self.prog_label = tk.Label(
            self.prog_outer, text="0%",
            font=("Arial", 9, "bold"),
            bg="#1f2937", fg="white"
        )
        self.prog_label.place(relx=0.5, rely=0.5, anchor="center")

        # ═══════════════════════════════════════════════════════
        # BOTTOM - LOG
        # ═══════════════════════════════════════════════════════

        btm = tk.Frame(self.root, bg="#1e293b")
        btm.pack(fill="x", padx=15, pady=(0, 10))

        btm_hdr = tk.Frame(btm, bg="#1f2937", height=32)
        btm_hdr.pack(fill="x")
        btm_hdr.pack_propagate(False)

        tk.Label(
            btm_hdr, text="  📋 Log",
            font=("Arial", 10, "bold"),
            bg="#1f2937", fg="white"
        ).pack(side="left", padx=8, pady=5)

        clr_btn = tk.Label(
            btm_hdr, text="Clear",
            font=("Arial", 9),
            bg="#1f2937", fg="#6b7280",
            cursor="hand2"
        )
        clr_btn.pack(side="right", padx=10, pady=5)
        clr_btn.bind("<Button-1>", lambda e: self._clear_log())

        log_scroll = tk.Scrollbar(btm)
        log_scroll.pack(side="right", fill="y")

        self.log_widget = tk.Text(
            btm,
            font=("Consolas", 10),
            bg="#0f172a", fg="#9ca3af",
            height=5,
            relief="flat",
            highlightthickness=0,
            yscrollcommand=log_scroll.set,
            state="disabled",
            wrap="word"
        )
        self.log_widget.pack(fill="x")
        log_scroll.config(command=self.log_widget.yview)

        # Log color tags
        self.log_widget.tag_configure("info", foreground="#9ca3af")
        self.log_widget.tag_configure("success", foreground="#10b981")
        self.log_widget.tag_configure("warning", foreground="#f59e0b")
        self.log_widget.tag_configure("error", foreground="#ef4444")

    # ═══════════════════════════════════════════════════════════
    # FILE OPERATIONS
    # ═══════════════════════════════════════════════════════════

    def _add_files(self):
        """Open file picker."""
        paths = filedialog.askopenfilenames(
            title="Select Files to Merge",
            filetypes=[
                ("Supported Files", "*.csv *.tsv *.xlsx *.xls *.xlsm"),
                ("CSV", "*.csv *.tsv"),
                ("Excel", "*.xlsx *.xls *.xlsm"),
                ("All", "*.*")
            ]
        )
        if paths:
            count = 0
            for p in paths:
                if p not in self.files:
                    self.files.append(p)
                    count += 1
            self._refresh_list()
            self._write_log(f"Added {count} file(s)")

    def _add_folder(self):
        """Open folder picker."""
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            found = []
            if self.use_subfolders.get():
                for root, dirs, fnames in os.walk(folder):
                    for fn in fnames:
                        fp = os.path.join(root, fn)
                        if is_ok_file(fp):
                            found.append(fp)
            else:
                for fn in os.listdir(folder):
                    fp = os.path.join(folder, fn)
                    if os.path.isfile(fp) and is_ok_file(fp):
                        found.append(fp)

            count = 0
            for p in sorted(found):
                if p not in self.files:
                    self.files.append(p)
                    count += 1

            self._refresh_list()
            if count:
                self._write_log(f"Added {count} file(s) from folder")
            else:
                self._write_log("No compatible files found in folder", "warning")

    def _remove_selected(self):
        """Remove selected files from list."""
        sel = list(self.listbox.curselection())
        if not sel:
            return
        for idx in reversed(sel):
            if idx < len(self.files):
                self.files.pop(idx)
        self._refresh_list()

    def _clear_files(self):
        """Clear all files."""
        self.files.clear()
        self._refresh_list()
        self._write_log("Cleared all files")

    def _refresh_list(self):
        """Refresh the file listbox."""
        self.listbox.delete(0, "end")

        if not self.files:
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
            self.file_badge.config(text=" 0 ")
            self.info_label.config(text="No files selected")
            return

        self.empty_label.place_forget()

        total_size = 0
        csv_n = 0
        xl_n = 0

        for p in self.files:
            name = os.path.basename(p)
            ext = os.path.splitext(p)[1].lower()
            sz = os.path.getsize(p) if os.path.exists(p) else 0
            total_size += sz

            if ext in CSV_EXT:
                icon = "📄"
                csv_n += 1
            else:
                icon = "📊"
                xl_n += 1

            self.listbox.insert("end", f"  {icon}  {name}  ({fmt_size(sz)})")

        self.file_badge.config(text=f" {len(self.files)} ")
        self.info_label.config(
            text=f"{len(self.files)} files  |  {fmt_size(total_size)}  |  CSV: {csv_n}  Excel: {xl_n}"
        )

    def _browse_output(self):
        """Browse for output location."""
        path = filedialog.asksaveasfilename(
            title="Save Merged File As",
            defaultextension=".csv",
            initialdir=get_downloads(),
            initialfile=make_filename(),
            filetypes=[("CSV", "*.csv"), ("All", "*.*")]
        )
        if path:
            self.out_var.set(path)

    # ═══════════════════════════════════════════════════════════
    # ████  THE START BUTTON HANDLER  ████
    # ═══════════════════════════════════════════════════════════

    def _on_start_clicked(self):
        """
        THIS IS THE MAIN START FUNCTION.
        Called when user clicks the green START button.
        """
        print("START BUTTON CLICKED!")  # Debug

        # ── Check 1: Any files? ──
        if not self.files:
            messagebox.showwarning(
                "No Files",
                "Please add files first!\n\n"
                "1. Click 'ADD FILES' to select files\n"
                "   - OR -\n"
                "2. Click 'ADD FOLDER' to select a folder"
            )
            return

        # ── Check 2: Output path ──
        output = self.out_var.get().strip()
        if not output:
            output = make_output_path()
            self.out_var.set(output)

        # ── Check 3: Output directory exists? ──
        out_dir = os.path.dirname(output)
        if out_dir and not os.path.exists(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create folder:\n{e}")
                return

        # ── Confirm ──
        n = len(self.files)
        ok = messagebox.askyesno(
            "Start Merge?",
            f"Merge {n} file(s) into one?\n\n"
            f"Output file:\n{os.path.basename(output)}\n\n"
            f"Saved to:\n{out_dir}\n\n"
            f"Click Yes to begin."
        )
        if not ok:
            return

        # ── Switch to processing mode ──
        self._write_log(f"Starting merge of {n} files...", "info")
        self.is_running = True
        self.cancel_flag = False

        # Hide START, show CANCEL
        self.btn_start.pack_forget()
        self.btn_cancel.pack(fill="x", padx=15, pady=(0, 5))

        # Reset progress
        self.prog_inner.place(x=0, y=0, relheight=1.0, relwidth=0)
        self.prog_label.config(text="0%")
        self.status_text.config(text="Starting merge...", fg="#fbbf24")

        # ── Launch worker thread ──
        self.thread = threading.Thread(
            target=self._worker,
            args=(output,),
            daemon=True
        )
        self.thread.start()

    def _worker(self, output):
        """Background worker that calls do_merge."""
        try:
            result = do_merge(
                files=self.files.copy(),
                output=output,
                add_source=self.opt_source.get(),
                rm_empty=self.opt_empty.get(),
                rm_dups=self.opt_dups.get(),
                progress_fn=lambda v, m: self.root.after(0, self._set_progress, v, m),
                status_fn=lambda m: self.root.after(0, self._set_status, m),
                log_fn=lambda m, l: self.root.after(0, self._write_log, m, l),
                cancel_check=lambda: self.cancel_flag
            )

            self.root.after(0, self._merge_finished, result)

        except Exception as e:
            tb = traceback.format_exc()
            self.root.after(0, self._merge_error, str(e), tb)

    def _on_cancel_clicked(self):
        """Cancel button handler."""
        if self.is_running:
            self.cancel_flag = True
            self._write_log("Cancelling...", "warning")

    # ═══════════════════════════════════════════════════════════
    # PROGRESS & STATUS
    # ═══════════════════════════════════════════════════════════

    def _set_progress(self, value, msg=""):
        """Update progress bar."""
        value = max(0, min(100, value))
        frac = value / 100.0
        self.prog_inner.place(x=0, y=0, relheight=1.0, relwidth=frac)
        self.prog_label.config(text=f"{value:.0f}%")
        if value >= 100:
            self.prog_label.config(bg="#10b981")
            self.prog_inner.config(bg="#10b981")
        if msg:
            self.status_text.config(text=msg)

    def _set_status(self, msg):
        """Update status label."""
        self.status_text.config(text=msg)

    def _restore_buttons(self):
        """Restore START button after processing."""
        self.is_running = False
        self.btn_cancel.pack_forget()
        self.btn_start.pack(fill="x", padx=15, pady=(0, 5))

    # ═══════════════════════════════════════════════════════════
    # MERGE COMPLETE / ERROR
    # ═══════════════════════════════════════════════════════════

    def _merge_finished(self, result):
        """Called when merge is done."""
        self._restore_buttons()

        if result["ok"]:
            self.status_text.config(text="✓ Merge completed!", fg="#10b981")
            self._show_results(result)
        else:
            self.status_text.config(text="✗ Merge failed", fg="#ef4444")
            if result["errors"]:
                messagebox.showerror("Failed", "\n".join(result["errors"]))

    def _merge_error(self, error, tb):
        """Called on fatal error."""
        self._restore_buttons()
        self.status_text.config(text="✗ Error occurred", fg="#ef4444")
        self._write_log(f"ERROR: {error}", "error")
        self._write_log(tb, "error")
        messagebox.showerror("Error", f"An error occurred:\n\n{error}")

    # ═══════════════════════════════════════════════════════════
    # RESULTS DIALOG
    # ═══════════════════════════════════════════════════════════

    def _show_results(self, res):
        """Show results popup."""
        dlg = tk.Toplevel(self.root)
        dlg.title("✓ Merge Complete")
        dlg.configure(bg="#1e293b")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        dw, dh = 520, 560
        x = self.root.winfo_x() + (self.root.winfo_width() - dw) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dh) // 2
        dlg.geometry(f"{dw}x{dh}+{x}+{y}")

        # Green success banner
        banner = tk.Frame(dlg, bg="#10b981", height=85)
        banner.pack(fill="x")
        banner.pack_propagate(False)

        tk.Label(
            banner, text="  ✓",
            font=("Arial", 40, "bold"),
            bg="#10b981", fg="white"
        ).pack(side="left", padx=20)

        bf = tk.Frame(banner, bg="#10b981")
        bf.pack(side="left", fill="both", expand=True, pady=15)

        tk.Label(
            bf, text="Merge Successful!",
            font=("Arial", 20, "bold"),
            bg="#10b981", fg="white", anchor="w"
        ).pack(fill="x")

        tk.Label(
            bf, text=f"{fmt_num(res['rows'])} rows from {res['done']} files",
            font=("Arial", 12),
            bg="#10b981", fg="white", anchor="w"
        ).pack(fill="x")

        # Stats table
        sf = tk.Frame(dlg, bg="#1e293b")
        sf.pack(fill="both", expand=True, padx=25, pady=20)

        data = [
            ("📁  Files Processed", f"{res['done']} / {res['total']}"),
            ("📊  Total Rows Written", fmt_num(res['rows'])),
            ("📋  Total Columns", str(res['cols'])),
            ("📑  Excel Sheets", str(res['sheets'])),
            ("🗑️  Empty Rows Removed", fmt_num(res['empty_rm'])),
            ("🔄  Duplicates Removed", fmt_num(res['dup_rm'])),
            ("📥  Input Size", fmt_size(res['in_size'])),
            ("📤  Output Size", fmt_size(res['out_size'])),
            ("⏱️  Processing Time", fmt_time(res['time'])),
        ]

        for label, value in data:
            r = tk.Frame(sf, bg="#1e293b")
            r.pack(fill="x", pady=4)

            tk.Label(
                r, text=label,
                font=("Arial", 11),
                bg="#1e293b", fg="#9ca3af",
                width=24, anchor="w"
            ).pack(side="left")

            tk.Label(
                r, text=value,
                font=("Arial", 11, "bold"),
                bg="#1e293b", fg="white",
                anchor="e"
            ).pack(side="right")

        # Output path
        pf = tk.Frame(dlg, bg="#0f172a")
        pf.pack(fill="x", padx=25, pady=(0, 15))

        tk.Label(
            pf, text="  Output saved to:",
            font=("Arial", 9),
            bg="#0f172a", fg="#6b7280"
        ).pack(anchor="w", padx=10, pady=(8, 0))

        tk.Label(
            pf, text=f"  {res['output']}",
            font=("Consolas", 9),
            bg="#0f172a", fg="#10b981",
            wraplength=460, anchor="w", justify="left"
        ).pack(anchor="w", padx=10, pady=(3, 8))

        # Action buttons
        abf = tk.Frame(dlg, bg="#1e293b")
        abf.pack(fill="x", padx=25, pady=(0, 20))

        def _open_folder():
            d = os.path.dirname(res['output'])
            try:
                if sys.platform == "win32":
                    os.startfile(d)
                elif sys.platform == "darwin":
                    os.system(f'open "{d}"')
                else:
                    os.system(f'xdg-open "{d}"')
            except Exception:
                pass

        def _open_file():
            try:
                if sys.platform == "win32":
                    os.startfile(res['output'])
                elif sys.platform == "darwin":
                    os.system(f'open "{res["output"]}"')
                else:
                    os.system(f'xdg-open "{res["output"]}"')
            except Exception:
                pass

        tk.Button(
            abf, text="📂 Open Folder",
            font=("Arial", 11),
            bg="#2563eb", fg="white",
            activebackground="#3b82f6",
            relief="flat", bd=0,
            cursor="hand2",
            padx=15, pady=8,
            command=_open_folder
        ).pack(side="left")

        tk.Button(
            abf, text="📄 Open File",
            font=("Arial", 11),
            bg="#7c3aed", fg="white",
            activebackground="#8b5cf6",
            relief="flat", bd=0,
            cursor="hand2",
            padx=15, pady=8,
            command=_open_file
        ).pack(side="left", padx=8)

        tk.Button(
            abf, text="  Close  ",
            font=("Arial", 11),
            bg="#4b5563", fg="white",
            activebackground="#6b7280",
            relief="flat", bd=0,
            cursor="hand2",
            padx=20, pady=8,
            command=dlg.destroy
        ).pack(side="right")

    # ═══════════════════════════════════════════════════════════
    # LOGGING
    # ═══════════════════════════════════════════════════════════

    def _write_log(self, msg, level="info"):
        """Write to log panel."""
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_widget.config(state="normal")
        self.log_widget.insert("end", f"[{ts}] {msg}\n", level)
        self.log_widget.see("end")
        self.log_widget.config(state="disabled")

    def _clear_log(self):
        """Clear log."""
        self.log_widget.config(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.config(state="disabled")

    # ═══════════════════════════════════════════════════════════
    # CLOSE
    # ═══════════════════════════════════════════════════════════

    def _on_close(self):
        """Handle window close."""
        if self.is_running:
            if messagebox.askyesno("Cancel?", "Merge in progress.\nCancel and exit?"):
                self.cancel_flag = True
                time.sleep(0.5)
                self.root.destroy()
        else:
            self.root.destroy()


# ═══════════════════════════════════════════════════════════════
# RUN
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 50)
    print("FILE MERGER PRO - Starting...")
    print("=" * 50)

    app = MergerApp()
    app.start()