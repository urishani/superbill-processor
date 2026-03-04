"""
Superbill Processor
-------------------
UI tool to:
 - Search for an input Superbill XLS/XLSX file by name
 - Select an output XLS/XLSX file
 - Validate and drop expected empty columns
 - Append new rows (with column remapping) to the output file
 - Skip duplicate rows (identified by cols A, B, O in input / A, C, I in output)
"""

import tkinter as tk
from tkinter import ttk, filedialog
import os
import threading

import pandas as pd
import openpyxl
from openpyxl import load_workbook

# ── Constants ────────────────────────────────────────────────────────────────

# Header row inside the Superbill file (0-based row index)
INPUT_HEADER_ROW = 6

# Expected completely-empty column names in the input file.
# The raw Excel headers contain newlines, so we normalise by replacing \n with space.
EXPECTED_EMPTY_COLS_NORMALISED = {
    "Primary Carrier",
    "Primary Policy",
    "Secondary Carrier",
    "Secondary Policy",
    "Tertiary Carrier",
    "Tertiary Policy",
    "Clinical Trial",
    "Seq No",
    "Comment",
}

# Input column indices (0-based, after reading with header=INPUT_HEADER_ROW)
# A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 O=14 P=15 Q=16 R=17 S=18 T=19 U=20 V=21
# AD=29 AE=30 AF=31
COL_MAP_INPUT_TO_OUTPUT = {
    0:  0,   # A  → A
    1:  2,   # B  → C
    2:  3,   # C  → D
    3:  4,   # D  → E
    4:  5,   # E  → F
    5:  6,   # F  → G
    6:  7,   # G  → H
    14: 8,   # O  → I
    15: 9,   # P  → J
    16: 10,  # Q  → K
    17: 11,  # R  → L
    18: 12,  # S  → M
    19: 13,  # T  → N
    20: 14,  # U  → O
    21: 15,  # V  → P
    29: 23,  # AD → X
    30: 24,  # AE → Y
    31: 25,  # AF → Z
}

# Identity: input col indices [A, B, O] and output col indices [A, C, I]
IDENTITY_INPUT_COLS  = [0, 1, 14]   # A, B, O
IDENTITY_OUTPUT_COLS = [0, 2, 8]    # A, C, I

# ── Helpers ──────────────────────────────────────────────────────────────────

def col_letter(idx):
    """Return Excel-style column letter for 0-based index."""
    result = ""
    n = idx + 1
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def normalise_col(name):
    """Replace newlines/extra spaces in column name."""
    if not isinstance(name, str):
        return str(name)
    return " ".join(name.split())


# ── Core processing ──────────────────────────────────────────────────────────

def process(input_path, output_path, log):
    """
    Main processing function. `log` is a callable that appends a message to the UI.
    Returns True on success, False on failure/abort.
    """

    # ── 1. Read input ────────────────────────────────────────────────────────
    log(f"Reading input: {os.path.basename(input_path)}")
    try:
        df_in = pd.read_excel(input_path, header=INPUT_HEADER_ROW)
    except Exception as e:
        log(f"ERROR reading input file: {e}")
        return False

    # Normalise column names
    df_in.columns = [normalise_col(c) for c in df_in.columns]

    # Drop trailing all-NaN rows (footer artefacts)
    df_in.dropna(how="all", inplace=True)
    df_in.reset_index(drop=True, inplace=True)

    log(f"  Input rows (after cleanup): {len(df_in)}  |  Columns: {len(df_in.columns)}")

    # ── 2. Detect empty columns & compare with expected list ─────────────────
    actual_empty = {c for c in df_in.columns if df_in[c].isna().all()}

    missing_from_actual  = EXPECTED_EMPTY_COLS_NORMALISED - actual_empty   # expected empty but NOT empty
    extra_actual_empty   = actual_empty - EXPECTED_EMPTY_COLS_NORMALISED   # empty but NOT expected

    if missing_from_actual or extra_actual_empty:
        log("⚠  DISCREPANCY between expected empty columns and actual empty columns:")
        if missing_from_actual:
            log("   Columns expected to be empty but are NOT empty in this file:")
            for c in sorted(missing_from_actual):
                log(f"     • {c}")
        if extra_actual_empty:
            log("   Columns that are empty but were NOT in the expected list:")
            for c in sorted(extra_actual_empty):
                log(f"     • {c}")
        log("   (Processing continues with the actually-empty columns removed.)")
    else:
        log("✓  Empty columns match expected list exactly.")

    # Drop all actually-empty columns
    df_in.drop(columns=list(actual_empty), inplace=True, errors="ignore")
    log(f"  Columns after dropping empty: {len(df_in.columns)}")

    # Reload original (with all columns) for index-based mapping
    try:
        df_raw = pd.read_excel(input_path, header=INPUT_HEADER_ROW)
    except Exception as e:
        log(f"ERROR re-reading input file: {e}")
        return False
    df_raw.columns = [normalise_col(c) for c in df_raw.columns]
    df_raw.dropna(how="all", inplace=True)
    df_raw.reset_index(drop=True, inplace=True)

    # ── 3. Read / prepare output file ────────────────────────────────────────
    output_exists = os.path.isfile(output_path)
    if output_exists:
        log(f"Output file exists – loading: {os.path.basename(output_path)}")
        try:
            wb_out = load_workbook(output_path)
            ws_out = wb_out.active
        except Exception as e:
            log(f"ERROR reading output file: {e}")
            return False

        # Read existing data to check for duplicates
        # First row is assumed to be the header in the output file
        out_rows = list(ws_out.iter_rows(values_only=True))
        if len(out_rows) < 1:
            existing_keys = set()
            next_row = 1
        else:
            # Build set of (A, C, I) keys from existing data rows (skip header row)
            existing_keys = set()
            for r in out_rows[1:]:
                def _val(row, idx):
                    v = row[idx] if idx < len(row) else None
                    return str(v).strip() if v is not None else ""
                key = (
                    _val(r, IDENTITY_OUTPUT_COLS[0]),
                    _val(r, IDENTITY_OUTPUT_COLS[1]),
                    _val(r, IDENTITY_OUTPUT_COLS[2]),
                )
                existing_keys.add(key)
            next_row = ws_out.max_row + 1
    else:
        log("Output file does not exist – will be created.")
        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        # Write a simple header row derived from the mapping
        header = [""] * (max(COL_MAP_INPUT_TO_OUTPUT.values()) + 1)
        raw_cols = list(df_raw.columns)
        for in_idx, out_idx in COL_MAP_INPUT_TO_OUTPUT.items():
            if in_idx < len(raw_cols):
                header[out_idx] = raw_cols[in_idx]
        ws_out.append(header)
        existing_keys = set()
        next_row = 2

    # ── 4. Check for duplicate rows ──────────────────────────────────────────
    raw_cols = list(df_raw.columns)
    duplicates = []
    rows_to_add = []

    for _, row in df_raw.iterrows():
        def cell(col_idx):
            if col_idx >= len(raw_cols):
                return ""
            v = row.iloc[col_idx]
            return str(v).strip() if pd.notna(v) else ""

        key = (
            cell(IDENTITY_INPUT_COLS[0]),
            cell(IDENTITY_INPUT_COLS[1]),
            cell(IDENTITY_INPUT_COLS[2]),
        )
        if key in existing_keys:
            duplicates.append(key)
        else:
            rows_to_add.append(row)

    if duplicates:
        log("")
        log(f"⚠  {len(duplicates)} DUPLICATE row(s) already in output – skipping them (matched by Date / Patient / Billing Code):")
        log(f"    {'Date of Service':<15}  {'Patient Name':<30}  {'Billing Code'}")
        log(f"    {'-'*15}  {'-'*30}  {'-'*15}")
        for d in duplicates:
            log(f"    {d[0]:<15}  {d[1]:<30}  {d[2]}")
        log("")
    else:
        log("✓  No duplicate rows found.")

    if not rows_to_add:
        log("ℹ  Nothing new to add – all input rows already exist in the output.")
        return True

    # ── 5. Append rows ───────────────────────────────────────────────────────
    max_out_col = max(COL_MAP_INPUT_TO_OUTPUT.values()) + 1
    added = 0
    for row in rows_to_add:
        out_row = [None] * max_out_col
        for in_idx, out_idx in COL_MAP_INPUT_TO_OUTPUT.items():
            if in_idx < len(raw_cols):
                v = row.iloc[in_idx]
                out_row[out_idx] = None if pd.isna(v) else v
        ws_out.append(out_row)
        added += 1

    # ── 6. Save output ───────────────────────────────────────────────────────
    try:
        wb_out.save(output_path)
    except Exception as e:
        log(f"ERROR saving output file: {e}")
        return False

    log("")
    log(f"✅  Done.  {added} row(s) added to {os.path.basename(output_path)}.")
    return True


# ── UI ────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Superbill Processor")
        self.resizable(True, True)
        self.minsize(700, 500)
        self._build_ui()

    # ── layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = dict(padx=10, pady=4)

        # ── Input file section ────────────────────────────────────────────────
        frm_in = ttk.LabelFrame(self, text="Input Superbill file")
        frm_in.pack(fill="x", **pad)

        self.input_path_var = tk.StringVar()
        ttk.Label(frm_in, text="File:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(frm_in, textvariable=self.input_path_var, width=55).grid(
            row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(frm_in, text="Browse…", command=self._browse_input).grid(row=0, column=2, padx=4)
        frm_in.columnconfigure(1, weight=1)

        # ── Output file section ───────────────────────────────────────────────
        frm_out = ttk.LabelFrame(self, text="Output file")
        frm_out.pack(fill="x", **pad)

        self.output_path_var = tk.StringVar()
        ttk.Label(frm_out, text="File:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(frm_out, textvariable=self.output_path_var, width=55).grid(
            row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(frm_out, text="Open…", command=self._browse_output).grid(
            row=0, column=2, padx=4)
        frm_out.columnconfigure(1, weight=1)

        # ── Run button ────────────────────────────────────────────────────────
        self.run_btn = ttk.Button(self, text="▶  Process", command=self._run, width=20)
        self.run_btn.pack(pady=6)

        # ── Messages pane ─────────────────────────────────────────────────────
        frm_log = ttk.LabelFrame(self, text="Messages")
        frm_log.pack(fill="both", expand=True, **pad)

        self.log_text = tk.Text(frm_log, state="disabled", wrap="word",
                                font=("Consolas", 9), background="#1e1e1e",
                                foreground="#d4d4d4", insertbackground="white")
        self.log_text.pack(side="left", fill="both", expand=True)
        sb2 = ttk.Scrollbar(frm_log, orient="vertical", command=self.log_text.yview)
        sb2.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=sb2.set)

        ttk.Button(self, text="Clear log", command=self._clear_log).pack(anchor="e", padx=10, pady=(0, 6))

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select input Superbill file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.input_path_var.set(path)

    def _browse_output(self):
        path = filedialog.askopenfilename(
            title="Select existing output file to append to",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.output_path_var.set(path)

    # ── logging ───────────────────────────────────────────────────────────────

    def _log(self, msg):
        """Thread-safe log append."""
        self.after(0, self._log_sync, msg)

    def _log_sync(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ── run ───────────────────────────────────────────────────────────────────

    def _run(self):
        input_path  = self.input_path_var.get().strip()
        output_path = self.output_path_var.get().strip()

        if not input_path:
            self._log("⚠  Please select an input file first.")
            return
        if not os.path.isfile(input_path):
            self._log(f"⚠  Input file not found: {input_path}")
            return
        if not output_path:
            self._log("⚠  Please specify an output file.")
            return

        self.run_btn.configure(state="disabled")
        self._log("=" * 60)
        self._log("Starting processing…")

        def worker():
            try:
                process(input_path, output_path, self._log)
            finally:
                self.after(0, lambda: self.run_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
