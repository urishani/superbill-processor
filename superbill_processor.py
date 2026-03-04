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
import shutil
import tempfile
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

def _last_nonempty_row(ws):
    """Return the 1-based index of the last row that has at least one non-None cell."""
    last = 0
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if any(v is not None for v in row):
            last = i
    return last


def process(input_path, output_path, log, confirm):
    """
    Main processing function.
      log(msg)        – append a line to the UI message pane
      confirm(msg)    – show a Proceed/Abort prompt; blocks until user responds;
                        returns True (proceed) or False (abort)
    Returns True on success/skip, False on abort/error.
    """

    # ── 1. Read input ────────────────────────────────────────────────────────
    log(f"Reading input: {os.path.basename(input_path)}")
    try:
        df_in = pd.read_excel(input_path, header=INPUT_HEADER_ROW)
    except Exception as e:
        log(f"ERROR reading input file: {e}")
        return False

    df_in.columns = [normalise_col(c) for c in df_in.columns]
    df_in.dropna(how="all", inplace=True)
    df_in.reset_index(drop=True, inplace=True)
    log(f"  Input rows (after cleanup): {len(df_in)}  |  Columns: {len(df_in.columns)}")

    # ── 2. Detect empty columns & compare with expected list ─────────────────
    actual_empty = {c for c in df_in.columns if df_in[c].isna().all()}
    missing_from_actual = EXPECTED_EMPTY_COLS_NORMALISED - actual_empty
    extra_actual_empty  = actual_empty - EXPECTED_EMPTY_COLS_NORMALISED

    if missing_from_actual or extra_actual_empty:
        log("⚠  DISCREPANCY between expected empty columns and actual empty columns:")
        if missing_from_actual:
            log("   Expected to be empty but are NOT empty:")
            for c in sorted(missing_from_actual):
                log(f"     • {c}")
        if extra_actual_empty:
            log("   Empty but NOT in the expected list:")
            for c in sorted(extra_actual_empty):
                log(f"     • {c}")
        log("   (Continuing – actually-empty columns will be removed.)")
    else:
        log("✓  Empty columns match expected list exactly.")

    # Reload original with all columns for index-based mapping
    try:
        df_raw = pd.read_excel(input_path, header=INPUT_HEADER_ROW)
    except Exception as e:
        log(f"ERROR re-reading input file: {e}")
        return False
    df_raw.columns = [normalise_col(c) for c in df_raw.columns]
    df_raw.dropna(how="all", inplace=True)
    df_raw.reset_index(drop=True, inplace=True)
    raw_cols = list(df_raw.columns)

    # ── 3. Load output file ───────────────────────────────────────────────────
    if not os.path.isfile(output_path):
        log("ERROR: output file does not exist. Please select an existing output file.")
        return False

    log(f"Loading output file: {os.path.basename(output_path)}")
    try:
        wb_out = load_workbook(output_path)
        ws_out = wb_out.active
    except Exception as e:
        log(f"ERROR reading output file: {e}")
        return False

    out_rows = list(ws_out.iter_rows(values_only=True))

    # Build identity key set from existing output rows (skip header row 0)
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

    # ── 4. Duplicate detection ────────────────────────────────────────────────
    duplicates   = []
    rows_to_add  = []

    for _, row in df_raw.iterrows():
        def cell(col_idx, _row=row):
            if col_idx >= len(raw_cols):
                return ""
            v = _row.iloc[col_idx]
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
        log(f"⚠  {len(duplicates)} duplicate row(s) already exist in the output:")
        log(f"    {'Date of Service':<15}  {'Patient Name':<30}  {'Billing Code'}")
        log(f"    {'-'*15}  {'-'*30}  {'-'*15}")
        for d in duplicates:
            log(f"    {d[0]:<15}  {d[1]:<30}  {d[2]}")
        log("")
        log(f"  {len(rows_to_add)} new row(s) would be appended.")
        if not confirm(f"{len(duplicates)} duplicate(s) found (listed above).\n"
                       f"Proceed with appending {len(rows_to_add)} new row(s)?"):
            log("⛔  Aborted by user. Output file was NOT modified.")
            return False
    else:
        log(f"✓  No duplicates found. {len(rows_to_add)} row(s) will be appended.")

    if not rows_to_add:
        log("ℹ  Nothing new to add – all input rows are already in the output.")
        return True

    # ── 5. Backup output file before writing ──────────────────────────────────
    backup_fd, backup_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(backup_fd)
    try:
        shutil.copy2(output_path, backup_path)
    except Exception as e:
        log(f"ERROR creating backup: {e}")
        return False

    # ── 6. Find last non-empty row and append after it ────────────────────────
    insert_after = _last_nonempty_row(ws_out)
    # Truncate any trailing empty rows beyond insert_after
    while ws_out.max_row > insert_after:
        ws_out.delete_rows(ws_out.max_row)

    log(f"  Appending after row {insert_after} (last non-empty row in output).")

    max_out_col = max(COL_MAP_INPUT_TO_OUTPUT.values()) + 1
    appended_keys = []
    for row in rows_to_add:
        out_row = [None] * max_out_col
        for in_idx, out_idx in COL_MAP_INPUT_TO_OUTPUT.items():
            if in_idx < len(raw_cols):
                v = row.iloc[in_idx]
                out_row[out_idx] = None if pd.isna(v) else v
        ws_out.append(out_row)
        # Track what we added by identity key
        def _cell(col_idx, _row=row):
            if col_idx >= len(raw_cols):
                return ""
            v = _row.iloc[col_idx]
            return str(v).strip() if pd.notna(v) else ""
        appended_keys.append((
            _cell(IDENTITY_INPUT_COLS[0]),
            _cell(IDENTITY_INPUT_COLS[1]),
            _cell(IDENTITY_INPUT_COLS[2]),
        ))

    # ── 7. Save to output file ────────────────────────────────────────────────
    try:
        wb_out.save(output_path)
    except Exception as e:
        log(f"ERROR saving output file: {e}")
        log("  Restoring backup…")
        shutil.copy2(backup_path, output_path)
        os.unlink(backup_path)
        return False

    log(f"  Saved. Verifying written data…")

    # ── 8. Verify: re-read output and confirm all appended rows are present ───
    try:
        wb_verify = load_workbook(output_path)
        ws_verify = wb_verify.active
        verified_keys = set()
        for r in ws_verify.iter_rows(values_only=True):
            def _vv(row, idx):
                v = row[idx] if idx < len(row) else None
                return str(v).strip() if v is not None else ""
            verified_keys.add((
                _vv(r, IDENTITY_OUTPUT_COLS[0]),
                _vv(r, IDENTITY_OUTPUT_COLS[1]),
                _vv(r, IDENTITY_OUTPUT_COLS[2]),
            ))
    except Exception as e:
        log(f"ERROR during verification read: {e}")
        verified_keys = set()

    missing_after_write = [k for k in appended_keys if k not in verified_keys]

    if missing_after_write:
        log("")
        log(f"⚠  VERIFICATION FAILED: {len(missing_after_write)} row(s) not found in output after save:")
        log(f"    {'Date of Service':<15}  {'Patient Name':<30}  {'Billing Code'}")
        log(f"    {'-'*15}  {'-'*30}  {'-'*15}")
        for k in missing_after_write:
            log(f"    {k[0]:<15}  {k[1]:<30}  {k[2]}")
        log("")
        if not confirm(f"Verification found {len(missing_after_write)} missing row(s).\n"
                       "Keep the (possibly incomplete) output, or Abort to restore the original?"):
            log("⛔  Aborted – restoring original output file from backup.")
            shutil.copy2(backup_path, output_path)
            os.unlink(backup_path)
            return False
        else:
            log("⚠  User chose to keep output despite verification discrepancy.")
    else:
        log(f"✓  Verification passed – all {len(appended_keys)} row(s) confirmed in output.")

    os.unlink(backup_path)
    log("")
    log(f"✅  Done.  {len(appended_keys)} row(s) added to {os.path.basename(output_path)}.")
    return True


# ── UI ────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Superbill Processor")
        self.resizable(True, True)
        self.minsize(700, 520)
        self._confirm_event = threading.Event()
        self._confirm_result = False
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

        # ── Confirm / Abort prompt (hidden until needed) ──────────────────────
        self.frm_confirm = ttk.LabelFrame(self, text="Confirmation required")
        # not packed yet – shown dynamically

        self.confirm_msg_var = tk.StringVar()
        ttk.Label(self.frm_confirm, textvariable=self.confirm_msg_var,
                  foreground="#cc8800", wraplength=640,
                  justify="left").pack(padx=8, pady=6, anchor="w")

        btn_row = ttk.Frame(self.frm_confirm)
        btn_row.pack(pady=(0, 8))
        ttk.Button(btn_row, text="✔  Proceed",
                   command=lambda: self._resolve_confirm(True),
                   width=16).pack(side="left", padx=12)
        ttk.Button(btn_row, text="✘  Abort",
                   command=lambda: self._resolve_confirm(False),
                   width=16).pack(side="left", padx=12)

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

        ttk.Button(self, text="Clear log", command=self._clear_log).pack(
            anchor="e", padx=10, pady=(0, 6))

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

    # ── confirmation helpers ──────────────────────────────────────────────────

    def _show_confirm(self, msg):
        """Called from the main thread to display the confirm panel."""
        self.confirm_msg_var.set(msg)
        self.frm_confirm.pack(fill="x", padx=10, pady=4)

    def _hide_confirm(self):
        self.frm_confirm.pack_forget()

    def _resolve_confirm(self, result: bool):
        """Called when user clicks Proceed or Abort."""
        self._hide_confirm()
        self._confirm_result = result
        self._confirm_event.set()

    def confirm(self, msg: str) -> bool:
        """
        Thread-safe blocking confirm. Called from the worker thread.
        Shows the prompt in the UI and waits for the user to click Proceed or Abort.
        """
        self._confirm_event.clear()
        self._confirm_result = False
        self.after(0, self._show_confirm, msg)
        self._confirm_event.wait()   # block worker thread
        return self._confirm_result

    # ── logging ───────────────────────────────────────────────────────────────

    def _log(self, msg):
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
        if not os.path.isfile(output_path):
            self._log(f"⚠  Output file not found: {output_path}")
            return

        self.run_btn.configure(state="disabled")
        self._log("=" * 60)
        self._log("Starting processing…")

        def worker():
            try:
                process(input_path, output_path, self._log, self.confirm)
            finally:
                self.after(0, lambda: self.run_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
