"""
Top-level UI for:
1) Fetching Lynx report into a work folder
2) Launching the existing merge UI (superbill_processor.py)
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from dotenv import dotenv_values, load_dotenv, set_key
from playwright.sync_api import sync_playwright

from lynx_flow import download_superbill, month_date_range


ROOT = Path(__file__).resolve().parent
ENV_PATH = ROOT / ".env"


def _env_get(key: str, default: str = "") -> str:
    return (os.environ.get(key) or default).strip()


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("Settings")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.url_var = tk.StringVar(value=_env_get("LYNX_URL"))
        self.user_var = tk.StringVar(value=_env_get("LYNX_USER"))
        self.pw_var = tk.StringVar(value=_env_get("LYNX_PASSWORD"))
        self.work_folder_var = tk.StringVar(value=_env_get("LYNX_WORK_FOLDER", "data"))

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Lynx URL").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.url_var, width=60).grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Label(frm, text="Lynx User").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.user_var, width=60).grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Label(frm, text="Lynx Password").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.pw_var, width=60, show="*").grid(row=2, column=1, sticky="ew", pady=4)

        ttk.Label(frm, text="Work Folder").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.work_folder_var, width=60).grid(row=3, column=1, sticky="ew", pady=4)

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="Save", command=self._save).pack(side="right", padx=4)

        frm.columnconfigure(1, weight=1)

    def _save(self) -> None:
        if not ENV_PATH.exists():
            ENV_PATH.write_text("", encoding="utf-8")
        set_key(str(ENV_PATH), "LYNX_URL", self.url_var.get().strip())
        set_key(str(ENV_PATH), "LYNX_USER", self.user_var.get().strip())
        set_key(str(ENV_PATH), "LYNX_PASSWORD", self.pw_var.get().strip())
        set_key(str(ENV_PATH), "LYNX_WORK_FOLDER", self.work_folder_var.get().strip() or "data")
        load_dotenv(ENV_PATH, override=True)
        self.destroy()


class WorkflowApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Superbill Fetch + Merge")
        self.geometry("820x500")
        self.minsize(760, 460)

        self.last_downloaded_path: str = ""
        self.fetch_running = False

        self._build_ui()
        self._refresh_defaults()

    def _build_ui(self) -> None:
        pad = dict(padx=10, pady=6)

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)
        ttk.Button(top, text="Settings", command=self._open_settings, width=14).pack(side="left")
        ttk.Label(
            top,
            text="Configure .env (URL/user/password/work folder), then fetch and launch merge UI.",
        ).pack(side="left", padx=10)

        fetch = ttk.LabelFrame(self, text="Fetch Stage")
        fetch.pack(fill="x", **pad)

        self.month_var = tk.StringVar()
        self.slow_mo_var = tk.IntVar(value=2000)
        self.headless_var = tk.BooleanVar(value=False)
        self.download_dir_var = tk.StringVar()

        ttk.Label(fetch, text="Month (MM/YYYY)").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(fetch, textvariable=self.month_var, width=16).grid(row=0, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(fetch, text="Slow-mo (ms)").grid(row=0, column=2, sticky="w", padx=16, pady=4)
        ttk.Entry(fetch, textvariable=self.slow_mo_var, width=10).grid(row=0, column=3, sticky="w", padx=4, pady=4)

        ttk.Checkbutton(fetch, text="Headless", variable=self.headless_var).grid(
            row=0, column=4, sticky="w", padx=16, pady=4
        )

        ttk.Label(fetch, text="Download folder").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(fetch, textvariable=self.download_dir_var, width=60).grid(
            row=1, column=1, columnspan=3, sticky="ew", padx=4, pady=4
        )
        ttk.Button(fetch, text="Browse…", command=self._browse_download_dir).grid(
            row=1, column=4, sticky="e", padx=6, pady=4
        )
        fetch.columnconfigure(3, weight=1)

        run_row = ttk.Frame(fetch)
        run_row.grid(row=2, column=0, columnspan=5, sticky="w", padx=6, pady=(6, 8))
        self.run_fetch_btn = ttk.Button(run_row, text="Run Fetch", command=self._run_fetch, width=18)
        self.run_fetch_btn.pack(side="left")
        self.fetch_status_var = tk.StringVar(value="Idle")
        ttk.Label(run_row, textvariable=self.fetch_status_var).pack(side="left", padx=12)

        result = ttk.LabelFrame(self, text="Fetch Result")
        result.pack(fill="x", **pad)
        self.result_name_var = tk.StringVar(value="Downloaded file: (none)")
        self.result_msg_var = tk.StringVar(value="Status: waiting")
        ttk.Label(result, textvariable=self.result_name_var).pack(anchor="w", padx=8, pady=3)
        ttk.Label(result, textvariable=self.result_msg_var).pack(anchor="w", padx=8, pady=3)

        merge = ttk.LabelFrame(self, text="Merge Stage")
        merge.pack(fill="x", **pad)
        ttk.Label(
            merge,
            text="Launch the existing merge UI from superbill_processor.py.",
        ).pack(anchor="w", padx=8, pady=4)
        ttk.Button(merge, text="Start Merge UI", command=self._open_merge_ui, width=18).pack(
            anchor="w", padx=8, pady=(0, 8)
        )

        log = ttk.LabelFrame(self, text="Messages")
        log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(log, height=10, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

    def _refresh_defaults(self) -> None:
        load_dotenv(ENV_PATH, override=True)
        month_hint = _env_get("LYNX_DEFAULT_MONTH")
        self.month_var.set(month_hint or self.month_var.get() or "")
        self.download_dir_var.set(_env_get("LYNX_WORK_FOLDER", "data"))

    def _open_settings(self) -> None:
        dlg = SettingsDialog(self)
        self.wait_window(dlg)
        self._refresh_defaults()
        self._log("Settings updated from .env")

    def _browse_download_dir(self) -> None:
        d = filedialog.askdirectory(title="Select download folder")
        if d:
            self.download_dir_var.set(d)

    def _open_merge_ui(self) -> None:
        target = ROOT / "superbill_processor.py"
        if not target.is_file():
            messagebox.showerror("Error", f"Cannot find merge UI script:\n{target}")
            return
        try:
            subprocess.Popen([sys.executable, str(target)], cwd=str(ROOT))
            self._log("Merge UI launched.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch merge UI:\n{e}")

    def _run_fetch(self) -> None:
        if self.fetch_running:
            return

        month = self.month_var.get().strip()
        try:
            month_date_range(month)
        except Exception as e:
            messagebox.showerror("Invalid month", str(e))
            return

        load_dotenv(ENV_PATH, override=True)
        url = _env_get("LYNX_URL")
        user = _env_get("LYNX_USER")
        pw = _env_get("LYNX_PASSWORD")
        if not url or not user or not pw:
            messagebox.showerror(
                "Missing settings",
                "Please set LYNX_URL, LYNX_USER, and LYNX_PASSWORD in Settings first.",
            )
            return

        download_dir = self.download_dir_var.get().strip() or _env_get("LYNX_WORK_FOLDER", "data")
        if not download_dir:
            download_dir = "data"
        slow_mo = max(0, int(self.slow_mo_var.get()))
        headless = bool(self.headless_var.get())

        self.fetch_running = True
        self.run_fetch_btn.configure(state="disabled")
        self.fetch_status_var.set("Running fetch...")
        self.result_msg_var.set("Status: running")
        self._log(f"Fetch start | month={month} headless={headless} slow_mo={slow_mo} download_dir={download_dir}")

        def worker() -> None:
            try:
                with sync_playwright() as p:
                    browser = p.chromium.launch(headless=headless, slow_mo=slow_mo)
                    context = browser.new_context(accept_downloads=True)
                    page = context.new_page()
                    page.on("dialog", lambda dialog: dialog.accept())
                    try:
                        saved = download_superbill(
                            page,
                            download_dir,
                            month,
                            interactive=False,
                        )
                    finally:
                        context.close()
                        browser.close()
                self.after(0, lambda: self._fetch_success(saved))
            except Exception as e:
                self.after(0, lambda: self._fetch_error(e, traceback.format_exc()))

        threading.Thread(target=worker, daemon=True).start()

    def _fetch_success(self, saved_path: str) -> None:
        self.fetch_running = False
        self.run_fetch_btn.configure(state="normal")
        self.fetch_status_var.set("Fetch finished successfully")
        self.last_downloaded_path = saved_path
        p = Path(saved_path)
        self.result_name_var.set(f"Downloaded file: {p.name}")
        self.result_msg_var.set(f"Status: success ({saved_path})")
        self._log(f"Fetch success: {saved_path}")

    def _fetch_error(self, err: Exception, tb: str) -> None:
        self.fetch_running = False
        self.run_fetch_btn.configure(state="normal")
        self.fetch_status_var.set("Fetch failed")
        self.result_msg_var.set(f"Status: error ({err})")
        self._log(f"Fetch error: {err}\n{tb}")

    def _log(self, msg: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


if __name__ == "__main__":
    load_dotenv(ENV_PATH, override=True)
    app = WorkflowApp()
    app.mainloop()
