"""
Download a Superbill export from Lynx (Playwright), then merge into the master workbook
using superbill_processor.process() (same backup/duplicate logic as the GUI).
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

from lynx_flow import FlowAborted, download_superbill, month_date_range
from superbill_processor import process


def _interactive_pause(message: str) -> None:
    """Wait for Enter in the terminal (for stepping through the run)."""
    input(f"{message}\nPress Enter to continue… ")


def _interactive_merge_or_skip() -> bool:
    """
    Ask whether to run superbill_processor on the downloaded file.
    Returns True to merge, False to exit without merging (download file is still discarded).
    """
    print("Download finished.")
    r = input("Merge into master workbook? [Y/n/abort]: ").strip().lower()
    if r in ("abort", "a", "q"):
        print("Aborted — merge not run.")
        return False
    if r in ("n", "no"):
        print("Skipping merge.")
        return False
    return True


def main() -> None:
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Fetch Lynx Superbill .xlsx and append into master via superbill_processor.",
        epilog=(
            "Requires LYNX_URL, LYNX_USER, LYNX_PASSWORD in .env. "
            "Interactive pauses are on by default; add --no-interactive for unattended runs. "
            "Example: python fetch_superbill_and_merge.py --month 03/2026 "
            '--master "data/Infusion Superbill.xlsx" --yes --no-interactive'
        ),
    )
    parser.add_argument(
        "--month",
        required=True,
        metavar="MM/YYYY",
        help="Report month (e.g. 03/2026). Lynx date range uses first-to-last day of that month.",
    )
    parser.add_argument(
        "--master",
        type=Path,
        required=True,
        help="Path to the existing master output workbook (.xlsx)",
    )
    parser.add_argument(
        "--download-dir",
        type=Path,
        default=Path(__file__).resolve().parent / "data",
        help="Directory to save the downloaded report (default: ./data)",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Auto-proceed on duplicate prompts during merge",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run browser headless (default off for easier debugging)",
    )
    parser.add_argument(
        "--slow-mo",
        type=int,
        default=0,
        metavar="MS",
        help="Slow down operations by this many milliseconds (0 = off)",
    )
    parser.add_argument(
        "--no-interactive",
        "--batch",
        "-b",
        action="store_false",
        dest="interactive",
        default=True,
        help=(
            "Skip terminal pauses (before browser, lynx_flow breakpoints, merge prompt). "
            "Default is interactive mode."
        ),
    )
    if len(sys.argv) <= 1:
        parser.print_help()
        sys.exit(0)
    args = parser.parse_args()

    if not (os.environ.get("LYNX_URL") or "").strip():
        print("ERROR: LYNX_URL must be set in .env or the environment.", file=sys.stderr)
        sys.exit(1)

    try:
        start_d, end_d = month_date_range(args.month)
    except ValueError as e:
        print(f"ERROR: invalid --month: {e}", file=sys.stderr)
        sys.exit(1)
    print(f"Superbill date range: {start_d} – {end_d}")

    master = args.master.resolve()
    if not master.is_file():
        print(f"ERROR: master file not found: {master}", file=sys.stderr)
        sys.exit(1)

    download_dir = args.download_dir.resolve()
    downloaded_path: Path | None = None

    try:
        if args.interactive:
            _interactive_pause("Next: Chromium will open and Lynx automation will run.")

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=args.headless, slow_mo=args.slow_mo)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            # Native window.alert / confirm / prompt (not custom HTML modals)
            page.on("dialog", lambda dialog: dialog.accept())
            try:
                saved = download_superbill(
                    page,
                    str(download_dir),
                    args.month,
                    interactive=args.interactive,
                )
                downloaded_path = Path(saved)
            except FlowAborted as e:
                print(f"Stopped: {e}", file=sys.stderr)
                sys.exit(130)
            finally:
                context.close()
                browser.close()

        if downloaded_path is None or not downloaded_path.is_file() or downloaded_path.stat().st_size == 0:
            print("ERROR: download did not produce a non-empty file.", file=sys.stderr)
            sys.exit(1)

        if args.interactive and not _interactive_merge_or_skip():
            sys.exit(0)

        def log(msg: str) -> None:
            print(msg)

        def confirm(msg: str) -> bool:
            if args.yes:
                return True
            return input(f"{msg}\nProceed? [y/N]: ").strip().lower() in ("y", "yes")

        ok = process(str(downloaded_path), str(master), log, confirm)
        sys.exit(0 if ok else 1)
    finally:
        pass


if __name__ == "__main__":
    main()
