"""
Lynx Superbill download. Credentials and base URL come from `.env` in the project root
(variables: LYNX_URL, LYNX_USER, LYNX_PASSWORD). Loaded automatically when this module is imported.

After download, Lynx may show "Access other apps and services on this device" — we click Allow by default.

Optional overrides (only if the default misses):
  LYNX_DOWNLOAD_CONFIRM_SELECTOR   CSS selector for the Allow control (highest precedence)
  LYNX_DOWNLOAD_CONFIRM_BUTTON     accessible name if not "Allow"

Interactive stepping is the default for fetch_superbill_and_merge.py; use --no-interactive
to turn it off. Add or remove _flow_pause(interactive, "label") in download_superbill() as needed.
"""

from __future__ import annotations

import calendar
import os
import re
from datetime import date
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import Page, expect

# Embed .env from the directory containing this file (project root)
_ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(_ENV_PATH)


class FlowAborted(Exception):
    """Raised when the user types abort at an interactive pause."""

    def __init__(self, step: str) -> None:
        self.step = step
        super().__init__(f"Aborted at step: {step}")


def _flow_pause(interactive: bool, step: str) -> None:
    """
    If interactive: wait in the terminal. Enter continues; 'abort' (or a/q) stops the flow.
    Edit this file to add or remove _flow_pause(...) calls where you want breakpoints.
    """
    if not interactive:
        return
    print(f"\n[lynx pause] {step}", flush=True)
    r = input("  [Enter] continue  |  type abort + Enter to stop: ").strip().lower()
    if r in ("abort", "a", "q"):
        raise FlowAborted(step)


def month_date_range(month_year: str) -> tuple[str, str]:
    """
    Parse mm/yyyy and return (start_date, end_date) as mm/dd/yyyy for the first and
    last calendar day of that month.
    """
    s = month_year.strip()
    m = re.fullmatch(r"(\d{1,2})/(\d{4})", s)
    if not m:
        raise ValueError("month must be mm/yyyy, e.g. 03/2026")
    month, year = int(m.group(1)), int(m.group(2))
    if not 1 <= month <= 12:
        raise ValueError("month must be between 01 and 12")
    last_d = calendar.monthrange(year, month)[1]
    start = date(year, month, 1)
    end = date(year, month, last_d)
    return (start.strftime("%m/%d/%Y"), end.strftime("%m/%d/%Y"))


_LYNX_APPS_PERMISSION_SNIPPET = "Access other apps and services on this device"


def _click_download_confirmation_if_configured(page: Page) -> None:
    """
    Clicks Allow on the Lynx permission prompt, or uses LYNX_DOWNLOAD_CONFIRM_* overrides.
    Order: env selector > env button name > built-in dialog + Allow.
    """
    sel = (os.environ.get("LYNX_DOWNLOAD_CONFIRM_SELECTOR") or "").strip()
    name = (os.environ.get("LYNX_DOWNLOAD_CONFIRM_BUTTON") or "").strip()
    if sel:
        page.locator(sel).first.click(timeout=15_000)
        return
    if name:
        page.get_by_role("button", name=name).click(timeout=15_000)
        return

    # Default: Lynx "Access other apps and services on this device" → Allow
    t = _LYNX_APPS_PERMISSION_SNIPPET
    for role in ("dialog", "alertdialog"):
        try:
            loc = page.get_by_role(role).filter(has_text=t)
            if loc.count() > 0:
                loc.first.get_by_role("button", name="Allow").click(timeout=15_000)
                return
        except Exception:
            continue
    try:
        page.locator("div").filter(has_text=t).first.get_by_role("button", name="Allow").click(
            timeout=15_000
        )
    except Exception:
        page.get_by_role("button", name="Allow").click(timeout=8_000)


def download_superbill(
    page: Page,
    download_path: str,
    month: str,
    *,
    interactive: bool = True,
) -> None:
    """
    Log into Lynx, run Superbill for the given month, save Excel to download_path.

    Parameters
    ----------
    page : Page
        Playwright page (caller uses accept_downloads=True).
    download_path : str
        Where to save the downloaded .xlsx.
    month : str
        Month as mm/yyyy; report covers the first through last day of that month
        (filled as mm/dd/yyyy in Start / End Date of Service).
    interactive : bool
        If True (default), pause at _flow_pause() breakpoints (terminal: Enter / abort).
    """
    url = (os.environ.get("LYNX_URL") or "").strip()
    user = (os.environ.get("LYNX_USER") or "").strip()
    password = (os.environ.get("LYNX_PASSWORD") or "").strip()
    if not url:
        raise RuntimeError("LYNX_URL is missing; set it in .env next to lynx_flow.py")
    if not user or not password:
        raise RuntimeError("LYNX_USER and LYNX_PASSWORD must be set in .env")

    start_str, end_str = month_date_range(month)

    page.goto(url)
    page.get_by_test_id("-input").fill(user)
    page.get_by_role("button", name="Next").click()
    page.get_by_test_id("-input").fill(password)
    page.get_by_role("button", name="Login").click()
    page.get_by_text("REPORTS").click()
    page.get_by_role("link", name="Superbill").click()
    # _flow_pause(interactive, "on Superbill screen (before date range)")
    page.get_by_role("textbox", name="Start Date of Service").click()
    page.get_by_role("textbox", name="Start Date of Service").press("ControlOrMeta+a")
    page.get_by_role("textbox", name="Start Date of Service").fill(start_str)
    page.get_by_role("textbox", name="End Date of Service").click()
    page.get_by_role("textbox", name="End Date of Service").press("ControlOrMeta+a")
    page.get_by_role("textbox", name="End Date of Service").fill(end_str)
    page.locator("#superbillForm").click()
    page.get_by_role("button", name="Run report").click()
    page.wait_for_timeout(1000)
    with page.expect_download() as download_info:
        page.get_by_role("button", name="DOWNLOAD EXCEL ALL").first.click()
    download = download_info.value
    print(f"[lynx] download_info.value: {download}", flush=True)
    try:
        print(f"[lynx] suggested_filename: {download.suggested_filename}", flush=True)
    except Exception:
        pass
    try:
        print(f"[lynx] url: {getattr(download, 'url', '')}", flush=True)
    except Exception:
        pass

    download.save_as(download_path)
    print(f"[lynx] download saved to {download_path}", flush=True)
    page.wait_for_timeout(1000)
    page.get_by_text("Log Out").click()


