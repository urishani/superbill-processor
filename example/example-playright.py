import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False, slow_mo=5000)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://example.com/")
    page.get_by_role("link", name="Learn more").click()
    page.get_by_role("link", name="IANA-managed Reserved Domains").click()
    page.get_by_role("cell", name="טעסט").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
