#!/usr/bin/env python3
"""Safe browser smoke test for a deployed Analitika instance.

Required environment variables:
  ANALITIKA_BASE_URL
  ANALITIKA_PASSWORD

The script logs in, opens every top-level module and its non-destructive
navigation controls, and fails on visible Streamlit exceptions. It never
submits warehouse writes or deletes data.
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

from playwright.sync_api import Page, sync_playwright

MODES = (
    "Обычный отчет",
    "Сравнение периодов",
    "Сувениры и касты на складе",
    "Заказ Sonu",
    "Заказ поставщику",
    "О программе",
)
WAREHOUSE_TABS = ("Главная", "Товары", "Поставки", "История", "Диагностика")


def _wait(page: Page) -> None:
    page.wait_for_timeout(900)
    page.wait_for_load_state("networkidle", timeout=30_000)


def _assert_clean(page: Page, context: str) -> None:
    errors = page.locator('[data-testid="stException"], text=/Traceback|Uncaught app exception/i')
    if errors.count():
        raise AssertionError(f"Streamlit exception in {context}: {errors.first.inner_text()[:1000]}")


def _click_text(page: Page, text: str) -> bool:
    locator = page.get_by_text(text, exact=True)
    if not locator.count():
        return False
    locator.first.click()
    _wait(page)
    return True


def main() -> int:
    base_url = os.getenv("ANALITIKA_BASE_URL", "").strip()
    password = os.getenv("ANALITIKA_PASSWORD", "")
    if not base_url or not password:
        print("Set ANALITIKA_BASE_URL and ANALITIKA_PASSWORD", file=sys.stderr)
        return 2
    artifacts = Path(os.getenv("ANALITIKA_E2E_ARTIFACTS", ".runtime/e2e"))
    artifacts.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 1100})
        page.goto(base_url, wait_until="networkidle", timeout=60_000)
        page.get_by_label("Пароль").fill(password)
        page.get_by_role("button", name="Войти").click()
        _wait(page)
        _assert_clean(page, "login")

        for mode in MODES:
            if not _click_text(page, mode):
                raise AssertionError(f"Top-level mode not found: {mode}")
            _assert_clean(page, mode)
            page.screenshot(path=str(artifacts / f"mode-{MODES.index(mode)+1}.png"), full_page=True)
            if mode == "Сувениры и касты на складе":
                for tab in WAREHOUSE_TABS:
                    if _click_text(page, tab):
                        _assert_clean(page, f"warehouse/{tab}")
                # Health refresh is read-only and validates all integrations.
                _click_text(page, "Проверить всё снова")
                _assert_clean(page, "warehouse/diagnostics-refresh")

        browser.close()
    print("E2E smoke passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
