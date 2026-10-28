r"""Playwright smoke test against a running Format Studio server.

Usage:
    .\.venv\Scripts\python.exe scripts\ui_smoke.py --url http://127.0.0.1:8765

Set FORMAT_STUDIO_EDGE to override the Edge executable path.
"""
from __future__ import annotations

import argparse
import os
import tempfile
from pathlib import Path

from playwright.sync_api import sync_playwright


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default="http://127.0.0.1:8765")
    parser.add_argument("--headless", action="store_true", default=True)
    args = parser.parse_args()
    edge = os.environ.get("FORMAT_STUDIO_EDGE", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
    root = Path(tempfile.mkdtemp(prefix="format-studio-ui-"))
    sample = root / "smoke.txt"
    sample.write_text("UI smoke test\n\nsecond paragraph\n", encoding="utf-8")
    errors: list[str] = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=edge, headless=args.headless)
        page = browser.new_page(viewport={"width": 1440, "height": 1000})
        page.on("console", lambda message: errors.append(message.text) if message.type == "error" else None)
        page.goto(args.url, wait_until="networkidle")
        page.get_by_text("\u8f6c\u6362\u5de5\u4f5c\u53f0").first.wait_for()
        page.locator("input[type=file]").first.set_input_files(str(sample))
        page.get_by_role("button", name="Markdown").first.click()
        page.get_by_role("button", name="\u5f00\u59cb\u8f6c\u6362").click()
        page.get_by_text("\u5df2\u5b8c\u6210").first.wait_for(timeout=30_000)
        page.goto(args.url.rstrip("/") + "/templates?tab=style", wait_until="networkidle")
        page.get_by_role("button", name="PDF \u6c34\u5370").click()
        assert page.get_by_text("\u542f\u7528 PDF \u6c34\u5370\uff08\u4ec5 PDF \u8f93\u51fa\uff09").count() == 1
        page.goto(args.url.rstrip("/") + "/settings", wait_until="networkidle")
        page.get_by_text("\u8f6c\u6362\u5f15\u64ce").first.wait_for()
        browser.close()
    if errors:
        raise SystemExit(f"console errors: {errors[:5]}")
    print("UI smoke test passed")


if __name__ == "__main__":
    main()