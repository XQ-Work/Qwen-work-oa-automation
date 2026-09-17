# -*- coding: utf-8 -*-
"""探针19：直接单测 browser_select 对乙方(干净页,唯一弹窗),看能否自动填出ID。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

FID = sys.argv[1] if len(sys.argv) > 1 else "field41325"
KW = sys.argv[2] if len(sys.argv) > 2 else "武汉汇谷自动化设备有限公司"
OUT = paths.STATE / "probe_bs.txt"


def main():
    lines = []
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        for pg in list(ctx.pages):
            try:
                pg.close()
            except Exception:
                pass
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            base = auth.login(page)
            page.goto(base + fc.ADD, wait_until="domcontentloaded", timeout=40000)
            frame = fc.find_form_frame(page)
            lines.append("调用 browser_select…")
            print("调用 browser_select…", flush=True)
            ok = fc.browser_select(page, frame, FID, KW, FID)
            try:
                v = frame.locator(f"input[name={FID}]").first.input_value()
            except Exception:
                v = "?"
            lines.append(f"ok={ok} {FID}={v!r}")
            print(f"ok={ok} {FID}={v!r}", flush=True)
            page.screenshot(path=str(paths.STATE / "probe_bs.png"), full_page=False)
            time.sleep(3)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            ctx.close()


if __name__ == "__main__":
    main()
