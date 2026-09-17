# -*- coding: utf-8 -*-
"""探针20：乙方行内联想——点框→打全称→退格删1字触发→等联想→点选→回读ID。"""
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
PRE = KW[:4]  # 删字后仍在前缀里，用于匹配下拉
OUT = paths.STATE / "probe_inline2.txt"


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
            n0 = len(page.frames)
            box = frame.locator(f"#innerContent{FID}div")
            box.first.click(force=True, timeout=5000)
            lines.append("clicked box")
            page.keyboard.type(KW, delay=90)
            page.keyboard.press("Backspace")   # 删1字触发联想
            time.sleep(1.8)
            page.screenshot(path=str(paths.STATE / "probe_inline2.png"), full_page=False)
            n1 = len(page.frames)
            lines.append(f"n_frames {n0}->{n1}")
            # 找联想下拉项(含前缀 PRE 的可见 li/div/a/td)
            found = None
            for j, f in enumerate(page.frames):
                try:
                    for sel in ("li", "div", "a", "td", "span"):
                        loc = f.locator(f"{sel}:visible").filter(has_text=PRE)
                        c = loc.count()
                        if c > 0:
                            lines.append(f"下拉候选 frame{j} {sel} n={c}")
                            if found is None and sel in ("li", "a", "td"):
                                found = (f, loc.first, sel)
                except Exception:
                    pass
            if found:
                f, el, sel = found
                try:
                    el.click(force=True, timeout=3000)
                    time.sleep(2)
                except Exception as e:
                    lines.append(f"点下拉失败 {str(e)[:60]}")
            try:
                v = frame.locator(f"input[name={FID}]").first.input_value()
            except Exception:
                v = "?"
            lines.append(f"RESULT {FID}={v!r}")
            print(f"RESULT {FID}={v!r}", flush=True)
            page.screenshot(path=str(paths.STATE / "probe_inline2_after.png"), full_page=False)
            time.sleep(30)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            ctx.close()


if __name__ == "__main__":
    main()
