# -*- coding: utf-8 -*-
"""探针13：乙方(field41325)字段框直接打简称能否联想匹配。
dump 该字段单元格 HTML + 可见可输入元素；往可见框打字看有无下拉。"""
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

FID = sys.argv[1] if len(sys.argv) > 1 else "field41325"
KW = sys.argv[2] if len(sys.argv) > 2 else "荆州浦华"
OUT = paths.STATE / "probe_type2.txt"


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
            # 该字段各候选输入框可见性
            for nm in (FID, FID + "__"):
                try:
                    tot = frame.locator(f"input[name={nm}]").count()
                    vis = frame.locator(f"input[name={nm}]:visible").count()
                    lines.append(f"input[{nm}] 总={tot} 可见={vis}")
                except Exception:
                    pass
            # dump 字段所在 td 的 HTML
            try:
                td = frame.locator(f"td[id={FID}_tdwrap], td:has(input[name='{FID}'])").first
                h = td.inner_html()
                lines.append("TD HTML:\n" + re.sub(r"\s+", " ", h)[:900])
            except Exception as e:
                lines.append(f"td dump err {str(e)[:80]}")
            # 找该字段附近可见可输入元素并打字
            typed = False
            for sel in (f"input[name={FID}__]:visible", f"input[name={FID}]:visible",
                        f"td[id={FID}_tdwrap] input:visible",
                        f"input.e8_browserInput:visible"):
                try:
                    loc = frame.locator(sel)
                    if loc.count() > 0:
                        loc.first.click(force=True, timeout=4000)
                        page.keyboard.type(KW, delay=120)
                        lines.append(f"typed into {sel!r}")
                        typed = True
                        break
                except Exception as e:
                    lines.append(f"type {sel!r} fail {str(e)[:60]}")
            if not typed:
                lines.append("无可打字可见框")
            time.sleep(3)
            page.screenshot(path=str(paths.STATE / "probe_type2.png"), full_page=False)
            # 联想下拉：含KW的可见元素(排除字段本身)
            hits = []
            for i, f in enumerate(page.frames):
                try:
                    c = f.locator("li:visible,div:visible,a:visible").filter(
                        has_text=KW).count()
                    if c:
                        hits.append((i, f.url[:20], c))
                except Exception:
                    pass
            lines.append(f"联想命中={hits[:10]}")
            try:
                lines.append(f"field value now={frame.locator(f'input[name={FID}]').first.input_value()!r}")
            except Exception:
                pass
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
