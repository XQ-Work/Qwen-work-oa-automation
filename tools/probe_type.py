# -*- coding: utf-8 -*-
"""探针2：测试"直接在字段输入框打简称→自动联想"是否可行(甲方 field41324)。
输出 -> oa_pilot/state/probe_type.txt + probe_type.png
"""
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_type.txt"
FID = sys.argv[1] if len(sys.argv) > 1 else "field41324"
KW = sys.argv[2] if len(sys.argv) > 2 else "荆州浦华"


def main():
    lines = []
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            base = auth.login(page)
            page.goto(base + fc.ADD, wait_until="domcontentloaded", timeout=40000)
            frame = fc.find_form_frame(page)
            lines.append(f"form={'OK' if frame else 'NONE'}")
            inp = frame.locator(f"input[name={FID}__]").first
            try:
                ro = inp.get_attribute("readonly")
                vt = inp.get_attribute("viewtype")
                cls = inp.get_attribute("class")
                lines.append(f"readonly={ro} viewtype={vt} class={cls}")
            except Exception as e:
                lines.append(f"attr err {str(e)[:80]}")
            # 尝试直接 fill
            try:
                inp.fill(KW)
                lines.append("fill OK")
            except Exception as e:
                lines.append(f"fill FAIL {str(e)[:100]}")
                try:
                    inp.click()
                    inp.type(KW, delay=80)
                    lines.append("type OK")
                except Exception as e2:
                    lines.append(f"type FAIL {str(e2)[:100]}")
            time.sleep(3)
            page.screenshot(path=str(paths.STATE / "probe_type.png"), full_page=False)
            # 扫描是否出现含关键词的联想项(新帧/新元素)
            hits = []
            for i, f in enumerate(page.frames):
                try:
                    n = f.locator(f"text={KW}").count()
                    if n:
                        hits.append((i, f.url[:24], n))
                except Exception:
                    pass
            lines.append(f"联想命中(KW出现处)={hits}")
            # 字段当前值
            try:
                lines.append(f"field value now={inp.input_value()!r}")
            except Exception:
                pass
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
