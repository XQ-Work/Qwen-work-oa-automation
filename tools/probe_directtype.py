# -*- coding: utf-8 -*-
"""探针5：不进弹窗，直接聚焦字段输入框打简称，测 e-cology 浏览器字段是否有联想(typeahead)。
用法: python tools/probe_directtype.py [field41324] [荆州浦华]
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_directtype.txt"
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
            # 候选输入框：单/双下划线，跨所有帧找可见的
            for name in (FID + "__", FID):
                for i, f in enumerate(page.frames):
                    try:
                        loc = f.locator(f"input[name={name}]")
                        n = loc.count()
                        if n == 0:
                            continue
                        vis = f.locator(f"input[name={name}]:visible").count()
                        lines.append(f"input[{name}] frame{i} n={n} vis={vis} url={f.url[:24]}")
                    except Exception:
                        pass
            # 尝试：在表单帧里对 field41324__ force 点击并打字
            tgt = frame.locator(f"input[name={FID}__]")
            lines.append(f"form-frame {FID}__ count={tgt.count()}")
            try:
                tgt.first.click(force=True, timeout=6000)
                lines.append("force click OK")
            except Exception as e:
                lines.append(f"force click FAIL {str(e)[:80]}")
            page.keyboard.type(KW, delay=120)
            time.sleep(3)
            page.screenshot(path=str(paths.STATE / "probe_directtype.png"), full_page=False)
            # 找联想下拉：含关键词的可见 li/div/a(排除表单帧本身字段)
            hits = []
            for i, f in enumerate(page.frames):
                try:
                    for sel in ("li:visible", "div:visible", "a:visible", "td:visible"):
                        c = f.locator(sel).filter(has_text=KW).count()
                        if c:
                            hits.append((i, sel.split(":")[0], c))
                except Exception:
                    pass
            lines.append(f"联想命中={hits[:15]}")
            try:
                lines.append(f"字段值 now={tgt.first.input_value()!r}")
            except Exception:
                pass
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
