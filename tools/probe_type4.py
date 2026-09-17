# -*- coding: utf-8 -*-
"""探针15：按 temptitle(标签) 定位乙方真正可见可输入框,打字测联想。窗口留60s便于观察。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

LABEL = sys.argv[1] if len(sys.argv) > 1 else "乙方名称"
FID = sys.argv[2] if len(sys.argv) > 2 else "field41325"
KW = sys.argv[3] if len(sys.argv) > 3 else "荆州浦华"
OUT = paths.STATE / "probe_type4.txt"


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
            # 列出表单帧里所有可见输入框 name/temptitle/bbox
            lines.append("== 表单帧可见输入框 ==")
            cand = None
            for el in frame.locator("input:visible").all():
                try:
                    nm = el.get_attribute("name") or ""
                    tt = el.get_attribute("temptitle") or ""
                    bb = el.bounding_box()
                    if bb and (LABEL in tt or "名称" in tt):
                        lines.append(f"  name={nm!r} temptitle={tt!r} bbox={ {k:round(v) for k,v in bb.items()} }")
                        if LABEL in tt and cand is None:
                            cand = el
                except Exception:
                    continue
            if cand is None:
                lines.append("未按 temptitle 找到乙方可见框")
            else:
                cand.click(timeout=5000)
                page.keyboard.type(KW, delay=130)
                lines.append(f"已对 temptitle={LABEL} 打 {KW}")
            time.sleep(3)
            page.screenshot(path=str(paths.STATE / "probe_type4.png"), full_page=False)
            # 联想命中
            hits = []
            for j, ff in enumerate(page.frames):
                try:
                    c = ff.locator("li:visible,div:visible,a:visible,td:visible").filter(
                        has_text=KW).count()
                    if c:
                        hits.append((j, c))
                except Exception:
                    pass
            lines.append(f"联想命中(KW可见元素帧:数)={hits[:12]}")
            try:
                lines.append(f"field41325 value={frame.locator(f'input[name={FID}]').first.input_value()!r}")
            except Exception:
                pass
            print("窗口保持60s，可观察/手动演示；期间我记录", flush=True)
            time.sleep(60)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            ctx.close()


if __name__ == "__main__":
    main()
