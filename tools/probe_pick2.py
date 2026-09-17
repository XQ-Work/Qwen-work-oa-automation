# -*- coding: utf-8 -*-
"""探针7：甲方专属——搜索后点【供应商名称全称】的 <a>，回读字段值验证选中。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_pick2.txt"
FID = "field41324"
FULL = sys.argv[1] if len(sys.argv) > 1 else "荆州浦华荆清水务有限公司"
KW = FULL[:4]  # 简称搜索


def fieldval(frame):
    for nm in (FID, FID + "__"):
        try:
            v = frame.locator(f"input[name={nm}]").first.input_value()
            if v:
                return v
        except Exception:
            pass
    return ""


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
            lines.append(f"before fieldval={fieldval(frame)!r}")
            expr = fc._browser_expr(frame.content(), FID)
            frame.evaluate("() => { %s; }" % expr)
            sframe = None
            for _ in range(15):
                time.sleep(0.7)
                for f in page.frames:
                    try:
                        if f.locator("input[name=searchName]:visible").count() > 0:
                            sframe = f
                            break
                    except Exception:
                        pass
                if sframe:
                    break
            if not sframe:
                lines.append("无可见 searchName")
                return
            sframe.locator("input[name=searchName]:visible").first.fill(KW)
            for g in [sframe] + [x for x in page.frames if x is not sframe]:
                try:
                    g.get_by_text("搜索", exact=True).first.click(timeout=2000)
                    break
                except Exception:
                    pass
            time.sleep(3)
            # 列出含简称的 <a>
            for i, f in enumerate(page.frames):
                try:
                    for a in f.locator("a").filter(has_text=KW).all():
                        lines.append(f"frame{i} <a> text={a.inner_text()[:30]!r} "
                                     f"onclick={(a.get_attribute('onclick') or '')[:40]!r} "
                                     f"href={(a.get_attribute('href') or '')[:30]!r}")
                except Exception:
                    pass
            # 点【全称】那条 <a>
            clicked = False
            for f in page.frames:
                if f is frame:
                    continue
                try:
                    a = f.locator("a").filter(has_text=FULL)
                    if a.count() == 0:
                        a = f.locator("a").filter(has_text=KW)
                    if a.count() > 0:
                        a.first.click(force=True, timeout=3000)
                        lines.append(f"clicked <a> full={FULL!r}")
                        clicked = True
                        break
                except Exception:
                    continue
            time.sleep(2.5)
            page.screenshot(path=str(paths.STATE / "probe_pick2_after.png"), full_page=False)
            lines.append(f"clicked={clicked} after fieldval={fieldval(frame)!r} "
                         f"url={page.url[:50]}")
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
