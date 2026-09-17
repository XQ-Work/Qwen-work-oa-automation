# -*- coding: utf-8 -*-
"""探针8：dump 甲方匹配行 <tr>/<td> 的 onclick，并试点行左边缘选中。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_pick3.txt"
FID = "field41324"
FULL = sys.argv[1] if len(sys.argv) > 1 else "荆州浦华荆清水务有限公司"
KW = FULL[:4]


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
            # dump 匹配行结构
            target_row = None
            for i, f in enumerate(page.frames):
                if f is frame:
                    continue
                try:
                    rows = f.locator("tr").filter(has_text=FULL)
                    for r in rows.all():
                        onc = r.get_attribute("onclick") or ""
                        lines.append(f"frame{i} <tr> onclick={onc[:70]!r}")
                        for td in r.locator("td").all():
                            t = (td.inner_text() or "").strip()[:14]
                            to = td.get_attribute("onclick") or ""
                            lines.append(f"    <td> text={t!r} onclick={to[:50]!r}")
                        if target_row is None:
                            target_row = (f, r)
                except Exception:
                    pass
            # 试点行左边缘(避开公司名链接)
            if target_row:
                f, r = target_row
                try:
                    r.click(position={"x": 2, "y": 6}, force=True, timeout=3000)
                    lines.append("clicked tr left-edge")
                except Exception as e:
                    lines.append(f"tr click fail {str(e)[:60]}")
                time.sleep(2.5)
                lines.append(f"after tr-click fieldval={fieldval(frame)!r}")
                page.screenshot(path=str(paths.STATE / "probe_pick3.png"), full_page=False)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
