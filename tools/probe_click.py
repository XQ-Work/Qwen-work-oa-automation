# -*- coding: utf-8 -*-
"""探针10：甲方——搜索后在弹窗帧里对结果行【真实鼠标点击】(非force),回读field41324。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_click.txt"
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
        for pg in ctx.pages[1:]:
            try:
                pg.close()
            except Exception:
                pass
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
                    continue
            time.sleep(3)
            # 候选帧：sframe 及其子帧(弹窗),排除主表单帧
            pool = list(fc._desc(sframe))
            pool += [f for f in page.frames if f is not frame and f not in pool]
            done = ""
            for f in pool:
                try:
                    a = f.locator("a").filter(has_text=FULL)
                    if a.count() == 0:
                        continue
                    bb = a.first.bounding_box()
                    lines.append(f"frame url={f.url[:26]} 结果<a>数={a.count()} bbox={bb}")
                    if not bb:
                        continue
                    # 真实鼠标点击行中央(避开链接文字:点 bbox 左侧 6px)
                    page.mouse.click(bb["x"] + 6, bb["y"] + bb["height"] / 2)
                    time.sleep(2)
                    done = fieldval(frame)
                    lines.append(f"real-click -> fieldval={done!r}")
                    if done:
                        break
                except Exception as e:
                    lines.append(f"click err {str(e)[:60]}")
                    continue
            page.screenshot(path=str(paths.STATE / "probe_click.png"), full_page=False)
            lines.append(f"FINAL fieldval={fieldval(frame)!r} url={page.url[:40]}")
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
