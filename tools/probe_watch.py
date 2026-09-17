# -*- coding: utf-8 -*-
"""探针9(协作)：开甲方弹窗并搜好简称,然后轮询等待【用户手动点中】,
一旦 field41324 变非空,立即 dump 那一行结果帧的真实结构,供复刻。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_watch.txt"
FID = "field41324"
FULL = sys.argv[1] if len(sys.argv) > 1 else "荆州浦华荆清水务有限公司"
KW = FULL[:4]
WAIT_SECONDS = int(sys.argv[2]) if len(sys.argv) > 2 else 180


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
        # 关掉多余标签,只留一个,减少帧污染
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
            if sframe:
                try:
                    sframe.locator("input[name=searchName]:visible").first.fill(KW)
                    for g in [sframe] + [x for x in page.frames if x is not sframe]:
                        try:
                            g.get_by_text("搜索", exact=True).first.click(timeout=2000)
                            break
                        except Exception:
                            continue
                except Exception:
                    pass
            lines.append("窗口已就绪:甲方弹窗已搜'荆州浦华',请在该窗口手动点中它")
            print(lines[-1], flush=True)
            # 轮询等待用户点中
            end = time.time() + WAIT_SECONDS
            got = ""
            while time.time() < end:
                got = fieldval(frame)
                if got:
                    break
                time.sleep(1.5)
            lines.append(f"fieldval now={got!r}")
            # dump 含全称的 <a> 及其所在 tr 结构
            for i, f in enumerate(page.frames):
                try:
                    a = f.locator("a").filter(has_text=FULL)
                    if a.count() > 0:
                        html = a.first.evaluate("el=>el.closest('tr')?el.closest('tr').outerHTML:el.outerHTML")
                        lines.append(f"frame{i} 行HTML:\n{html[:1500]}")
                        break
                except Exception:
                    continue
            page.screenshot(path=str(paths.STATE / "probe_watch.png"), full_page=False)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
