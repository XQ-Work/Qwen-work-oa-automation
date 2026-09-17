# -*- coding: utf-8 -*-
"""探针14：在所有帧里找乙方输入框,force点击打简称,测联想下拉并选中回填。"""
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
OUT = paths.STATE / "probe_type3.txt"


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
            # 全帧定位乙方输入框(单/双下划线)
            target = None
            for i, f in enumerate(page.frames):
                for nm in (FID + "__", FID):
                    try:
                        loc = f.locator(f"input[name={nm}]")
                        n = loc.count()
                        if n == 0:
                            continue
                        vis = f.locator(f"input[name={nm}]:visible").count()
                        lines.append(f"frame{i} input[{nm}] n={n} vis={vis} url={f.url[:24]}")
                        if vis > 0 and target is None:
                            target = (f, loc.filter(visible=True).first)
                    except Exception:
                        pass
            if target is None:
                # 退而求其次:用 force 对任意 field41325__ 点击打字
                for i, f in enumerate(page.frames):
                    try:
                        loc = f.locator(f"input[name={FID}]").first
                        if loc.count() > 0:
                            target = (f, loc)
                            lines.append(f"fallback target frame{i}")
                            break
                    except Exception:
                        pass
            if target is None:
                lines.append("完全找不到乙方输入框")
                return
            f, el = target
            try:
                el.click(force=True, timeout=5000)
            except Exception as e:
                lines.append(f"click {str(e)[:60]}")
            page.keyboard.type(KW, delay=130)
            time.sleep(3)
            page.screenshot(path=str(paths.STATE / "probe_type3.png"), full_page=False)
            try:
                lines.append(f"typed value={el.input_value()!r}")
            except Exception:
                pass
            # 找联想下拉:打字后新出现的、含KW的可见 li/div/a(排除刚打的框)
            hits = []
            for j, ff in enumerate(page.frames):
                try:
                    for sel in ("li:visible", "div:visible", "a:visible", "span:visible"):
                        c = ff.locator(sel).filter(has_text=KW).count()
                        if c:
                            hits.append((j, sel.split(":")[0], c))
                except Exception:
                    pass
            lines.append(f"联想命中={hits[:12]}")
            # 试着点联想项
            for j, ff in enumerate(page.frames):
                try:
                    it = ff.locator("li:visible,div:visible,a:visible").filter(has_text=KW)
                    if it.count() > 0:
                        it.first.click(force=True, timeout=3000)
                        time.sleep(2)
                        try:
                            v = frame.locator(f"input[name={FID}]").first.input_value()
                        except Exception:
                            v = ""
                        lines.append(f"点联想项 frame{j} -> field41325={v!r}")
                        if v:
                            break
                except Exception:
                    continue
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
