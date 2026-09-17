# -*- coding: utf-8 -*-
"""探针4：甲方弹窗里定位【可见】searchName，填词+点搜索，验证结果是否被过滤。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_search.txt"
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
            expr = fc._browser_expr(frame.content(), FID)
            frame.evaluate("() => { %s; }" % expr)
            time.sleep(4)
            # 1) 列出所有 searchName
            for i, f in enumerate(page.frames):
                try:
                    allsn = f.locator("input[name=searchName]")
                    n = allsn.count()
                    if n == 0:
                        continue
                    vis = f.locator("input[name=searchName]:visible").count()
                    title = ""
                    try:
                        title = allsn.first.get_attribute("title") or ""
                    except Exception:
                        pass
                    lines.append(f"[{i}] url={f.url[:30]} searchName n={n} vis={vis} title={title!r}")
                except Exception:
                    pass
            # 2) 找可见 searchName 填词
            filled = False
            for f in page.frames:
                try:
                    sn = f.locator("input[name=searchName]:visible")
                    if sn.count() > 0:
                        sn.first.fill(KW)
                        lines.append(f"fill visible searchName OK kw={KW}")
                        filled = True
                        # 点本帧或任意帧的 搜索
                        for g in [f] + [x for x in page.frames if x is not f]:
                            try:
                                g.get_by_text("搜索", exact=True).first.click(timeout=2000)
                                lines.append("click 搜索 OK")
                                break
                            except Exception:
                                continue
                        break
                except Exception:
                    continue
            if not filled:
                lines.append("无可见 searchName")
            time.sleep(3)
            page.screenshot(path=str(paths.STATE / "probe_search.png"), full_page=False)
            # 3) 统计结果行(含KW前缀的 <a>)
            for i, f in enumerate(page.frames):
                try:
                    c = f.locator("a").filter(has_text=KW).count()
                    if c:
                        lines.append(f"[{i}] 结果<a>含{KW} = {c}")
                except Exception:
                    pass
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
