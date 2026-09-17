# -*- coding: utf-8 -*-
"""探针17(过程捕获)：保持窗口，用户操作乙方期间每秒记录 activeElement、
帧数(判断是否弹对话框)、乙方值；变非空后 dump 单元格并结束。"""
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
WAIT = int(sys.argv[2]) if len(sys.argv) > 2 else 120
OUT = paths.STATE / "probe_proc.txt"


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
            n0 = len(page.frames)
            lines.append(f"初始 n_frames={n0}，请现在操作乙方({FID})")
            print("请现在操作乙方", flush=True)
            end = time.time() + WAIT
            lastv = ""
            tick = 0
            while time.time() < end:
                # 用 document.title 侧信道读 activeElement(表单帧)
                try:
                    frame.evaluate(
                        "() => { try{var a=document.activeElement;"
                        "document.title='AE|'+(a?a.tagName:'#'+(a.id||'')+'.'+(a.className||''))"
                        "+'|nf='+parent.frames.length;}catch(e){} }")
                except Exception:
                    pass
                try:
                    title = page.title()
                except Exception:
                    title = "?"
                try:
                    v = frame.locator(f"input[name={FID}]").first.input_value()
                except Exception:
                    v = "?"
                nf = len(page.frames)
                lines.append(f"[{tick}s] nf={nf}(+{nf-n0}) title={title!r} {FID}={v!r}")
                if v and v != "?" and v != lastv:
                    page.screenshot(path=str(paths.STATE / "probe_proc.png"), full_page=False)
                    try:
                        td = frame.locator(f"td[id={FID}_tdwrap]").first
                        lines.append("乙方TD:\n" + re.sub(r"\s+", " ", td.inner_html())[:1000])
                    except Exception:
                        pass
                    lines.append("抓到非空值，结束")
                    break
                lastv = v
                tick += 1
                time.sleep(1.5)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
