# -*- coding: utf-8 -*-
"""通用捕获探针：打开某放大镜字段弹窗→等用户手动选一次→回读该字段值(记录ID)。
用法: python tools/probe_capture.py <fid> <显示名> [等待秒]
例:   python tools/probe_capture.py field41307 印章类型
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

FID = sys.argv[1] if len(sys.argv) > 1 else "field41307"
LABEL = sys.argv[2] if len(sys.argv) > 2 else "印章类型"
WAIT = int(sys.argv[3]) if len(sys.argv) > 3 else 180
OUT = paths.STATE / f"capture_{FID}.txt"


def val(frame, nm):
    try:
        return frame.locator(f"input[name={nm}]").first.input_value()
    except Exception:
        return ""


def main():
    lines = [f"fid={FID} label={LABEL}"]
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
            lines.append(f"expr={'found' if expr else 'NONE'}")
            if expr:
                try:
                    frame.evaluate("() => { %s; }" % expr)
                except Exception as e:
                    lines.append(f"open err {str(e)[:80]}")
            msg = f"请在窗口里手动选中【{LABEL}】，我在等…(最多{WAIT}s)"
            lines.append(msg)
            print(msg, flush=True)
            end = time.time() + WAIT
            got = ""
            while time.time() < end:
                got = val(frame, FID)
                if got:
                    break
                time.sleep(1.5)
            disp = val(frame, FID + "__")
            lines.append(f"RESULT {FID}={got!r} {FID}__={disp!r}")
            print(f"RESULT {FID}={got!r} disp={disp!r}", flush=True)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
