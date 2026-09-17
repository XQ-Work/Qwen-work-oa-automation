# -*- coding: utf-8 -*-
"""探针18：测乙方"行内联想"——点 tabindex 框直接打字,看是行内下拉还是弹对话框。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

FID = sys.argv[1] if len(sys.argv) > 1 else "field41325"
KW = sys.argv[2] if len(sys.argv) > 2 else "武汉汇谷"
OUT = paths.STATE / "probe_inline.txt"


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
            # 点 tabindex 框
            box = frame.locator(f"#innerContent{FID}div")
            lines.append(f"box count={box.count()}")
            try:
                box.first.click(force=True, timeout=5000)
                lines.append("clicked tabindex box")
            except Exception as e:
                lines.append(f"click fail {str(e)[:80]}")
            time.sleep(1)
            page.keyboard.type(KW, delay=160)
            time.sleep(6)
            n1 = len(page.frames)
            lines.append(f"n_frames {n0} -> {n1} (弹对话框? {'是' if n1-n0>10 else '否/行内'})")
            page.screenshot(path=str(paths.STATE / "probe_inline.png"), full_page=False)
            # 行内下拉:含KW的可见元素
            hits = []
            for j, f in enumerate(page.frames):
                try:
                    for sel in ("li:visible", "div:visible", "a:visible", "td:visible"):
                        c = f.locator(sel).filter(has_text=KW).count()
                        if c:
                            hits.append((j, sel.split(":")[0], c))
                except Exception:
                    pass
            lines.append(f"含{KW}可见元素={hits[:12]}")
            try:
                lines.append(f"field41325 now={frame.locator(f'input[name={FID}]').first.input_value()!r}")
            except Exception:
                pass
            print("窗口保持40s可观察", flush=True)
            time.sleep(40)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            ctx.close()


if __name__ == "__main__":
    main()
