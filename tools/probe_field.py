# -*- coding: utf-8 -*-
"""探针3：测试甲方可见输入框 field41324__ 打字是否触发联想(typeahead)。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_field.txt"
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
            nf = frame.locator(f"input[name={FID}__]:visible").count()
            lines.append(f"visible field{FID}__ count={nf}")
            if nf == 0:
                lines.append("字段可见框不存在→只能走弹窗")
                return
            el = frame.locator(f"input[name={FID}__]:visible").first
            el.click(timeout=8000)
            time.sleep(1.5)
            # 点击后是否弹出对话框(新帧)? 记录点击后帧数
            n_after_click = len(page.frames)
            lines.append(f"click后 n_frames={n_after_click}")
            # 逐字打
            page.keyboard.type(KW, delay=120)
            time.sleep(3)
            page.screenshot(path=str(paths.STATE / "probe_field.png"), full_page=False)
            # 找联想项：含关键词的可见 a/li/div(排除字段本身)
            hits = []
            for i, f in enumerate(page.frames):
                try:
                    c = f.locator(f"a:visible,li:visible,div:visible").filter(
                        has_text=KW).count()
                    if c:
                        hits.append((i, f.url[:22], c))
                except Exception:
                    pass
            lines.append(f"联想命中(KW可见元素)={hits[:12]}")
            try:
                lines.append(f"field value now={el.input_value()!r}")
            except Exception:
                pass
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
