# -*- coding: utf-8 -*-
"""探针12：探 OA 供应商查询目录 /wui/main.jsp，看能否按名字读到供应商ID。"""
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

URL = "/wui/main.jsp"
KW = sys.argv[1] if len(sys.argv) > 1 else "供应商"
OUT = paths.STATE / "probe_supdir.txt"


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
        page = ctx.new_page()
        try:
            base = auth.login(page)
            page.goto(base + URL, wait_until="domcontentloaded", timeout=40000)
            time.sleep(4)
            # 关掉"您已经登录系统"弹窗
            for f in page.frames:
                try:
                    f.get_by_text("确定", exact=True).first.click(timeout=1500)
                except Exception:
                    continue
            time.sleep(2)
            lines.append(f"url_now={page.url}")
            page.screenshot(path=str(paths.STATE / "probe_supdir.png"), full_page=True)
            # 逐帧 dump 链接 + 命中关键词
            for i, f in enumerate(page.frames):
                try:
                    html = f.content() or ""
                except Exception:
                    continue
                hits = re.findall(r"<a\b[^>]*href=\"([^\"]*)\"[^>]*>([^<]{1,24})</a>", html)
                hits = [(h, t.strip()) for h, t in hits if t.strip()]
                key = [(h, t) for h, t in hits
                       if any(k in t for k in ("供应商", "查询", "高级搜索"))]
                if key:
                    lines.append(f"== frame{i} url={f.url[:40]} 命中链接 ==")
                    for h, t in key[:30]:
                        lines.append(f"  {t} -> {h[:90]}")
                elif hits:
                    lines.append(f"[frame{i} url={f.url[:34]} 链接数={len(hits)}] "
                                 + " | ".join(t for _, t in hits[:12]))
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
