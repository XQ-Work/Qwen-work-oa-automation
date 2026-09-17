# -*- coding: utf-8 -*-
"""探针：打开合同会签的甲方(field41324)放大镜，逐帧定位真正的搜索框。
输出 -> oa_pilot/state/probe_browser.txt（utf-8，避开控制台 GBK）
"""
import io
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_contract as fc  # noqa: E402
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = paths.STATE / "probe_browser.txt"
FID = sys.argv[1] if len(sys.argv) > 1 else "field41324"
KW = sys.argv[2] if len(sys.argv) > 2 else "荆州浦华荆清水务有限公司"


def census(page, form_frame, lines):
    lines.append(f"== n_frames={len(page.frames)} ==")
    for i, f in enumerate(page.frames):
        try:
            html = f.content()
        except Exception as e:
            lines.append(f"[{i}] url={f.url[:40]} ERR {str(e)[:40]}")
            continue
        ninp = len(re.findall(r"<input\b", html, re.I))
        nvis = f.locator("input:visible").count()
        ntd = f.locator("td").count()
        is_form = "YES" if f is form_frame else "no"
        has_sup = "供应商名称" in html
        has_q = "查询条件" in html
        has_search = ("搜索" in html) or ("查询" in html)
        lines.append(f"[{i}] form={is_form} url={f.url[:36]} "
                     f"input={ninp} vis={nvis} td={ntd} "
                     f"供应商名称={has_sup} 查询条件={has_q} 搜索={has_search}")
        if has_sup or has_q or ("请选择" in html):
            lines.append(f"  --- 对话框帧[{i}] 可见输入框 ---")
            try:
                for el in f.locator("input:visible").all():
                    try:
                        lines.append("    inp name=%r id=%r type=%r ph=%r cls=%r" % (
                            el.get_attribute("name"), el.get_attribute("id"),
                            el.get_attribute("type"), el.get_attribute("placeholder"),
                            el.get_attribute("class")))
                    except Exception:
                        pass
            except Exception as e:
                lines.append(f"    inp dump err {str(e)[:60]}")
            lines.append(f"  --- 对话框帧[{i}] 可点元素 ---")
            try:
                for sel in ("a:visible", "button:visible",
                            "input[type=button]:visible", "input[type=submit]:visible"):
                    for el in f.locator(sel).all():
                        try:
                            txt = (el.inner_text() or el.get_attribute("value") or "").strip()
                            oc = el.get_attribute("onclick") or ""
                            if txt or "earch" in oc:
                                lines.append("    %s text=%r onclick=%r" % (
                                    sel.split(":")[0], txt[:16], oc[:60]))
                        except Exception:
                            pass
            except Exception as e:
                lines.append(f"    btn dump err {str(e)[:60]}")
            m = re.search(r".{0,40}名称.{0,300}", html, re.S)
            if m:
                lines.append("    ~~~ " + re.sub(r"\s+", " ", m.group(0))[:360])


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
            lines.append(f"form_frame={'OK' if frame else 'NONE'} "
                         f"url={frame.url[:50] if frame else ''}")
            if not frame:
                return
            expr = fc._browser_expr(frame.content(), FID)
            lines.append(f"expr41324={'found' if expr else 'NONE'}")
            if expr:
                try:
                    frame.evaluate("() => { %s; }" % expr)
                except Exception as e:
                    lines.append(f"evaluate err {str(e)[:120]}")
                time.sleep(4)
                page.screenshot(path=str(paths.STATE / "probe_browser.png"),
                                full_page=False)
                census(page, frame, lines)
                # 尝试在所有帧里用"名称"标签定位输入框并填词
                for i, f in enumerate(page.frames):
                    try:
                        inp = f.locator("td:has-text('名称') + td input, "
                                        "label:has-text('名称') ~ input, "
                                        "input[placeholder*='名称']")
                        if inp.count() > 0:
                            inp.first.fill(KW)
                            lines.append(f"[fill] frame{i} via label-adjacent ok")
                            break
                    except Exception:
                        continue
                time.sleep(1)
        finally:
            OUT.write_text("\n".join(lines), encoding="utf-8")
            print("wrote", OUT)
            time.sleep(2)
            ctx.close()


if __name__ == "__main__":
    main()
