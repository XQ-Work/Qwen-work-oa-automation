# -*- coding: utf-8 -*-
"""级联探针第二轮：项目状态→项目类型→事项类型 全组合 dump（合同 476402）"""
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

ADD = ("/workflow/request/AddRequest.jsp?workflowid=476402"
       "&isagent=0&beagenter=0&f_weaver_belongto_userid=")
WATCH = {"41328": "项目状态", "41314": "项目类型", "41327": "事项类型"}


def dump_selects(frame):
    try:
        html = frame.content()
    except Exception:
        return {}
    by_name = {}
    for m in re.finditer(r"<select\b([^>]*)>(.*?)</select>", html, re.S):
        a = dict(re.findall(r'(\w+)\s*=\s*"([^"]*)"', m.group(1)))
        opts = [t.strip() for _, t in re.findall(
            r'<option[^>]*value="([^"]*)"[^>]*>([^<]*)</option>', m.group(2)) if t.strip()]
        by_name[a.get("name", "")] = opts
    return {label: by_name.get(f"field{fid}", []) for fid, label in WATCH.items()}


def main():
    result = []
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        base = auth.login(page)
        page.goto(base + ADD, wait_until="domcontentloaded", timeout=40000)
        time.sleep(10)
        frame = None
        for _ in range(15):
            for f in page.frames:
                try:
                    h = f.content()
                except Exception:
                    continue
                if h and h.count('name="field') > 50:
                    frame = f
                    break
            if frame:
                break
            time.sleep(2)
        if frame is None:
            print("[fail] 无表单帧")
            ctx.close()
            return 2
        # 前置链：水务板块→项目公司合同审批→事项审批任一支（项目状态依赖它）
        pre = [("field41343", "水务板块"), ("field41366", "水务项目公司合同审批"),
               ("field41448", "运营"), ("field41447", "污水"),
               ("field41367", "项目公司非长期固定采购生产物资")]
        for fid, lab in pre:
            try:
                frame.locator(f"select[name={fid}]:visible").first.select_option(label=lab)
                time.sleep(2.5)
            except Exception as e:
                print(f"[pre-fail {fid}={lab}] {str(e)[:80]}", flush=True)
        states = dump_selects(frame).get("项目状态", [])
        print(f"[项目状态全集] {states}", flush=True)
        for st in [s for s in states if s != "项目状态"]:
            try:
                frame.locator("select[name=field41328]:visible").first.select_option(label=st)
            except Exception as e:
                print(f"[skip {st}] {str(e)[:60]}", flush=True)
                continue
            time.sleep(3)
            snap = dump_selects(frame)
            types = snap.get("项目类型", [])
            result.append({"项目状态": st, "项目类型": types, "事项类型_by_项目类型": {}})
            print(f"[{st}] 项目类型={types}", flush=True)
            for ty in [t for t in types if t != "项目类型"]:
                try:
                    frame.locator("select[name=field41314]:visible").first.select_option(label=ty)
                    time.sleep(2.5)
                    items = dump_selects(frame).get("事项类型", [])
                    result[-1]["事项类型_by_项目类型"][ty] = items
                    print(f"   [{ty}] 事项类型={items}", flush=True)
                except Exception as e:
                    print(f"   [type-fail {ty}] {str(e)[:60]}", flush=True)
        (paths.STATE / "flowmap" / "合同_cascade2.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print("[ok] -> 合同_cascade2.json", flush=True)
        ctx.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
