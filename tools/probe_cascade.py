# -*- coding: utf-8 -*-
"""级联下拉探针：按路径逐级选择并 dump 各下拉的真实选项（合同会签 476402）"""
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

WF = "476402"
ADD = ("/workflow/request/AddRequest.jsp?workflowid=" + WF +
       "&isagent=0&beagenter=0&f_weaver_belongto_userid=")
WATCH = {"41366": "水务板块", "41306": "投资类型", "41367": "水务板块事项审批",
         "41328": "项目状态", "41314": "项目类型", "41327": "事项类型",
         "41448": "水务项目状态", "41447": "水务项目种类"}

# (字段, 选项文本) 选择序列；None 表示只 dump
STEPS = [
    (None, None),
    ("field41343", "水务板块"),
    (None, None),
    ("field41366", "水务项目公司合同审批"),
    (None, None),
    ("field41448", "运营"),
    ("field41447", "污水"),
    (None, None),
]


def dump_selects(frame):
    """CDP 导出 HTML 离线解析（帧 JS 环境不可靠，all_inner_texts 会撞注入冲突）"""
    import re
    try:
        html = frame.content()
    except Exception:
        return {label: "ERR content()" for label in WATCH.values()}
    by_name = {}
    for m in re.finditer(r"<select\b([^>]*)>(.*?)</select>", html, re.S):
        a = dict(re.findall(r'(\w+)\s*=\s*"([^"]*)"', m.group(1)))
        opts = [t.strip() for _, t in re.findall(
            r'<option[^>]*value="([^"]*)"[^>]*>([^<]*)</option>', m.group(2)) if t.strip()]
        by_name[a.get("name", "")] = opts
    return {label: by_name.get(f"field{fid}", []) for fid, label in WATCH.items()}


def main():
    log = []
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
                    html = f.content()
                except Exception:
                    continue
                if html and html.count('name="field') > 50:
                    frame = f
                    break
            if frame:
                break
            time.sleep(2)
        if frame is None:
            print("[fail] 未找到表单帧")
            ctx.close()
            return 2
        for sel_field, sel_label in STEPS:
            if sel_field:
                try:
                    frame.locator(f"select[name={sel_field}]:visible").first.select_option(
                        label=sel_label)
                    print(f"[select] {sel_field} = {sel_label}", flush=True)
                    time.sleep(3)
                except Exception as e:
                    print(f"[select-fail] {sel_field}={sel_label}: {str(e)[:100]}", flush=True)
            snap = dump_selects(frame)
            log.append({"after": f"{sel_field}={sel_label}" if sel_field else "init",
                        "options": snap})
            print(json.dumps(snap, ensure_ascii=False)[:600], flush=True)
        # 关键一步：遍历 水务板块事项审批 的每个分支，看子级选项联动
        try:
            real = dump_selects(frame).get("水务板块事项审批", [])
            real = [o for o in real if o != "水务板块事项审批"]
            for b in real:
                try:
                    frame.locator("select[name=field41367]:visible").first.select_option(label=b)
                except Exception as e:
                    print(f"[branch-skip {b}] {str(e)[:80]}", flush=True)
                    continue
                time.sleep(3)
                snap = dump_selects(frame)
                log.append({"after": f"field41367={b}", "options": snap})
                print(f"[branch {b}]", json.dumps(
                    {k: snap.get(k) for k in ("投资类型", "项目状态", "项目类型", "事项类型")},
                    ensure_ascii=False)[:400], flush=True)
        except Exception as e:
            print("[branch-fail]", str(e)[:150], flush=True)
        (paths.STATE / "flowmap" / "合同_cascade.json").write_text(
            json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
        page.screenshot(path=str(paths.STATE / "flowmap" / "合同_cascade.png"),
                        full_page=True)
        print("[ok] -> 合同_cascade.json")
        time.sleep(2)
        ctx.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
