# -*- coding: utf-8 -*-
"""通用发起页诊断工具：抓任意 workflowid 的表单结构（只读，不填不存）

用法：python tools/diag_flow.py --workflowid 444401 --label 供应商
输出：state/flowmap/{label}_fields.json + {label}_form.html
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from oa_common import paths, auth  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

ADD_PATH = ("/workflow/request/AddRequest.jsp"
            "?workflowid={wf}&isagent=0&beagenter=0&f_weaver_belongto_userid=")


def parse_fields(html):
    """离线解析表单控件：input(空元素)/textarea/select 分别匹配，按位置排序"""
    matches = []  # (pos, tag, attrs, inner)
    for m in re.finditer(r'<input\b([^>]*)>', html, re.I):
        matches.append((m.start(), "input", m.group(1), ""))
    for m in re.finditer(r'<textarea\b([^>]*)>(.*?)</textarea>', html, re.S | re.I):
        matches.append((m.start(), "textarea", m.group(1), m.group(2)))
    for m in re.finditer(r'<select\b([^>]*)>(.*?)</select>', html, re.S | re.I):
        matches.append((m.start(), "select", m.group(1), m.group(2)))
    matches.sort(key=lambda x: x[0])

    fields = []
    for pos, tag, attr_str, inner in matches:
        a = dict(re.findall(r'(\w[\w-]*)\s*=\s*"([^"]*)"', attr_str))
        name = a.get("name", "")
        if not name.startswith("field"):
            continue
        info = {"name": name, "id": a.get("id", ""), "tag": tag,
                "type": a.get("type", "").lower(), "value": a.get("value", "")[:80],
                "cls": a.get("class", "")[:50]}
        if tag == "select":
            info["options"] = re.findall(
                r'<option[^>]*value="([^"]*)"[^>]*>([^<]{0,30})', inner)[:30]
        ctx = html[max(0, pos - 1500):pos]
        cells = re.findall(r"<td[^>]*>(.*?)</td>", ctx, re.S)
        label = ""
        for c in reversed(cells):
            t = re.sub(r"<[^>]+>", "", c)
            t = re.sub(r"\s+", "", t).replace("&nbsp;", "")
            if t:
                label = t[:16]
                break
        info["label"] = label
        fields.append(info)
    seen, uniq = set(), []
    for f in fields:
        if f["name"] in seen:
            continue
        seen.add(f["name"])
        uniq.append(f)
    return uniq


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workflowid", required=True)
    ap.add_argument("--label", required=True)
    args = ap.parse_args()
    out = paths.STATE / "flowmap"
    out.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        base = auth.login(page)
        page.goto(base + ADD_PATH.format(wf=args.workflowid),
                  wait_until="domcontentloaded", timeout=40000)
        time.sleep(10)

        best = None
        for attempt in range(10):
            cand = None
            for f in page.frames:
                try:
                    html = f.content()
                except Exception:
                    continue
                n = html.count('name="field') if html else 0
                if n >= 3 and (cand is None or n > cand[0]):
                    cand = (n, f, html)
            if cand:
                best = cand
                print(f"[hit] 表单帧 field控件x{cand[0]} url={cand[1].url[:90]}", flush=True)
                break
            print(f"[scan {attempt}] 未找到表单帧（{len(page.frames)}帧，CDP普查）", flush=True)
            time.sleep(3)

        if not best:
            page.screenshot(path=str(out / f"{args.label}_fail.png"))
            print("[fail] 未定位表单frame", flush=True)
            ctx.close()
            return 2

        n, frame, html = best
        (out / f"{args.label}_form.html").write_text(html, encoding="utf-8")
        fields = parse_fields(html)
        (out / f"{args.label}_fields.json").write_text(
            json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")
        page.screenshot(path=str(out / f"{args.label}.png"), full_page=True)
        print(f"[ok] HTML {len(html)}字符，解析 field 控件 {len(fields)} 个", flush=True)
        time.sleep(2)
        ctx.close()
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
