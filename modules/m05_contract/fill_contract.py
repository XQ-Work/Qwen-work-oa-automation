# -*- coding: utf-8 -*-
"""m05 合同会签一条龙（070302 / workflowid=476402）

读合同发起台账 + 合同分支映射表 -> 登录 -> 严格级联顺序填表 -> 放大镜选择
-> 摘要 -> 印章区 -> 附件 -> --save 存草稿回写。提交永远人工。

用法：python modules/m05_contract/fill_contract.py --row 3 [--save]
"""
import argparse
import datetime
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from oa_common import paths, auth, db  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

LEDGER = paths.DATA / "合同发起台账.xlsx"
BRANCH = paths.DATA / "合同分支映射表.xlsx"
WF = "476402"
ADD = ("/workflow/request/AddRequest.jsp?workflowid=" + WF +
       "&isagent=0&beagenter=0&f_weaver_belongto_userid=")
KEEP_OPEN_SECONDS = 30
BROWSER = False  # 默认不自动进放大镜弹窗(易卡),留人工;--browser 才尝试

L = {  # 台账列
    2: "场景", 3: "合同名称", 4: "金额", 5: "甲方", 6: "乙方", 7: "关联关键词",
    8: "拟签日期", 9: "摘要覆盖", 10: "合同路径", 11: "状态"}
# 映射表列
M = {2: "事项类型", 3: "板块事项审批", 4: "项目状态W", 5: "项目种类",
     6: "项目状态", 7: "项目类型", 8: "合同内容", 9: "合同所属类型",
     10: "合同类型", 11: "采购类", 12: "人员所属"}
# OA 字段编号
F = {"所属板块": "field41343", "水务板块": "field41366", "项目状态W": "field41448",
     "项目种类": "field41447", "板块事项审批": "field41367", "投资类型": "field41306",
     "采购类": "field503130", "人员所属": "field175130", "关联审批": "field497630",
     "项目状态": "field41328", "项目类型": "field41314", "合同内容": "field48686",
     "合同所属类型": "field44486", "合同类型": "field41310", "新合同名称": "field41336",
     "母合同": "field41329", "合同总金额": "field41345", "甲方": "field41324",
     "乙方": "field41325", "债权类": "field53306", "报装": "field57106",
     "摘要": "field41340", "印章类型": "field41307", "印章分部": "field106131",
     "合同关联分部甲": "field378630",
     "监印人": "field316130", "文档名称": "field106130", "法人章": "field106133"}

# 甲方=本公司，供应商浏览框里的记录ID(实测捕获)；用于 JS 直填、不走易卡的弹窗
COMPANY_SUPPLIER_ID = {"荆州浦华荆清水务有限公司": "18404"}

# 本公司其它固定浏览器字段 → 记录ID(实测捕获)，同样 JS 直填
FIXED_FIELD_IDS = {
    "field41307": "1528",    # 印章类型
    "field106131": "325",    # 印章所属分部
    "field378630": "325",    # 合同关联分部甲
}


def read_row(r):
    wb = load_workbook(LEDGER, data_only=True)
    ws = wb["合同台账"]
    rec = {k: ("" if ws.cell(row=r, column=c).value is None
               else str(ws.cell(row=r, column=c).value).strip())
           for c, k in L.items()}
    wb.close()
    if not rec["合同名称"]:
        raise SystemExit(f"合同台账第{r}行合同名称为空")
    return rec


def read_branch(scenario):
    wb = load_workbook(BRANCH, data_only=True)
    ws = wb["分支映射"]
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value or "").strip() == scenario:
            return {k: ("" if ws.cell(row=r, column=c).value is None
                        else str(ws.cell(row=r, column=c).value).strip())
                    for c, k in M.items()}
    wb.close()
    raise SystemExit(f"映射表中无场景『{scenario}』")


def find_form_frame(page):
    for _ in range(20):
        best = None
        for f in page.frames:
            try:
                html = f.content()
            except Exception:
                continue
            n = html.count('name="field') if html else 0
            if n >= 20 and (best is None or n > best[0]):
                best = (n, f)
        if best:
            return best[1]
        time.sleep(2)
    return None


def sel(frame, fid, label):
    if not label:
        return
    frame.locator(f"select[name={fid}]:visible").first.select_option(label=label)


def txt(frame, fid, val):
    if val == "" or val is None:
        return
    frame.locator(f"input[name={fid}]:visible").first.fill(str(val))


def area(frame, fid, val):
    if val:
        frame.locator(f"textarea[name={fid}]").first.fill(val)


def _browser_expr(html, fid):
    """从帧 HTML 提取该字段的 onShowBrowser2 调用表达式"""
    m = re.search(r'onclick="(onShowBrowser2\([^"]*?%s[^"]*?)"' % fid.replace("field", ""),
                  html)
    if not m:
        return None
    return m.group(1).replace("\\'", "'")


def _tin(f):
    """可见的文本类输入框(排除隐藏/按钮/单选/复选)"""
    return f.locator("input:visible:not([type=hidden]):not([type=button])"
                     ":not([type=submit]):not([type=image]):not([type=radio])"
                     ":not([type=checkbox])")


def _desc(f):
    out = [f]
    try:
        for c in f.child_frames:
            out += _desc(c)
    except Exception:
        pass
    return out


def _dialog_frames(page, before):
    """evaluate 后新出现的、URL 命中浏览器对话框特征的帧(含其子帧),去重"""
    BR = ("browsermain", "commonbrowser", "formmode/bro", "formmode/sea",
          "checkbrowser", "/browser")
    roots = []
    for f in page.frames:
        if id(f) in before:
            continue
        u = (f.url or "").lower()
        if any(b in u for b in BR):
            roots.append(f)
    seen, res = set(), []
    for r in roots:
        for f in _desc(r):
            if id(f) not in seen:
                seen.add(id(f))
                res.append(f)
    return res


def _browser_search(frames, keyword, tag, owner):
    """在浏览器对话框帧子树里：填 searchName(或首个可见输入) -> 搜索 -> 点匹配行"""
    inp, inframe = None, None
    for f in frames:
        try:
            sn = f.locator("input[name=searchName]")
            if sn.count() > 0:
                inp, inframe = sn.first, f
                break
        except Exception:
            continue
    if inp is None:
        for f in frames:
            try:
                t = _tin(f)
                if t.count() > 0:
                    inp, inframe = t.first, f
                    break
            except Exception:
                continue
    if inp is None:
        print(f"[browser {tag}] 对话框内无搜索框", flush=True)
        return False
    try:
        inp.fill(keyword)
    except Exception:
        try:
            inp.type(keyword, delay=25)
        except Exception:
            pass
    clicked = False
    for f in [inframe] + [x for x in frames if x is not inframe]:
        for lab in ("搜索", "查询"):
            try:
                f.get_by_text(lab, exact=True).first.click(timeout=2000)
                clicked = True
                break
            except Exception:
                continue
        if clicked:
            break
    if not clicked:
        try:
            inp.press("Enter")
        except Exception:
            pass
    time.sleep(3)
    try:
        owner.screenshot(path=str(paths.STATE / f"browser_{tag}_result.png"), full_page=False)
    except Exception:
        pass
    mk = keyword[:4] if len(keyword) > 4 else keyword  # 网格中公司名常截断,用前缀匹配行
    # 选中：点含关键词前缀的结果行(点行即回填并关闭)
    for f in frames:
        try:
            cell = f.locator("td,div,a").filter(has_text=mk).first
            if cell.count() > 0:
                cell.click(timeout=3000)
                time.sleep(1.5)
                return True
        except Exception:
            continue
    # 兜底：选择链接
    for f in frames:
        try:
            pick = f.locator("a:has-text('选择'),input[value='选择']")
            if pick.count() > 0:
                pick.first.click(timeout=3000)
                time.sleep(1.5)
                return True
        except Exception:
            continue
    print(f"[browser {tag}] 搜索后无匹配行(前缀{mk})", flush=True)
    return False


def _jianming(name):
    """简称：取前 4 个连续字符(城市+字号),匹配更稳;过短易命中过多"""
    s = str(name or "").strip()
    return s[:4] if len(s) > 4 else s


def set_field_by_id(frame, fid, id_val, name_val):
    """JS 直填浏览器字段：隐藏值框=fid 存ID、显示框=fid__ 存名称、名称span 同步。
    表单帧 evaluate 传参会撞 e-cology 上下文坑(refs.set)，故把值内联进纯字符串。"""
    js = ("(function(){var fid=%s,idv=%s,nm=%s;"
          "var h=document.querySelector('input[name=\"'+fid+'\"]');"
          "var d=document.querySelector('input[name=\"'+fid+'__\"]');"
          "if(h){h.value=idv;}"
          "if(d){d.value=nm;}"
          "var sp=document.getElementById(fid+'span');"
          "if(sp){sp.innerHTML='<span class=\"e8_showNameClass\">'+nm+'</span>';}"
          "})()") % (json.dumps(fid), json.dumps(str(id_val)), json.dumps(str(name_val)))
    try:
        frame.evaluate(js)
    except Exception as e:
        print(f"[jsid {fid}] evaluate 异常 {str(e)[:100]}", flush=True)


def browser_select(page, frame, fid, keyword, tag):
    """放大镜选择：触发 onShowBrowser2 开弹窗 -> 全页找【可见】searchName 框(隐藏预建帧
    的不可见,天然唯一) -> 填简称 -> 点搜索 -> 点结果 <a>(点行即回填关闭)。"""
    try:
        expr = _browser_expr(frame.content(), fid)
        if not expr:
            print(f"[browser {tag}] 未找到 onShowBrowser2 表达式", flush=True)
            return False
        try:
            frame.evaluate("() => { %s; }" % expr)
        except Exception as e:
            print(f"[browser {tag}] evaluate 异常 {str(e)[:120]}", flush=True)
        kw = _jianming(keyword)
        sframe = None
        deadline = time.time() + 12
        while time.time() < deadline:
            time.sleep(0.7)
            for f in page.frames:
                try:
                    if f.locator("input[name=searchName]:visible").count() > 0:
                        sframe = f
                        break
                except Exception:
                    continue
            if sframe:
                break
        if sframe is None:
            print(f"[browser {tag}] 未见可见 searchName n_frames={len(page.frames)}",
                  flush=True)
            try:
                page.screenshot(path=str(paths.STATE / f"browser_{tag}_fail.png"),
                                full_page=False)
            except Exception:
                pass
            return False
        time.sleep(1)
        try:
            nvis = sum(1 for f in page.frames
                       if f.locator("input[name=searchName]:visible").count() > 0)
            print(f"[diag {tag}] sframe.url={sframe.url[:30]} 可见searchName帧数={nvis}",
                  flush=True)
        except Exception:
            pass
        sn = sframe.locator("input[name=searchName]:visible").first
        try:
            sn.fill(kw)
        except Exception:
            try:
                sn.type(kw, delay=30)
            except Exception:
                pass

        def _try_select():
            pool = _desc(sframe) + [f for f in page.frames
                                    if f is not frame and f is not sframe]
            for f in pool:
                try:
                    a = f.locator("a").filter(has_text=keyword)
                    if a.count() == 0:
                        a = f.locator("a").filter(has_text=kw)
                    cnt = a.count()
                except Exception:
                    continue
                for k in range(min(cnt, 5)):
                    try:
                        bb = a.nth(k).bounding_box()
                        if not bb:
                            continue
                        page.mouse.click(bb["x"] + 6, bb["y"] + bb["height"] / 2)
                        time.sleep(1.5)
                        try:
                            v = frame.locator(f"input[name={fid}]").first.input_value()
                        except Exception:
                            v = ""
                        if v:
                            return v
                    except Exception:
                        continue
            return ""

        def _attempt(f, how):
            try:
                if how == "img":
                    f.locator("span.searchImg, img[src*='search-input'],"
                              " .e8_btn_top_first").first.click(timeout=1500)
                elif how == "js":
                    f.evaluate("() => { try{onBtnSearchClick()}catch(e){} }")
                elif how == "btn":
                    f.get_by_text("搜索").first.click(timeout=1500)
                elif how == "enter":
                    sn.press("Enter")
                return True
            except Exception:
                return False

        seq = [("img", sframe), ("js", sframe), ("btn", sframe), ("enter", sframe)]
        for f in page.frames:
            if f is sframe:
                continue
            u = (f.url or "").lower()
            if any(k in u for k in ("formmode", "browsermain", "commonbrowser",
                                    "systeminfo/b")):
                seq += [("img", f), ("js", f), ("btn", f)]
        for how, f in seq:
            if not _attempt(f, how):
                continue
            time.sleep(2.2)
            val = _try_select()
            if val:
                print(f"[browser {tag}] 选中成功 via {how}@{f.url[:14]} val={val}",
                      flush=True)
                try:
                    page.screenshot(path=str(paths.STATE / f"browser_{tag}_result.png"),
                                    full_page=False)
                except Exception:
                    pass
                return True
        try:
            page.screenshot(path=str(paths.STATE / f"browser_{tag}_result.png"),
                            full_page=False)
        except Exception:
            pass
        print(f"[browser {tag}] 未能选中(简称{kw})", flush=True)
        return False
    except Exception as e:
        print(f"[browser-fail {tag}] {str(e)[:150]}", flush=True)
        return False


def _contract_name(rec):
    """合同名称=事项名称(关联关键词)+后缀；场景含'服务'→服务合同，否则→采购合同"""
    subj = (rec.get("关联关键词") or rec.get("合同名称") or "").strip()
    if not subj:
        return ""
    suffix = "服务合同" if "服务" in (rec.get("场景") or "") else "采购合同"
    return subj + suffix


def build_steps(rec, br):
    """严格级联顺序：所属板块→水务板块→项目状态W/种类→板块事项审批→投资类型
    →采购类/人员所属→关联审批(放大镜)→项目状态→项目类型→合同三下拉→文本
    →母合同/债权/报装→甲乙方(放大镜)→摘要→印章区"""
    steps = []
    steps.append(("sel", F["所属板块"], "水务板块"))
    steps.append(("sel", F["水务板块"], "水务项目公司合同审批"))
    steps.append(("wait", 3, None))
    steps.append(("sel", F["项目状态W"], br["项目状态W"]))
    steps.append(("sel", F["项目种类"], br["项目种类"]))
    steps.append(("sel", F["板块事项审批"], br["板块事项审批"]))
    steps.append(("wait", 3, None))
    steps.append(("auto_invest", F["投资类型"], None))  # 单选项自动选
    steps.append(("sel", F["采购类"], br["采购类"]))
    steps.append(("sel", F["人员所属"], br["人员所属"]))
    steps.append(("wait", 2, None))
    steps.append(("sel", F["项目状态"], br["项目状态"]))
    steps.append(("wait", 2, None))
    steps.append(("sel", F["项目类型"], br["项目类型"]))
    steps.append(("sel", F["合同内容"], br["合同内容"]))
    steps.append(("sel", F["合同所属类型"], br["合同所属类型"]))
    steps.append(("sel", F["合同类型"], br["合同类型"]))
    steps.append(("txt", F["新合同名称"], _contract_name(rec)))
    steps.append(("txt", F["合同总金额"], f"{float(rec['金额']):.2f}" if rec["金额"] else ""))
    steps.append(("sel", F["母合同"], "否"))
    _jid = COMPANY_SUPPLIER_ID.get(rec["甲方"])
    if _jid:
        steps.append(("jsid", F["甲方"], (_jid, rec["甲方"])))
    else:
        steps.append(("browser", F["甲方"], rec["甲方"]))
    steps.append(("browser", F["乙方"], rec["乙方"]))
    steps.append(("sel", F["债权类"], "否"))
    steps.append(("sel", F["报装"], "否"))
    steps.append(("area", F["摘要"], rec["摘要覆盖"] or _contract_name(rec)))
    for _key, _nm in (("印章类型", rec["甲方"] + "合同专用章"),
                      ("印章分部", rec["甲方"]),
                      ("合同关联分部甲", rec["甲方"])):
        _fid = F[_key]
        if _jid and _fid in FIXED_FIELD_IDS:
            steps.append(("jsid", _fid, (FIXED_FIELD_IDS[_fid], _nm)))
        else:
            steps.append(("browser", _fid, _nm))
    steps.append(("txt", F["文档名称"], _contract_name(rec)))
    steps.append(("sel", F["法人章"], "否"))
    return steps


def apply_step(page, frame, kind, fid, val, step):
    try:
        if kind == "wait":
            time.sleep(fid)
        elif kind == "auto_invest":
            html = frame.content()
            m = re.search(r'<select[^>]*name="%s"[^>]*>(.*?)</select>' % fid, html, re.S)
            if m:
                opts = [t.strip() for _, t in re.findall(
                    r'<option[^>]*value="([^"]*)"[^>]*>([^<]*)</option>', m.group(1))
                    if t.strip()]
                real = [o for o in opts if o != "投资类型"]
                if len(real) == 1:
                    sel(frame, fid, real[0])
                    step(f"{fid}={real[0]}", True, "单选项自动")
        elif kind == "sel":
            if not val:
                step(f"{fid}", True, "空值跳过")
                return
            sel(frame, fid, val)
            step(f"{fid}", True, str(val))
        elif kind == "jsid":
            idv, namev = val
            set_field_by_id(frame, fid, idv, namev)
            try:
                cur = frame.locator(f"input[name={fid}]").first.input_value()
            except Exception:
                cur = ""
            step(f"直填{fid}", cur == str(idv), f"id={cur!r} 期望{idv}")
        elif kind == "txt":
            txt(frame, fid, val)
            step(f"{fid}", True, str(val)[:30])
        elif kind == "area":
            area(frame, fid, val)
            step(f"{fid}", True, str(val)[:30])
        elif kind == "browser":
            if not BROWSER:
                step(f"放大镜{fid}", True, "跳过-请人工填")
                return
            ok = browser_select(page, frame, fid, val, fid)
            step(f"放大镜{fid}", ok, val)
    except Exception as e:
        step(f"{kind} {fid}", False, str(e)[:150])


def write_back(r, status, reqid=""):
    wb = load_workbook(LEDGER)
    ws = wb["合同台账"]
    ws.cell(row=r, column=11, value=status)
    if reqid:
        ws.cell(row=r, column=12, value=reqid)
    ws.cell(row=r, column=13,
            value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
    wb.save(LEDGER)


def main():
    global BROWSER
    ap = argparse.ArgumentParser()
    ap.add_argument("--row", type=int, required=True)
    ap.add_argument("--save", action="store_true")
    ap.add_argument("--browser", action="store_true",
                    help="尝试自动填放大镜字段(默认跳过,留人工)")
    args = ap.parse_args()
    BROWSER = args.browser
    rec = read_row(args.row)
    br = read_branch(rec["场景"])
    report = {"row": args.row, "rec": rec, "branch": br, "steps": []}

    def step(name, ok, err=""):
        report["steps"].append({"step": name, "ok": bool(ok), "err": str(err)[:150]})
        print(f"[{'ok ' if ok else 'ERR'}] {name} {err}", flush=True)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        for pg in ctx.pages[1:]:
            try:
                pg.close()
            except Exception:
                pass
        try:
            base = auth.login(page)
            step("登录", True, base)
            page.goto(base + ADD, wait_until="domcontentloaded", timeout=40000)
            frame = find_form_frame(page)
            if frame is None:
                step("定位表单帧", False)
                raise RuntimeError("无表单帧")
            step("定位表单帧", True, frame.url[:60])
            for kind, fid, val in build_steps(rec, br):
                apply_step(page, frame, kind, fid, val, step)
            time.sleep(2)
            page.screenshot(path=str(paths.STATE / "contract_filled.png"),
                            full_page=True)
            step("整页截图", True)
            (paths.STATE / "contract_report.json").write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            if args.save:
                signals = []
                page.on("dialog", lambda d: (signals.append(d.message or ""), d.accept()))
                page.locator("input.e8_btn_top[value=保存]").first.click(timeout=8000)
                reqid, ok = "", False
                for i in range(60):
                    time.sleep(2)
                    try:
                        body = page.inner_text("body")
                    except Exception:
                        body = ""
                    if "正在保存" in body or "保存中" in body:
                        continue
                    m = re.search(r"requestid=(\d+)", page.url or "", re.I)
                    if m:
                        reqid, ok = m.group(1), True
                        break
                    if any("成功" in s for s in signals):
                        ok = True
                        break
                    if i > 3 and "AddRequest" not in (page.url or ""):
                        m3 = re.search(r"requestid=(\d+)", page.url or "", re.I)
                        reqid = m3.group(1) if m3 else ""
                        ok = True
                        break
                page.screenshot(path=str(paths.STATE / "contract_saved.png"),
                                full_page=True)
                step("保存草稿", ok, f"requestid={reqid}")
                if ok:
                    write_back(args.row, "已存草稿", reqid)
            else:
                print(f"[hold] 干跑：浏览器保持 {KEEP_OPEN_SECONDS//60} 分钟", flush=True)
                time.sleep(KEEP_OPEN_SECONDS)
        finally:
            (paths.STATE / "contract_report.json").write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            try:
                ctx.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()
