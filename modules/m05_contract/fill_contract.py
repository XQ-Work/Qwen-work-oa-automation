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
import shutil
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
BROWSER = True  # 默认自动填放大镜(行内联想已稳);--no-browser 才跳过留人工

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


def find_annex_frame(page):
    """附件控件在独立嵌套帧(含 #field-annexupload)，非表单主帧。"""
    for _ in range(10):
        for f in page.frames:
            try:
                if f.query_selector("#field-annexupload"):
                    return f
            except Exception:
                pass
        time.sleep(2)
    return None


# 找出“文档上传”明细表对应的加行按钮 id（$addbuttonN$）
_JS_ADDBTN = (
    "(function(){var bs=document.querySelectorAll(\"button[id^='$addbutton']\");"
    "for(var i=0;i<bs.length;i++){var t=bs[i].closest('table');"
    "if(t&&/文档上传/.test(t.textContent||''))return bs[i].id;}return '';})()")


def _mark_idx_js(idx):
    """清旧标记，给第 idx 行(field54276_{idx} 所在 tr)的 file/text input 打标记。"""
    return (
        "(function(){var o=document.querySelectorAll('input[data-upl],input[data-dname]');"
        "for(var x=0;x<o.length;x++){o[x].removeAttribute('data-upl');o[x].removeAttribute('data-dname');}"
        "var h=document.querySelector(\"input[name='field54276_%d']\");if(!h)return 0;"
        "var tr=h.closest('tr');if(!tr)return 0;"
        "var fi=tr.querySelectorAll('input[type=file]');for(var j=0;j<fi.length;j++)fi[j].setAttribute('data-upl','1');"
        "var tx=tr.querySelectorAll('input[type=text]');for(var k=0;k<tx.length;k++)tx[k].setAttribute('data-dname','1');"
        "return fi.length;})()") % idx


def _docid_js(idx):
    return ("(function(){var h=document.querySelector(\"input[name='field54276_%d']\");"
            "return h?(h.value||''):'';})()") % idx


def _state_js(idx):
    """返回 'docid|busy'，busy=1 表示该行仍有 Uploading/Pending。"""
    return ("(function(){var h=document.querySelector(\"input[name='field54276_%d']\");"
            "var tr=h?h.closest('tr'):null;var t=tr?tr.textContent.replace(/\\s+/g,' '):'';"
            "var busy=(t.indexOf('Uploading')>=0||t.indexOf('Pending')>=0)?1:0;"
            "return (h?(h.value||''):'')+'|'+busy;})()") % idx


# 推 SWFUpload 队列(多文件时第一文件完成后需再 startUpload 才带上后续文件)
_JS_STARTQ = (
    "(function(){if(window.SWFUpload&&SWFUpload.instances){for(var k in SWFUpload.instances){"
    "try{var u=SWFUpload.instances[k];if(u.getStats().files_queued>0)u.startUpload();}catch(e){}}}})()")


# 附件取件根目录（测试期=文档整理；最终位置待定，见 backlog）
ATTACH_ROOT = paths.DATA / "OA附件库" / "文档整理"
# 4 个必传件（材料名），文件名规则 = 材料名（事项）.pdf
REQUIRED_DOCS = ["事项审批", "比质比价报告单", "供应商报价单及资质", "采购合同"]


def resolve_attachments(kw, root=None, contract_path=None):
    """按 材料名（kw）.pdf 在 root 递归找 4 必传件；缺任一则抛 FileNotFoundError。
    采购合同优先用台账 J 列路径。排除 已用印/盖章/扫描 版。"""
    root = Path(root or ATTACH_ROOT)
    got, missing = {}, []
    for base in REQUIRED_DOCS:
        f = None
        if base == "采购合同" and contract_path and Path(contract_path).exists():
            f = Path(contract_path)
        else:
            want = f"{base}（{kw}）.pdf"
            for cand in sorted(root.rglob(want)):
                if any(s in cand.name for s in ("已用印", "盖章", "扫描")):
                    continue
                f = cand
                break
        if f is None:
            missing.append(f"{base}（{kw}）.pdf")
        else:
            got[base] = f
    if missing:
        raise FileNotFoundError("缺少文件：" + "、".join(missing))
    return got


def _set_docname(frame, idx, name):
    """填文档名称：优先 Playwright fill(派发真实事件,OA才认)，JS 写值兜底。"""
    try:
        loc = frame.locator(f"input[name='field54275_{idx}']")
        if loc.count():
            loc.first.fill(name)
            return 1
    except Exception:
        pass
    js = ("(function(){var e=document.querySelector(\"input[name='field54275_%d']\");"
          "if(!e)return 0;e.value=%s;"
          "e.dispatchEvent(new Event('input',{bubbles:true}));"
          "e.dispatchEvent(new Event('change',{bubbles:true}));return 1;})()") % (idx, json.dumps(name))
    try:
        return frame.evaluate(js)
    except Exception:
        return 0


def _attach_rows(rec):
    kw = rec.get("关联关键词") or rec.get("合同名称") or ""
    kind = "服务合同" if "服务" in (rec.get("场景") or "") else "采购合同"
    return [
        (["事项审批"], "事项审批"),
        (["比质比价报告单", "供应商报价单及资质"], "比质比价报告单、供应商报价单及资质"),
        (["采购合同"], rec.get("合同名称") or f"{kw}{kind}"),
    ]


def upload_begin(page, got, rec, step):
    """起传：加 3 行 + 塞文件 + 推队列，快速返回(不在此等上传完成)。
    返回 state 供 upload_finish 收口；失败返回 None。"""
    frame = find_annex_frame(page)
    if frame is None:
        step("附件-定位帧", False, "无附件帧")
        return None
    btn_id = frame.evaluate(_JS_ADDBTN)
    if not btn_id:
        step("附件-无文档上传表", False)
        return None
    rows = _attach_rows(rec)
    added = []
    for idx, (files, dname) in enumerate(rows):
        try:
            frame.locator("button[id='%s']" % btn_id).first.click(force=True, timeout=4000)
        except Exception as e:
            step(f"附件-行{idx+1}加行", False, str(e)[:60]); continue
        time.sleep(1.2)
        frame.evaluate(_mark_idx_js(idx))
        fi = frame.locator("input[type=file][data-upl='1']")
        if fi.count() == 0:
            step(f"附件-行{idx+1}", False, "无file input"); continue
        plist = [str(got[b]) for b in files]
        try:
            fi.last.set_input_files(plist, timeout=8000)
            frame.evaluate(_JS_STARTQ)
            added.append(idx)
        except Exception as e:
            step(f"附件-行{idx+1}塞文件", False, str(e)[:70])
    step("附件-起传", bool(added), f"已塞{len(added)}行,后台上传中")
    return {"frame": frame, "rows": rows, "added": added, "got": got}


def upload_finish(state, rec, step):
    """收口：等所有已加行 docid 落定且无 Uploading/Pending，再按 name 填文档名称，据实报告。"""
    if not state:
        return False
    frame, rows, added, got = state["frame"], state["rows"], state["added"], state["got"]
    docids = {}
    for _ in range(90):
        frame.evaluate(_JS_STARTQ)          # 持续推队列,确保多文件/并发都带上
        allok = bool(added)
        for idx in added:
            try:
                raw = frame.evaluate(_state_js(idx)) or "|"
            except Exception:
                raw = "|"
            d, b = (raw.split("|") + ["0"])[:2]
            docids[idx] = d
            if not d or d == "NULL" or b != "0":
                allok = False
        if allok:
            break
        time.sleep(1)
    for idx, (files, dname) in enumerate(rows):
        if idx in added:
            _set_docname(frame, idx, dname)
    allok = True
    for idx, (files, dname) in enumerate(rows):
        d = docids.get(idx, "")
        ok = bool(d and d != "NULL")
        step(f"附件-行{idx+1}[{dname}]", ok,
             " + ".join(got[b].name for b in files) + f" docid={d}")
        allok = allok and ok
    return allok


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
    """行内联想选供应商：点字段框→打全称→退格删1字触发→等下拉→点含前缀的项→回读校验。"""
    try:
        box = frame.locator(f"#innerContent{fid}div")
        if box.count() == 0:
            box = frame.locator(f"td[id={fid}_tdwrap]")
        if box.count() == 0:
            print(f"[browser {tag}] 找不到联想框", flush=True)
            return False
        box.first.click(force=True, timeout=6000)
        page.keyboard.type(keyword, delay=90)
        page.keyboard.press("Backspace")   # 删1字触发 keyup 联想查询
        time.sleep(1.8)
        pre = keyword[:4] if len(keyword) > 4 else keyword
        picked = False
        for sel in ("li", "a", "td", "div"):
            try:
                loc = frame.locator(f"{sel}:visible").filter(has_text=pre)
                if loc.count() > 0:
                    loc.first.click(force=True, timeout=3000)
                    picked = True
                    break
            except Exception:
                continue
        time.sleep(1.5)
        try:
            v = frame.locator(f"input[name={fid}]").first.input_value()
        except Exception:
            v = ""
        if v:
            print(f"[browser {tag}] 联想选中 val={v}", flush=True)
            return True
        print(f"[browser {tag}] 联想未命中(前缀{pre} picked={picked})", flush=True)
        try:
            page.screenshot(path=str(paths.STATE / f"browser_{tag}_fail.png"),
                            full_page=False)
        except Exception:
            pass
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
    ap.add_argument("--no-browser", action="store_true",
                    help="跳过放大镜字段(留人工);默认自动填")
    ap.add_argument("--no-attach", action="store_true",
                    help="跳过附件上传;默认按台账J列合同路径上传")
    args = ap.parse_args()
    BROWSER = not args.no_browser
    attach = not args.no_attach
    rec = read_row(args.row)
    br = read_branch(rec["场景"])
    attach_files = None
    if attach:
        try:
            attach_files = resolve_attachments(
                rec.get("关联关键词") or "", contract_path=rec.get("合同路径"))
            print("[附件] 4件齐：" + " | ".join(
                f"{k}={v.name}" for k, v in attach_files.items()), flush=True)
        except FileNotFoundError as e:
            raise SystemExit(f"[附件] {e} —— 已中止(不填表/不上传/不保存)")
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
            # 先起传附件(后台上传),再填表单——上传耗时与填表时间重叠
            up_state = None
            if attach and attach_files:
                try:
                    up_state = upload_begin(page, attach_files, rec, step)
                except Exception as e:
                    step("附件-起传", False, str(e)[:100])
            for kind, fid, val in build_steps(rec, br):
                apply_step(page, frame, kind, fid, val, step)
            time.sleep(2)
            page.screenshot(path=str(paths.STATE / "contract_filled.png"),
                            full_page=True)
            step("整页截图", True)
            if attach and up_state:
                try:
                    upload_finish(up_state, rec, step)
                except Exception as e:
                    step("附件-收口", False, str(e)[:100])
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
                    time.sleep(2)
                    print("[postsave-url] " + (page.url or ""), flush=True)
                    try:
                        vf = find_annex_frame(page) or frame
                        tbl = vf.evaluate(
                            "(function(){var tds=document.querySelectorAll('td'),t;"
                            "for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传')"
                            "{t=tds[i].closest('table');break;}}"
                            "if(!t)return 'no table';"
                            "var ids=[];t.querySelectorAll('input[type=hidden]').forEach(function(e){"
                            "if(/^field\\d+_\\d+$/.test(e.name)&&/^\\d/.test(e.value||''))ids.push(e.name+'='+e.value);});"
                            "var nm=[];t.querySelectorAll('input[type=text]').forEach(function(e){"
                            "if(/^field\\d+_\\d+$/.test(e.name))nm.push(e.name+'=\"'+(e.value||'')+'\"');});"
                            "return 'names:'+nm.join(',')+' || docids:'+ids.join(',');})()")
                        step("存后回读-附件", "docids:" in tbl and "=" in tbl.split("docids:")[-1], tbl)
                    except Exception as e:
                        step("存后回读-附件", False, str(e)[:80])
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
