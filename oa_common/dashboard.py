# -*- coding: utf-8 -*-
"""进度总览 dashboard（定稿版式 v15）：真实数据聚合生成 记录/进度总览.xlsx。

页1 进度总览：一行一个项目；流程三列 + 附件十列矩阵（列头=应交清单，
       格=已上传/未上传/—，红绿底色）+ 供应商/时间/下一步/灯。
页2 速查：固定手册模板（命名→清单矩阵→范例→灯→口径→口诀）。

判定口径（定稿）：
  流程状态  未完成=无草稿；已完成=草稿已存/已提交未办结；已归档=办结回写
  已上传    项目夹有实物 且 对应流程草稿已存（双确认）；未上传=缺其一
  附件列     仅在该流程已发起后激活（未到=—），水瀑式点亮
  供应商    合同乙方 → 发票销售方(付款台账收款单位) → 回单收款人 逐级
  灯        红=超付/草稿失败/名实存疑；黄=有未上传或队列待发；绿=无事
数据源：记录/项目档案 + 三张发起台账 + moneyline主档 + intake队列 + docindex。
CLI：python -m oa_common.dashboard      （intake run/flush 尾部自动调用）
"""
import datetime as dt
import json
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.cell.cell import MergedCell

from oa_common import paths, moneyline

OUT = paths.BOARD
F_N = lambda sz=9, b=False, col=None: Font(name="微软雅黑", size=sz, bold=b,
                                           color=col or "000000")
CTR = Alignment(wrap_text=True, vertical="center", horizontal="center")
THIN = Border(*[Side(style="thin", color="DDDDDD")]*4)
BLUE = PatternFill("solid", fgColor="4472C4")
G1 = PatternFill("solid", fgColor="8EA9DB")
G2 = PatternFill("solid", fgColor="A9D08E")
UP_F = PatternFill("solid", fgColor="E7F4E4")
DN_F = PatternFill("solid", fgColor="FDE9E9")
LAMP = {"绿": "C6EFCE", "黄": "FFEB9C", "红": "FFC7CE"}
S1, S2, S3 = "未完成", "已完成", "已归档"
UP, DN, NA = "已上传", "未上传", "—"

HEAD = ["编号", "事项", "类别", "合同金额", "供应商（全称）", "创建时间", "最近办理日期",
        "流程｜事项审批", "流程｜合同会签", "流程｜付款流程",
        "附件｜申购单", "附件｜事项审批", "附件｜报告单", "附件｜报价单", "附件｜采购合同",
        "附件｜流程会签", "附件｜用印合同", "附件｜发票", "附件｜付款单", "附件｜回单",
        "下一步该干嘛", "灯"]
# 附件列：(匹配材料词, 排除词, 激活条件键)
ATTACH = [
    ("申购单", None, "m01"), ("事项审批", None, "m05"),
    ("比质比价报告单|报告单", None, "m05"), ("供应商报价单及资质|报价单", None, "m05"),
    ("采购合同", "已用印|盖章", "m05"),
    ("流程会签", None, "m06"), ("采购合同", "已用印|盖章|invert", "m06"),
    ("发票|dzfp", None, "m06"), ("付款单", None, "m06"),
    ("回单", None, "m07"),
]


def norm(s):
    return re.sub(r"[（）()\s\-_·、,，【】]", "", str(s or ""))


def _rows(path, sheet):
    if not Path(path).exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    out = [[(c if c is not None else "") for c in row]
           for row in ws.iter_rows(min_row=2, values_only=True)]
    wb.close()
    return [r for r in out if any(str(x).strip() for x in r)]


def load_data():
    """项目档案目录 + 三台账 + 队列 + 主档，按项目编号归组。"""
    projs = {}
    for root, archived in ((paths.ARCHIVE, False),
                           (paths.ARCHIVED_SMALL, True), (paths.ARCHIVED_LARGE, True)):
        if not root.is_dir():
            continue
        for d in sorted(root.iterdir()):
            if d.is_dir() and re.match(r"^\d{4}-\d{3}-", d.name):
                code, matter = moneyline.code_from_path(d.name)
                projs[code] = dict(code=code, matter=matter, dir=d, files={},
                                   m01=[], m05=[], m06=[], archived=archived)
                for sub in d.rglob("*"):
                    if sub.is_file():
                        projs[code]["files"].setdefault(sub.parent.name[:2], []).append(sub)
    def link(rows, key, proj_col, name_cols):
        for r in rows:
            code = str(r[proj_col - 1] if len(r) >= proj_col else "").strip()
            if code in projs:
                projs[code][key].append(r)
                continue
            # 无编号列的行按名称回退匹配
            for c in name_cols:
                nm = norm(r[c - 1] if len(r) >= c else "")
                if len(nm) >= 3:
                    for p in projs.values():
                        if nm in norm(p["matter"]) or norm(p["matter"]) in nm:
                            p[key].append(r)
                            break
                    else:
                        continue
                    break
    link(_rows(paths.LEDGER_OA, "发起台账"), "m01", 17, [2, 5])
    link(_rows(paths.LEDGER_CONTRACT, "合同台账"), "m05", 15, [3, 7])
    link(_rows(paths.LEDGER_PAY, "付款台账"), "m06", 15, [2])
    q = {}
    if paths.STATE.joinpath("intake_queue.json").exists():
        for item in json.loads(paths.STATE.joinpath("intake_queue.json").read_text(encoding="utf-8")):
            if item.get("status") in ("queued", "failed", "parsing"):
                q.setdefault(item.get("matter", ""), []).append(item)
    return projs, q


def has_file(files, words, neg=None, matter=""):
    for f in files:
        n = norm(f.name)
        if any(norm(w) in n for w in words) and \
           not (neg and any(norm(x) in n for x in neg)):
            return True
    return False


DONE_WORDS = ("已存草稿", "已保存草稿", "已提交")
def flow_state(rows, col_status, done_words=DONE_WORDS, arch_words=("已付款", "已办结")):
    if not rows:
        return S1
    st = [str(r[col_status]) for r in rows]
    if any(any(w in s for w in arch_words) for s in st) and \
       all(any(w in s for w in arch_words + ("已办结",)) for s in st):
        return S3
    if any(any(w in s for w in done_words) for s in st):
        return S2
    return S1


def build():
    projs, queue = load_data()
    wb = Workbook()
    ws = wb.active
    ws.title = "进度总览"
    for c, h in enumerate(HEAD, 1):
        cell = ws.cell(1, c, h)
        cell.font = F_N(9, True, "FFFFFF")
        cell.fill = G1 if h.startswith("流程") else (G2 if h.startswith("附件") else BLUE)
        cell.alignment = CTR
    ws.row_dimensions[1].height = 26
    widths = [9, 12, 5, 8, 16, 9, 9, 7, 7, 7, 8, 8, 8, 8, 8, 8, 8, 8, 8, 7, 16, 5]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    ws.freeze_panes = "B2"

    r = 2
    order = sorted(projs.values(), key=lambda p: (p.get("archived"), p["code"]))
    for p in order:
        files = p["files"]
        st01 = flow_state(p["m01"], 12)
        st05 = flow_state(p["m05"], 10)
        st06 = flow_state(p["m06"], 10)
        lane_b = bool(p["m01"]) and not p["m05"] and _amount(p) < 2000
        active = {"m01": bool(p["m01"]),
                  "m05": (not lane_b) and p["m05"] and st05 != S1,
                  "m06": p["m06"] and st06 != S1,
                  "m07": any("已付款" in str(x[10]) for x in p["m06"])}
        dirs = {"m01": ("01", "04"), "m05": ("01", "02", "03", "04"),
                "m06": ("05",), "m07": ("05",)}
        cells = []
        pay_started = bool(p["m06"])
        for words, neg, key in ATTACH:
            words = words.split("|")
            negs = [x for x in (neg or "").split("|") if x and x != "invert"] or None
            if not active.get(key):
                cells.append(NA)
                continue
            pool = [f for d in dirs[key] for f in files.get(d, [])]
            if key == "m05" and words == ["采购合同"]:
                ok = has_file(pool, ["采购合同"], ["已用印", "盖章"])
            elif key == "m06" and words == ["采购合同"]:      # 用印合同列
                ok = has_file(pool, ["已用印", "盖章"])
            elif key == "m07":
                ok = any(f.suffix.lower() in (".png", ".jpg") for f in pool) or active["m07"]
            else:
                ok = has_file(pool, words, negs)
            if key == "m06" and not pay_started:
                cells.append(NA)     # 付款流程未建台账行：附件列先不逼问
            else:
                cells.append(UP if ok else DN)   # 已上传=档案有实物；流程进度列管草稿
        supplier = _supplier(p)
        amount = _amount(p)
        created = _cdate(p)
        last = _lastdate(p)
        flags = moneyline.flags(p["code"])
        miss = [HEAD[10 + i].split("｜")[1] for i, v in enumerate(cells) if v == DN]
        qk = queue.get(p["matter"], [])
        if flags or any(i["status"] == "failed" for i in qk):
            lamp, nxt = "红", "；".join(flags + ["有草稿发送失败，人工核对后重试 flush"])
        elif miss or qk or st01 == S1 or (st05 == S1 and not lane_b and p["m01"]):
            lamp = "黄"
            parts = []
            if qk:
                parts.append("跑 flush 发草稿")
            if miss:
                parts.append("补投：" + "、".join(miss) + "（事项名）投 输入/")
            if st01 == S1 and p["m01"]:
                parts.append("事项审批台账待确认/待发")
            nxt = "；".join(parts)
        else:
            lamp, nxt = "绿", "无（正常/已闭环）"
        vals = [p["code"], p["matter"], _cat(p), f"{amount:.2f}" if amount else NA,
                supplier, created, last, st01,
                (NA if lane_b else st05), st06] + cells + [nxt, lamp]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = F_N()
            cell.alignment = CTR
            cell.border = THIN
            if v in (S2, S3):
                cell.font = F_N(col="2E7D32", b=(v == S3))
            elif v == S1:
                cell.font = F_N(col="B45309")
            elif v == UP:
                cell.font = F_N(col="2E7D32")
                cell.fill = UP_F
            elif v == DN:
                cell.font = F_N(col="C62828", b=True)
                cell.fill = DN_F
            if c == len(HEAD):
                cell.fill = PatternFill("solid", fgColor=LAMP[v])
        if p.get("archived"):
            ws.cell(r - 0 if False else r, 1).value = str(ws.cell(r, 1).value) + "（已归档）"
            for c in range(1, len(HEAD) + 1):
                cell = ws.cell(r, c)
                cell.font = F_N(col="999999")
                cell.fill = PatternFill(fill_type=None)
            ws.cell(r, len(HEAD) - 0).value = "—"   # 归档行不再催办
        r += 1
    if r == 2:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(HEAD))
        cell = ws.cell(2, 1, "（暂无项目——把第一批材料投进 输入/ 并跑 intake_all.py run）")
        cell.font = F_N(col="808080")
        cell.alignment = CTR
    _guide(wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    return OUT, r - 2


def _amount(p):
    for x in p["m05"]:
        try:
            return float(str(x[3]).replace(",", ""))
        except (ValueError, IndexError):
            pass
    for x in reversed(p["m01"]):
        try:
            return float(str(x[3]).replace(",", ""))
        except (ValueError, IndexError):
            pass
    return 0


def _supplier(p):
    for x in p["m05"]:
        if str(x[5]).strip():
            return str(x[5]).strip()
    for x in p["m06"]:
        if str(x[3]).strip():
            return str(x[3]).strip()
    return ""


def _cat(p):
    for x in p["m01"]:
        t = str(x[2])
        return "货物" if "货物" in t else ("服务" if "服务" in t else "")
    return ""


def _cdate(p):
    try:
        return dt.datetime.fromtimestamp(p["dir"].stat().st_ctime).strftime("%Y-%m-%d")
    except OSError:
        return ""


def _lastdate(p):
    ds = []
    for d in p["files"].values():
        for f in d:
            ds.append(dt.datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d"))
    for rows, col in ((p["m01"], 14), (p["m05"], 12), (p["m06"], 12)):
        for x in rows:
            v = re.search(r"\d{4}-\d{2}-\d{2}", str(x[col]))
            if v:
                ds.append(v.group())
    return max(ds)[:10] if ds else ""


# ── 页2：速查手册（v15 定稿模板）──────────────────────────────
def _guide(wb):
    g = wb.create_sheet("速查")
    N = 8
    BD = Border(*[Side(style="thin", color="BFBFBF")]*4)
    LB = PatternFill("solid", fgColor="D6E4F0")
    GRe, YE, RE_, OR = (PatternFill("solid", fgColor=c) for c in
                       ("C6EFCE", "FFEB9C", "FFC7CE", "FFF2CC"))
    def band(r, text):
        g.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N)
        for c in range(1, N + 1):
            g.cell(r, c).border = BD
            g.cell(r, c).fill = BLUE
        a = g.cell(r, 1, text)
        a.font = F_N(10, True, "FFFFFF")
        a.alignment = CTR
        g.row_dimensions[r].height = 24

    def _row(r, cells, header=False, fill=None, h=22):
        for i in range(1, N + 1):
            cell = g.cell(r, i)
            v = cells[i - 1] if i - 1 < len(cells) else None
            if not isinstance(cell, MergedCell) and v not in (None, ""):
                cell.value = v
            cell.alignment = CTR
            cell.border = BD
            cell.font = F_N(10, True, "1F3864") if header else F_N(10)
            if header:
                cell.fill = LB
            elif fill and i == 1:
                cell.fill = fill
        g.row_dimensions[r].height = h

    def kv(r, k, v):
        _row(r, [k, v])
        g.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
    g.merge_cells("A1:H1")
    t = g.cell(1, 1, "速查手册")
    t.font = F_N(14, True, "1F3864")
    t.alignment = CTR
    g.row_dimensions[1].height = 32
    band(3, "一、命名总格式")
    kv(4, "", "材料名（事项名）.pdf　——　事项名＝项目夹“2026-NNN-”后的字；"
              "可带金额：采购合同（潜污泵采购-金额：15877）.pdf")
    g.cell(4, 1).value = None
    band(6, "二、材料清单（投料即按此名，系统按列头核对）")
    _row(7, ["材料名", "投料后归入", "事项审批必传", "合同会签必传", "付款申请必传",
             "免命名分流", "备注"], header=True)
    g.merge_cells("G7:H7")
    items = [
        ["申购单", "01_事项审批", "必传", "—", "—", "", "从 输入/ 投"],
        ["事项审批", "01_事项审批", "—", "必传", "—", "", "批复扫描/打印件"],
        ["报告单", "03_比质比价报告单", "—", "必传", "—", "", "全名：比质比价报告单"],
        ["报价单", "02_供应商报价及资质", "—", "必传", "—", "", "全名：供应商报价单及资质"],
        ["采购合同", "04_采购合同", "—", "必传", "—", "", "未盖章原件"],
        ["流程会签", "05_付款及验收", "—", "—", "必传", "", ""],
        ["用印合同", "05_付款及验收", "—", "—", "必传", "", "采购合同（事项名 已用印）"],
        ["发票", "05_付款及验收", "—", "—", "必传", "dzfp_ 自动识别", "电子发票pdf"],
        ["付款单", "05_付款及验收", "—", "—", "必传", "系统回流自动改名", ""],
        ["回单", "银行回单流水线", "—", "—", "—", "含 回单/交易流水/银行流水", "匹配后复制进05夹"],
        ["其他", "06_其他", "—", "—", "—", "", "非常规附件（检测报告、赠品单据等）"],
        ["证照类", "供应商档案", "—", "—", "—", "", "营业执照等（供应商全称）"],
    ]
    r = 8
    for it in items:
        _row(r, it)
        g.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        r += 1
    r += 1
    band(r, "三、范例：以「潜污泵采购」为例")
    r += 1
    for k, v, fl in [
        ("事项审批所需附件", "申购单（潜污泵采购）.pdf", OR),
        ("合同会签所需附件", "事项审批（潜污泵采购）.pdf、报告单（潜污泵采购）.docx、"
         "报价单（潜污泵采购）.pdf、采购合同（潜污泵采购）.pdf ※未盖章原件", None),
        ("付款流程所需附件", "流程会签（潜污泵采购）.pdf、采购合同（潜污泵采购 已用印）.pdf、"
         "发票（潜污泵采购）.pdf、付款单（潜污泵采购）.pdf", None),
        ("结清回单", "不用改名：回单_××_15877.00_2026-10-08.png 直接丢 输入/，系统自动识别归档", GRe)]:
        _row(r, [k, v], fill=fl, h=30)
        g.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
        r += 1
    r += 1
    band(r, "四、灯与状态")
    r += 1
    for k, v, fl in [("绿灯", "无需你动手：正常推进或已闭环", GRe),
                     ("黄灯", "在等你：缺材料待投 / 等银行回单——看“未上传”红格与下一步列", YE),
                     ("红灯", "出了错须人工：超付、草稿失败、名实存疑（文件名与正文不符）", RE_),
                     ("流程状态", "未完成＝还没发起　｜　已完成＝草稿已存/已提交未办结　｜　已归档＝办结回写", None)]:
        _row(r, [k, v], fill=fl)
        g.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
        r += 1
    r += 1
    band(r, "五、字段口径")
    r += 1
    for k, v in [("已上传", "项目夹已收到该实物文件（流程走到哪步看左侧三列，缺什么看这里）"),
                 ("未上传", "把“材料名（事项名）”文件投进 输入/，系统自动归档并挂传"),
                 ("供应商（全称）", "合同乙方 → 发票销售方 → 回单收款人 逐级取"),
                 ("创建时间 / 最近办理", "项目建档（首次收料）日 / 该项目最近一次动作时间")]:
        _row(r, [k, v])
        g.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
        r += 1
    r += 1
    band(r, "六、投料三句口诀")
    r += 1
    for k, v in [("一句", "要发起流程的料，一律改成「材料名（事项名）」再丢进 输入/"),
                 ("二句", "回单、入库单、电子发票不用改名，系统自己认"),
                 ("三句", "丢完跑 intake_all.py run；发草稿攒到 flush 一次性办")]:
        _row(r, [k, v])
        g.merge_cells(start_row=r, start_column=2, end_row=r, end_column=N)
        r += 1
    for col, w in zip("ABCDEFGH", [20, 22, 12, 12, 12, 20, 22, 6]):
        g.column_dimensions[col].width = w
    g.sheet_view.showGridLines = False


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    out, n = build()
    print(f"进度总览已生成：{out}（项目 {n} 个）")
