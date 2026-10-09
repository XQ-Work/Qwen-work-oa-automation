# -*- coding: utf-8 -*-
"""收件导流器（intake dispatcher）——立即归档、攒批发草稿。

用法：
  python intake_all.py scan            只预报：每个收件文件会被路由到哪，不动文件
  python intake_all.py run             收料：识别→归档进项目夹→需要发草稿的入队（不开浏览器）
  python intake_all.py flush [--only m05,m06]   批量发草稿：按队列依次跑 fill_*（开浏览器）
  python intake_all.py status          队列/收件箱/异常 三态一览

投递约定：{data_root}/输入 是唯一投料口（平铺，不建子文件夹）。
  首选命名 材料名（事项名）.pdf（可选带金额）——按词表精确路由，见 ROUTES；
  未按规范命名的按文件名特征兜底分流（dzfp_/入库单/申购单_→单据类；回单/流水→M07），见 TYPE_ROUTES。
  例：合同（滤布采购）.pdf   申购单（PE配件-金额：999.91）.pdf   dzfp_2642….pdf
两者都认不出的 → 移入 工作/复核/收件异常 等人工。run 永不打开浏览器；
发草稿只在 flush 时按队列串行执行，失败项留在队列可重试。
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
from oa_common import paths, archive  # noqa: E402

EXCEPT_DIR = paths.EXCEPT
QUEUE_F = paths.STATE / "intake_queue.json"
DONE_F = paths.STATE / "intake_done.json"
RECV_DIR = paths.RECEIPT_DIR
DANJU_IN = paths.DANJU_BASE / "10_输入"
LEDGER_CONTRACT = paths.LEDGER_CONTRACT
LEDGER_PAY = paths.LEDGER_PAY

# 材料名(别名) -> (规范名, 归档阶段子夹, 后续动作)
# 动作: archive=纯归档 | m02/m05/m06=解析+排队发草稿 | m07=转回单夹 | danju=转单据类
ROUTES = [
    (["申购单", "请购单"],              "申购单",       "01_事项审批",     "m02"),
    (["事项审批"],                      "事项审批",     "01_事项审批",     "archive"),
    (["比质比价报告", "比质比价报告单", "比价报告"], "比质比价报告", "03_比质比价报告单", "archive"),
    (["报价单及资质", "报价单", "供应商资质"],       "报价单及资质", "02_供应商报价及资质", "archive"),
    (["合同", "采购合同", "服务合同"],   "合同",         "04_采购合同",     "m05"),
    (["发票", "电子发票"],               "发票",         "05_付款及验收",   "m06"),
    (["付款单"],                         "付款单",       "05_付款及验收",   "archive"),
    (["其他"],                         "其他",         "06_其他",       "archive"),
    (["验收资料", "验收单"],             "验收资料",     "05_付款及验收",   "archive"),
    (["回单", "转账回单", "银行回单"],   "回单",         None,              "m07"),
    (["入库单", "入库单扫描件"],         "入库单",       None,              "danju"),
]
_ALIAS = {a: (canon, stage, act) for aliases, canon, stage, act in ROUTES for a in aliases}

# 未按「材料名（事项名）」命名时的类型兜底分流（平铺投料口，按名字特征直达流水线进料口）
TYPE_ROUTES = [
    (re.compile(r"^dzfp_|数电票|电子发票|发票_\d|^增值税.*发票", re.I), "danju", "单据类·发票记账"),
    (re.compile(r"^入库单", re.I), "danju", "单据类·入库(自动分拣)"),
    (re.compile(r"^(申购单|事项审批)_", re.I), "danju", "单据类·配套单据"),
    (re.compile(r"回单|交易流水|银行流水", re.I), "recv", "M07·回单扫描"),
]

NAME_RE = re.compile(r"^(?P<mat>.+?)[（(](?P<inner>[^（）()]+)[)）]\s*$")
AMT_RE = re.compile(r"[-_\s]*金额[:：]\s*([\d.]+)\s*$")


def utf8_console():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    if __import__("os").name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass


def now():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def content_route(f):
    """PDF 首页文字层关键词兜底：文件名不认识时看内容。"""
    if f.suffix.lower() != ".pdf":
        return None
    try:
        import pymupdf as fitz
        with fitz.open(str(f.resolve())) as d:
            head = d[0].get_text()[:600]
    except Exception:
        return None
    if "入库单" in head or "采购入库" in head:
        return ("danju", "单据类·入库(内容识别)")
    if "发票号码" in head or "电子发票" in head:
        return ("danju", "单据类·发票记账(内容识别)")
    if "回单" in head or "交易流水" in head:
        return ("recv", "M07·回单扫描(内容识别)")
    return None


def classify(stem):
    """文件名 -> (规范材料名, 事项名, 金额或None, 动作, 阶段夹)；认不出返回 None+原因。"""
    m = NAME_RE.match(stem.strip())
    if not m:
        return None, "缺「材料名（事项名）」括号命名"
    mat = re.sub(r"\s+", "", m.group("mat"))
    hit = _ALIAS.get(mat)
    if not hit:
        return None, f"材料名「{mat}」不在词表"
    canon, stage, act = hit
    inner = m.group("inner").strip()
    amt = None
    am = AMT_RE.search(inner)
    if am:
        amt = float(am.group(1))
        inner = inner[:am.start()]
    matter = re.sub(r"\s+", "", inner)
    if len(matter) < 2:
        return None, "括号内事项名过短"
    # 采购合同带用印标记 = 付款侧必传的盖章扫描版：剥掉标记归 05 夹，不走合同抽取链
    if canon == "合同" and re.search(r"已用印|已盖章|用印版", matter):
        matter = re.sub(r"(已用印|已盖章|用印版)", "", matter)
        if not matter:
            return None, "用印合同缺事项名"
        return ("用印合同", matter, amt, "archive", "05_付款及验收"), None
    return (canon, matter, amt, act, stage), None


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_json(p, default):
    return json.loads(Path(p).read_text(encoding="utf-8")) if Path(p).exists() else default


def save_json(p, obj):
    Path(p).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def place(src, dst_dir, dst_name):
    """移动+改名（幂等：同哈希目标已存在则只清收件箱原件）。返回最终路径。"""
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / dst_name
    if dst.exists():
        if sha256(src) == sha256(dst):
            src.unlink()
            return dst
        stem, suf = dst.stem, dst.suffix
        i = 2
        while (dst_dir / f"{stem}_{i}{suf}").exists():
            i += 1
        dst = dst_dir / f"{stem}_{i}{suf}"
    shutil.move(str(src), str(dst))
    return dst


def find_ledger_row(ledger, sheet, col, needle):
    from openpyxl import load_workbook
    wb = load_workbook(ledger, read_only=True)
    ws = wb[sheet]
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        v = row[col - 1] if len(row) >= col else None
        if v and needle in str(v):
            wb.close()
            return r
    wb.close()
    return None


def run_sub(cmd, cwd=None):
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    out = (p.stdout or "") + (p.stderr or "")
    return p.returncode, out


# ---------------- 各动作的解析链（run 阶段，不开浏览器） ----------------

def chain_m05(item, dst):
    """合同：抽取+回写合同发起台账 -> 队列带 row。"""
    code, out = run_sub([sys.executable, str(BASE / "modules/m05_contract/extract_contract.py"),
                         str(dst), "--commit"])
    item["extract_log"] = out.strip().splitlines()[-1] if out.strip() else ""
    row = find_ledger_row(LEDGER_CONTRACT, "合同台账", 10, str(dst.resolve()))
    if row:
        item["row"] = row
        item["status"] = "queued"
    else:
        item["status"] = "parse_failed"   # 扫描件/要素缺失：文件已归档，等人工/视觉兜底
    return item


def chain_m06(item, dst):
    """发票：解析+追加付款发起台账 -> 队列带 row。"""
    code, out = run_sub([sys.executable, str(BASE / "modules/m06_payment/intake_payment.py"),
                         str(dst), "--contract", item["matter"] + "合同", "--commit"])
    item["extract_log"] = out.strip().splitlines()[-1] if out.strip() else ""
    row = find_ledger_row(LEDGER_PAY, "付款台账", 10, str(dst.resolve()))
    if row:
        item["row"] = row
        item["status"] = "queued"
    else:
        item["status"] = "parse_failed"
    return item


def chain_m02(item, dst):
    item["status"] = "needs_agent"   # 申购单要素需视觉识别：intake.py scan -> add -> fill_oa
    item["extract_log"] = "待智能体：python modules/m02_shxiang/intake.py scan"
    try:                             # 文件名带金额则预估金额入主档（金额链起点）
        from oa_common import moneyline
        code, matter = moneyline.code_from_path(dst)
        if code and item.get("amount"):
            moneyline.set_amount(code, matter, "est", item["amount"], source=dst.name)
            item["extract_log"] += f"；est={item['amount']} 已入主档 {code}"
    except Exception as e:
        item["extract_log"] += f"；[moneyline] {e}"
    return item


# ---------------- 顶层命令 ----------------

def _index(path, material="", proj=None):
    """投料即建档：正文入全文索引+错投检测（失败不拦收料）；selftest 沙箱不触真索引。"""
    if _SELFTEST:
        return
    try:
        from oa_common import docindex
        code = matter = ""
        if proj is not None:
            m = re.match(r"^(\d{4}-\d{3})-(.+)$", proj.name)
            code, matter = (m.group(1), m.group(2)) if m else ("", "")
        r = docindex.ingest(path, material=material, code=code, matter=matter)
        if r["status"] == "待视觉":
            print("    [索引] 扫描件已登记元数据，待第二阶段OCR")
        elif r["warn"]:
            print(f"    ⚠ [索引] {r['warn']}")
    except Exception as e:
        print(f"    [索引] 失败(不影响收料): {str(e)[:80]}")


_SELFTEST = False


def cmd_scan(inbox, live=False, ledger_mode="real"):
    global _SELFTEST
    _SELFTEST = (ledger_mode == "copy")
    queue = load_json(QUEUE_F, [])
    done = load_json(DONE_F, {})
    moved, bad = 0, 0
    if not inbox.is_dir():
        print("收件箱不存在:", inbox)
        return queue
    for f in sorted(inbox.iterdir()):
        if not f.is_file() or f.name.startswith("~"):
            continue
        plan, err = classify(f.stem)
        if not plan:
            # 类型兜底：未按命名规范投递的，按文件名特征分流进各流水线进料口
            tr = next(((k, w) for rx, k, w in TYPE_ROUTES if rx.search(f.name)), None)
            if not tr:
                tr = content_route(f)
            if tr:
                kind, why = tr
                tgt = DANJU_IN if kind == "danju" else (RECV_DIR / "输入")
                h = sha256(f)
                key = f"{f.name}|{h[:16]}"
                if key in done:
                    print(f"[重复投递] {f.name}  （原件已在 {done[key]}，移入 复核/内容幂等 人工处理）")
                    if live:
                        dup = paths.EXCEPT.parent / "内容幂等"
                        place(f, dup, f.name)
                    continue
                print(f"[类型分流] {f.name}  ->  {why}")
                if live:
                    moved += 1
                    dst = place(f, tgt, f.name)
                    done[key] = str(dst)
                    tp = "danju" if kind == "danju" else "m07"
                    queue.append(dict(id=h[:12], ts=now(), type=tp, matter="",
                                      file=str(dst), amount=None, row=None,
                                      status="queued", attempts=0, last_err=""))
                    save_json(DONE_F, done)
                    save_json(QUEUE_F, queue)
                    _index(dst, material=kind)
                continue
            bad += 1
            print(f"[异常] {f.name}  ->  {err}")
            if live:
                EXCEPT_DIR.mkdir(parents=True, exist_ok=True)
                shutil.move(str(f), str(EXCEPT_DIR / f.name))
            continue
        canon, matter, amt, act, stage = plan
        proj = archive.project_dir(matter, create=True)
        if act == "m07":
            tgt = RECV_DIR / "输入"
        elif act == "danju":
            tgt = DANJU_IN
        else:
            tgt = proj / stage
        h = sha256(f)
        key = f"{f.name}|{h[:16]}"   # 同名同内容才算重复投递（同内容不同材料名要允许）
        if key in done:
            print(f"[重复投递] {f.name}  （原件已在 {done[key]}，移入 复核/内容幂等 人工处理）")
            if live:
                dup = paths.EXCEPT.parent / "内容幂等"
                place(f, dup, f.name)
            continue
        show = (str(tgt.relative_to(paths.DATA)) if act in ("m07", "danju")
                else f"{proj.name}/{stage}")
        print(f"[{act}] {f.name}  ->  {show}"
              f"{'  金额=' + str(amt) if amt else ''}")
        if not live:
            continue
        moved += 1
        dst = place(f, tgt, f.name)
        done[key] = str(dst)
        _index(dst, material=canon, proj=proj)
        if canon == "事项审批":          # 交付区同步出册：事项审批导出PDF
            try:
                from oa_common import deliver
                deliver.stage("事项审批", [dst], matter)
            except Exception as e:
                print("    [交付] 失败(不影响收料): " + str(e)[:60])
        if act == "archive":
            save_json(DONE_F, done)
            continue
        item = dict(id=h[:12], ts=now(), type=act, matter=matter, file=str(dst),
                    amount=amt, row=None, status="queued", attempts=0, last_err="")
        if act in ("m05", "m06", "m02"):
            if ledger_mode == "copy":       # 自测模式：不动正式台账
                item["extract_log"] = "(自测跳过解析)"
            else:
                item = {"m05": chain_m05, "m06": chain_m06, "m02": chain_m02}[act](item, dst)
        queue.append(item)
        save_json(DONE_F, done)
        save_json(QUEUE_F, queue)
    if live:
        save_json(DONE_F, done)
        save_json(QUEUE_F, queue)
        n_q = sum(1 for q in queue if q["status"] == "queued")
        print(f"完成：本次收料 {moved} 个，异常 {bad} 个；队列待发草稿 {n_q} 项（flush 执行）")
        if ledger_mode != "copy":
            _board()
        if not ledger_mode == "copy" and moved:
            try:    # 投料即建档收尾：本次新登记的扫描件顺手OCR（有缓存，秒级）
                from oa_common import docindex
                import io, contextlib
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    docindex.ocr_backfill()
                line = buf.getvalue().strip().splitlines()[-1] if buf.getvalue().strip() else ""
                if line and "成功 0，未成功 0" not in line:
                    print("[索引OCR] " + line)
            except Exception as e:
                print("[索引OCR] 失败(不影响收料): " + str(e)[:80])
    return queue


def _board():
    """收料/发草稿后自动重建 记录/进度总览.xlsx（失败不拦主流程）。"""
    try:
        from oa_common import dashboard
        out, n = dashboard.build()
        print(f"[总览] 已刷新 {out.name}（项目 {n} 个）")
    except Exception as e:
        print("[总览] 刷新失败(不影响本次操作): " + str(e)[:100])


def cmd_flush(only=None):
    queue = load_json(QUEUE_F, [])
    todo = [q for q in queue if q["status"] in ("queued", "failed")
            and (not only or q["type"] in only)]
    if not todo:
        print("队列里没有可发的草稿（scan 看预报 / run 收料）。")
        return
    by_type = {}
    for q in todo:
        by_type.setdefault(q["type"], []).append(q)
    ran_danju = False
    for t, items in by_type.items():
        print(f"\n### {t}: {len(items)} 项")
        if t in ("danju", "m07") and items:      # 组批：跑一次全家，组内统一记账
            one = items[0]
            one["attempts"] = one.get("attempts", 0) + 1
            cmd = ([sys.executable, str(BASE / "danju/adapter.py"), "run"] if t == "danju"
                   else [sys.executable, str(BASE / "modules/m07_receipt/match_archive.py")])
            code, out = run_sub(cmd)
            tail = out.strip().splitlines()[-1] if out.strip() else ""
            for q in items:
                q["attempts"] = q.get("attempts", 0) + 1
                q["status"] = "done" if code == 0 else "failed"
                q["last_err"] = tail[:120]
            save_json(QUEUE_F, queue)
            print(f"  [{'ok ' if code == 0 else 'ERR'}] {t} 组批 x{len(items)} 一次执行  {tail[:70]}")
            continue
        for q in items:
            if q["type"] in ("m05", "m06") and not q.get("row"):
                q["status"] = "needs_review"
                q["last_err"] = "无台账行（解析失败/扫描件/自测残留），人工核对台账后补 row 再 flush"
                print(f"  [跳过] {q['matter']}：{q['last_err']}")
                save_json(QUEUE_F, queue)
                continue
            q["attempts"] = q.get("attempts", 0) + 1
            try:
                if t == "m05":
                    code, out = run_sub([sys.executable, str(BASE / "modules/m05_contract/fill_contract.py"),
                                         "--row", str(q["row"]), "--save"])
                elif t == "m06":
                    code, out = run_sub([sys.executable, str(BASE / "modules/m06_payment/fill_payment.py"),
                                         "--row", str(q["row"]), "--save"])
                elif t == "m07":
                    code, out = run_sub([sys.executable, str(BASE / "modules/m07_receipt/match_archive.py")])
                elif t == "danju":
                    code, out = run_sub([sys.executable, str(BASE / "danju/adapter.py"), "run"])
                    ran_danju = True
                else:
                    continue
                tail = out.strip().splitlines()[-1] if out.strip() else ""
                m = re.search(r"requestid[=:]\s*(\d+)", out)
                if code == 0:
                    q["status"] = "done"
                    q["requestid"] = m.group(1) if m else ""
                    q["last_err"] = tail[:120]
                    print(f"  [ok] {q['matter']} -> 草稿 {q['requestid'] or tail[:60]}")
                else:
                    q["status"] = "failed"
                    q["last_err"] = (m.group(0) if m else tail)[:200]
                    print(f"  [ERR 第{q['attempts']}次] {q['matter']}: {q['last_err'][:80]}")
            except Exception as e:
                q["status"] = "failed"
                q["last_err"] = str(e)[:200]
                print(f"  [EXC] {q['matter']}: {e}")
            save_json(QUEUE_F, queue)
    left = [q for q in queue if q["status"] in ("queued", "failed")]
    print(f"\n剩余待重试 {len(left)} 项。")
    if ran_danju:
        print("\n— 单据类产出回流 —")
        cmd_reflow()


DANJU_OUT = paths.DANJU_BASE / "20_输出"
REFLOW_LOG = paths.STATE / "danju_reflow.json"


def cmd_verify(code=None):
    """项目本地一致性对账：主档金额链 vs 台账状态 vs 项目夹附件齐备（不开浏览器）。
    OA 端附件回读在 fill_* 保存后已打印，此处只核本地档案完整性。"""
    from oa_common import moneyline
    from openpyxl import load_workbook
    root = paths.ARCHIVE
    if not root.is_dir():
        print("档案根不存在:", root)
        return
    dirs = [d for d in sorted(root.iterdir()) if d.is_dir() and "示例" not in d.name]
    if code:
        dirs = [d for d in dirs if d.name.startswith(code)]
    def stage_files(d, stage, *pats):
        p = d / stage
        if not p.is_dir():
            return []
        import fnmatch
        return [f for f in p.iterdir()
                if f.is_file() and any(fnmatch.fnmatch(f.name, pt) for pt in pats)]
    def ledger_codes(ledger, sheet, col, rmax_col=None):
        out = {}
        p = ledger
        if not p.exists():
            return out
        wb = load_workbook(p, read_only=True)
        ws = wb[sheet]
        for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            c = row[col - 1] if len(row) >= col else None
            if c:
                out.setdefault(str(c).strip(), []).append(
                    (r, str(row[rmax_col - 1] or "") if rmax_col and len(row) >= rmax_col else ""))
        wb.close()
        return out
    hc = ledger_codes(paths.LEDGER_CONTRACT, "合同台账", 15, 11)
    pc = ledger_codes(paths.LEDGER_PAY, "付款台账", 15, 11)
    ac = ledger_codes(paths.LEDGER_OA, "发起台账", 17, 13)
    for d in dirs:
        cd, matter = moneyline.code_from_path(d.name)
        if not cd or (code and cd != code and d.name != code):
            continue
        miss = []
        if not stage_files(d, "01_事项审批", "*"):
            miss.append("缺01事项审批")
        if not stage_files(d, "02_供应商报价及资质", "*"):
            miss.append("缺02报价资质")
        if not stage_files(d, "03_比质比价报告单", "*"):
            miss.append("缺03比质比价")
        ht = stage_files(d, "04_采购合同", "*")
        if not ht:
            miss.append("缺04合同")
        elif not stage_files(d, "04_采购合同", "*已用印*"):
            miss.append("合同未用印")
        if not stage_files(d, "05_付款及验收", "付款单*"):
            miss.append("缺06付款单")
        if not stage_files(d, "05_付款及验收", "发票*", "*发票（*）*"):
            miss.append("缺06发票")
        h = hc.get(cd, [])
        ps = pc.get(cd, [])
        as_ = ac.get(cd, [])
        pinfo = moneyline.get(cd) or {}
        fl = "/".join(moneyline.flags(cd)) if pinfo else ""
        line = (f"{d.name}  事项审批[{','.join(s or '-' for _, s in as_) or '—'}] "
                f"合同[{','.join(s or '-' for _, s in h) or '—'}] "
                f"付款[{','.join(s or '-' for _, s in ps) or '—'}] "
                f"主档:合同{pinfo.get('contract_amount') or '—'}/已付{pinfo.get('paid_total') or 0}")
        print(line + ("  ⚠ " + "；".join(miss) if miss else "  ✓齐备")
              + (f"  [{fl}]" if fl else ""))


def _norm_sup(s):
    return re.sub(r"(有限公司|经营部|商行|商店|店|厂|（个体工商户）|科|技)", "", str(s or ""))


def _match_payee_rows(sup, amt):
    """在付款台账找 收款单位≈sup 且 |金额-amt|≤1 的行 -> [(row, matter)]"""
    from openpyxl import load_workbook
    p = paths.LEDGER_PAY
    if not p.exists():
        return []
    wb = load_workbook(p, read_only=True)
    ws = wb["付款台账"]
    hits = []
    ns = _norm_sup(sup)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 5 or not row[3]:
            continue
        name = str(row[3])
        try:
            ramt = float(row[4])
        except (TypeError, ValueError):
            continue
        nn = _norm_sup(name)
        if abs(ramt - float(amt)) <= 1.5 and (ns and (ns in nn or nn in ns)):
            cn = re.sub(r"(采购合同|服务合同|合同)$", "", str(row[1] or ""))
            hits.append((row[0] if row[0] else 0, cn or name))
    wb.close()
    return hits


def _danju_usage(sup, amt):
    """兜底：从单据类库里按 供应商+金额 找发票用途（Lane B 无付款台账行时当事项名）。"""
    try:
        import sqlite3
        try:
            dj = Path(paths.DANJU_BASE)
        except Exception:
            dj = None
        dbp = (dj or BASE / "danju") / "60_数据" / "单据.db"
        if not dbp.exists():
            return None
        conn = sqlite3.connect("file:" + str(dbp).replace(chr(92), "/") + "?mode=ro", uri=True)
        rows = conn.execute("SELECT usage, seller FROM invoices WHERE ROUND(amount)=ROUND(?)",
                            (float(amt),)).fetchall()
        conn.close()
        for usage, seller in rows:
            s1, s2 = _norm_sup(sup), _norm_sup(seller)
            if usage and s1 and (s1 in s2 or s2 in s1):
                return str(usage).strip()
    except Exception:
        pass
    return None


def cmd_reflow(dry=False):
    """单据类产出回流：把 danju/20_输出 新产生的 付款单/验收单 按台账匹配改名投收件箱。
    匹配次序：付款台账(收款+金额) → 单据类发票用途兜底 → 人工。
    首次运行只登记存量（历史件不回流），此后仅新增文件参与。"""
    first_run = not REFLOW_LOG.exists() and not dry
    log = load_json(REFLOW_LOG, {})
    manual = []
    n_seed = 0
    for sub, mat in [("付款单", "付款单"), ("验收单", "验收资料"), ("发票", "发票")]:
        d = DANJU_OUT / sub
        if not d.is_dir():
            continue
        for f in sorted(d.iterdir()):
            if not f.is_file() or f.suffix.lower() not in (".pdf", ".xlsx", ".docx"):
                continue
            key = f"{sub}|{f.name}|{f.stat().st_size}"
            if key in log:
                continue
            import datetime as _d
            today = _d.date.today()
            if first_run and _d.date.fromtimestamp(f.stat().st_mtime) < today:
                log[key] = "seed存量"
                n_seed += 1
                continue
            m = re.search(r"_([1-9]\d{1,8})$", f.stem)     # 末尾金额段
            parts = f.stem.split("_")
            sup = ""
            if m and len(parts) >= 2:
                # 去掉大类前缀词，取真正的供应商段
                parts = [x for x in parts if x not in ("发票", "付款单", "验收单", "入库单扫描件", "申购单", "事项审批")]
                amt = m.group(1)
                for j in range(len(parts) - 1, 0, -1):
                    if parts[j] == "_" + amt or parts[j] == amt:
                        sup = parts[j - 1]
                        break
            else:
                amt = None
            if not (amt and sup and not sup.isdigit()):
                manual.append((f, "文件名解析不出供应商/金额"))
                continue
            src = f                                      # PDF优先：同目录或 30_待打印 里的同名PDF
            for cand in (f.parent / (f.stem + ".pdf"), BASE.parent.parent / "工作/单据/30_待打印" / (f.stem + ".pdf")):
                pass
            from oa_common import paths as _pp
            for cand in (f.parent / (f.stem + ".pdf"), _pp.DANJU_BASE / "30_待打印" / (f.stem + ".pdf")):
                if cand.exists():
                    src = cand
                    break
            hits = _match_payee_rows(sup, amt)
            uniq = sorted({mt for _, mt in hits})
            if not uniq:
                usage = _danju_usage(sup, amt)
                if usage:
                    uniq = [usage]
            try:                           # 交付区：转PDF+按事项合并（在用途兜底之后，口径一致）
                from oa_common import deliver
                deliver.stage(sub, [src], uniq[0] if len(uniq) == 1 else sup)
            except Exception as e:
                print("  [交付] 失败(不影响回流):", str(e)[:60])
            dst = None
            if len(uniq) == 1:
                tag = "" if hits else "（按发票用途）"
                dst = paths.INBOX / f"{mat}（{uniq[0]}）{src.suffix}"
                print(f"[回流] {f.name} -> {dst.name}{tag}")
                if not dry:
                    paths.INBOX.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(str(src), str(dst))
            else:
                manual.append((f, f"事项不唯一{uniq or '无命中'}"))
            if not dry and dst:
                log[key] = str(dst)
                save_json(REFLOW_LOG, log)
    if first_run:
        if not dry:
            save_json(REFLOW_LOG, log)
        print(f"首次初始化：登记存量 {n_seed} 件不回流（只登记基线），此后仅新增文件参与。"
              " 需要把某个存量件回流时，删除 state/danju_reflow.json 中对应行即可。")
    if manual and not first_run:
        print("— 以下需人工按「材料名（事项名）」命名后投收件箱 —")
        for f, why in manual:
            print(f"  {f.name}  ({why})")


def _split_code(code_or_folder):
    """把 '2026-005' / '2026-005-事项名' / '2026-005-事项名〔办结…〕' 归一为 (code, folder_for_matter)。"""
    m = re.match(r"^(\d{4}-\d{3})", code_or_folder.strip())
    code = m.group(1) if m else code_or_folder.strip()
    return code


def cmd_archive(code, reason="", undo=False):
    """双轨归档 / 取出。归档=记录/项目档案原地不动，从交付区组装定稿册进归档架 + 记台账卡片 +
    刷台账快照；取出=把定稿册移入 运行/state/归档取出回收 并删台账卡片行。全系统零删除。
    code 支持 2026-NNN 前缀或完整夹名。"""
    from oa_common import archive as arc
    code = _split_code(code)

    def locate(root, want):
        if not Path(root).is_dir():
            return None
        for d in sorted(Path(root).iterdir()):
            if d.is_dir() and (d.name == want or d.name.startswith(want + "-")):
                return d
        return None

    src = locate(paths.ARCHIVE, code)                       # 任务轨源夹（过程全量）
    old = next((x for x in (locate(r, code)
                            for r in (paths.ARCHIVED_SMALL, paths.ARCHIVED_LARGE)) if x), None)

    if undo:
        moved = arc.remove_browse_folder(code)
        if not moved:
            print("归档架里没有:", code)
            return 1
        _archive_del_row(code)
        arc.refresh_snapshots()
        try:
            from oa_common import dashboard
            dashboard.build()
        except Exception:
            pass
        print(f"[取出] {code} 定稿册已移入 运行/state/归档取出回收（记录/项目档案未受影响），台账卡片行已删")
        return 0

    if not src and not old:
        print("项目档案里没有:", code)
        return 1
    folder = (src or old).name
    code2, matter = arc._code_matter(folder)
    amt = _shelf_amount(folder)
    big = amt is not None and amt >= 2000
    shelf = paths.ARCHIVED_LARGE if big else paths.ARCHIVED_SMALL
    if old:                                                   # 换架/重归：先回收旧册
        arc.remove_browse_folder(code)
    dest, files = arc.build_browse_folder(code, matter, amt, big)
    _archive_write_row(shelf, f"{code}-{matter}", reason)
    snap = arc.refresh_snapshots()
    try:
        from oa_common import dashboard
        dashboard.build()
    except Exception:
        pass
    print(f"[归档] {code} {matter} → {dest.parent.name}/{dest.name}（金额{amt if amt is not None else '未知,归日常'}）")
    print(f"  定稿册 {len(files)} 份（取自交付区）；台账快照刷新 {len(snap)} 项")
    return 0


def _shelf_amount(folder_name):
    """分架依据：主档合同金额 → 事项审批台账金额 → 申购需求台账，取到即用。"""
    try:
        from oa_common import moneyline
        code, matter = moneyline.code_from_path(folder_name)
        p = moneyline.get(code)
        if p and p.get("contract_amount"):
            return float(p["contract_amount"])
        from openpyxl import load_workbook
        for led, sheet, kcol, acol in (
                (paths.LEDGER_OA, "发起台账", 2, 4),
                (paths.LEDGER_SG, None, 1, None)):
            try:
                wb = load_workbook(led, read_only=True, data_only=True)
                ws = wb[sheet] if sheet else wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    kv = str(row[kcol - 1] or "")
                    if kv and (matter in kv or kv.startswith(matter)):
                        if acol:
                            v = row[acol - 1]
                            if v not in (None, ""):
                                wb.close()
                                return float(str(v).replace(",", ""))
                wb.close()
            except Exception:
                pass
    except Exception:
        pass
    return None


HDR_LARGE = ["编号", "创建时间", "事项名称", "事项类别", "供应商", "合同金额", "税率",
             "签订时间", "应付金额", "已付金额", "未付金额", "付款时间", "其他"]
HDR_SMALL = ["编号", "创建时间", "审批状态", "审批日期", "入库单单号", "付款", "供应商", "其他"]


def _find_rows(ledger, sheet, pred, cols):
    from openpyxl import load_workbook
    out = []
    if not Path(ledger).exists():
        return out
    wb = load_workbook(ledger, read_only=True, data_only=True)
    ws = wb[sheet] if sheet in wb.sheetnames else wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            if pred(row):
                out.append([row[c - 1] if len(row) >= c else "" for c in cols])
        except Exception:
            pass
    wb.close()
    return out


def _collect_info(folder_name):
    """按用户口径汇集归档行：创建时间=事项审批发起；类别仅 采购类/服务类。"""
    from oa_common import moneyline
    code, matter = moneyline.code_from_path(folder_name)
    info = dict(code=code, matter=matter, created="", category="采购类",
                supplier="", contract_amt="", tax="", pay_amt="", paid="", unpaid="",
                pay_times="", appr_status="", appr_date="", receipts="")
    nrm = lambda s: re.sub(r"[\s（）()]", "", str(s or ""))
    nm = nrm(matter)

    def hit(r, idx):
        v = r[idx] if len(r) > idx else None
        return "" if v is None else str(v).strip()

    r1 = _find_rows(paths.LEDGER_OA, "发起台账",
                    lambda r: (hit(r, 16) == code) or (nm and nm in nrm(hit(r, 1))),
                    [3, 4, 13, 15])
    if r1:
        info["appr_status"] = hit(r1[0], 2)
        info["appr_date"] = hit(r1[0], 3)[:10]
        info["created"] = info["appr_date"]
        info["category"] = "服务类" if "服务" in hit(r1[0], 0) else "采购类"
        if hit(r1[0], 1):
            info["pay_amt"] = hit(r1[0], 1)
    r5 = _find_rows(paths.LEDGER_CONTRACT, "合同台账",
                    lambda r: (hit(r, 14) == code) or (nm and nm in nrm(hit(r, 2))),
                    [4, 6, 13])
    if r5:
        info["contract_amt"] = hit(r5[0], 0)
        info["supplier"] = hit(r5[0], 1)
    r6 = _find_rows(paths.LEDGER_PAY, "付款台账",
                    lambda r: (hit(r, 14) == code) or (nm and nm in nrm(hit(r, 1))),
                    [4, 5, 11, 13])
    if r6:
        info["supplier"] = info["supplier"] or hit(r6[0], 0)
        if len(r6) == 1 and not info["contract_amt"]:
            info["contract_amt"] = hit(r6[0], 1)
        paid = [x for x in r6 if "已付款" in hit(x, 2)]
        if paid:
            try:
                info["paid"] = round(sum(float(hit(x, 1)) for x in paid), 2)
            except ValueError:
                pass
            info["pay_times"] = "、".join(sorted({hit(x, 3)[:10] for x in paid if hit(x, 3)}))
    try:
        p = moneyline.get(code)
        if p:
            info["contract_amt"] = info["contract_amt"] or str(p.get("contract_amount") or "")
            if info["paid"] == "":
                info["paid"] = str(p.get("paid_total") or "")
            if p.get("gap") is not None and info["contract_amt"]:
                info["unpaid"] = p["gap"]
            if p.get("category") and "服务" in str(p["category"]):
                info["category"] = "服务类"
            info["tax"] = p.get("tax") or ""
            if not info["pay_amt"]:
                info["pay_amt"] = str(p.get("appr_amount") or p.get("contract_amount")
                                      or p.get("est_amount") or "")
    except Exception:
        pass
    if info["contract_amt"] and info["unpaid"] == "" and info["paid"] != "":
        try:
            info["unpaid"] = round(float(info["contract_amt"]) - float(info["paid"] or 0), 2)
        except ValueError:
            pass
    try:                                   # 单据类库：单号关联 + 供应商按金额反查兜底
        import sqlite3
        dbp = Path(paths.DANJU_BASE) / "60_数据" / "单据.db"
        if dbp.exists():
            conn = sqlite3.connect("file:" + str(dbp).replace(chr(92), "/") + "?mode=ro", uri=True)
            if info["supplier"]:
                core = re.sub(r"(有限公司|经营部|商行|（个体工商户）)", "", info["supplier"])[:4]
                rs = conn.execute("SELECT receipt_id FROM receipts WHERE supplier LIKE ?",
                                  ("%" + core + "%",)).fetchall()
                info["receipts"] = "、".join(x[0] for x in rs)
            if not info["supplier"] and info["pay_amt"]:
                rs = conn.execute(
                    "SELECT DISTINCT supplier FROM receipts WHERE ROUND(total)=ROUND(?)",
                    (float(str(info["pay_amt"]).replace(",", "")),)).fetchall()
                if len(rs) == 1:
                    info["supplier"] = rs[0][0]
                    rs2 = conn.execute(
                        "SELECT receipt_id FROM receipts WHERE ROUND(total)=ROUND(?)",
                        (float(str(info["pay_amt"]).replace(",", "")),)).fetchall()
                    if not info["receipts"]:
                        info["receipts"] = "、".join(x[0] for x in rs2)
            conn.close()
    except Exception:
        pass
    return info


def _archive_write_row(shelf_dir, folder_name, reason):
    """写两张分身（归档/ + 记录/台账/），同编号换架自动迁行。"""
    from openpyxl import Workbook, load_workbook
    large = str(shelf_dir).startswith(str(paths.ARCHIVED_LARGE))
    hdr = HDR_LARGE if large else HDR_SMALL
    sheet = "采购合同" if large else "日常采购"
    info = _collect_info(folder_name)
    vals = dict(编号=info["code"], 创建时间=info["created"], 事项名称=info["matter"],
                事项类别=info["category"], 供应商=info["supplier"],
                合同金额=info["contract_amt"], 税率=info["tax"], 签订时间=now()[:10],
                应付金额=info["contract_amt"] or info["pay_amt"], 已付金额=info["paid"],
                未付金额=info["unpaid"], 付款时间=info["pay_times"],
                其他=reason or "办结归档", 审批状态=info["appr_status"],
                审批日期=info["appr_date"], 入库单单号=info["receipts"],
                付款=info["pay_amt"])
    for led in paths.ARCHIVE_LOGS:
        wb = load_workbook(led) if led.exists() else Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        ws = wb[sheet] if sheet in wb.sheetnames else wb.create_sheet(sheet)
        if ws.max_row == 0 or ws.cell(1, 1).value != "编号":
            ws.append(hdr)
        key = str(info["code"])
        tgt_r = None
        for rr in range(2, ws.max_row + 1):
            if str(ws.cell(rr, 1).value or "").strip() == key:
                tgt_r = rr
                break
        r = tgt_r or ws.max_row + 1
        for i, h in enumerate(hdr, 1):
            ws.cell(r, i, vals.get(h, ""))
        other = "日常采购" if large else "采购合同"
        if other in wb.sheetnames:
            w2 = wb[other]
            for rr in range(w2.max_row, 1, -1):
                if str(w2.cell(rr, 1).value or "").strip() == key:
                    w2.delete_rows(rr)
        led.parent.mkdir(parents=True, exist_ok=True)
        wb.save(led)
    print(f"[归档台账] {sheet}/{info['code']} {info['matter']} 已记行")


def _archive_del_row(folder_name):
    from openpyxl import load_workbook
    code = "-".join(folder_name.split("-")[:2])
    for led in paths.ARCHIVE_LOGS:
        if not led.exists():
            continue
        wb = load_workbook(led)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for rr in range(ws.max_row, 1, -1):
                if str(ws.cell(rr, 1).value or "").strip() == code:
                    ws.delete_rows(rr)
        wb.save(led)


def cmd_status():
    queue = load_json(QUEUE_F, [])
    print("== 队列 ==")
    from collections import Counter
    c = Counter(q["status"] for q in queue)
    for k, v in sorted(c.items()):
        print(f"  {k}: {v}")
    for q in queue:
        if q["status"] in ("queued", "failed", "parse_failed", "needs_agent", "needs_review"):
            print(f"  [{q['status']}] {q['type']} {q['matter']} row={q.get('row')} {q.get('last_err','')[:60]}")
    n_box = len([f for f in paths.INBOX.iterdir() if f.is_file()]) if paths.INBOX.is_dir() else 0
    n_ex = len([f for f in EXCEPT_DIR.iterdir() if f.is_file()]) if EXCEPT_DIR.is_dir() else 0
    print(f"收件箱待收料: {n_box} 个文件；收件异常待处理: {n_ex} 个")


def main():
    utf8_console()
    ap = argparse.ArgumentParser(description="收件导流器：立即归档、攒批发草稿")
    ap.add_argument("cmd", choices=["scan", "run", "flush", "status", "verify", "reflow", "deliver", "archive"])
    ap.add_argument("--inbox", default=str(paths.INBOX))
    ap.add_argument("--archive-root", default=None, help="覆盖项目档案根（自测用）")
    ap.add_argument("--only", default="", help="flush 只跑这些类型，逗号分隔，如 m05,m06")
    ap.add_argument("--project", default="", help="verify 只看某个项目编号")
    ap.add_argument("--dry-run", action="store_true", help="reflow 只预览不投递")
    ap.add_argument("--selftest", action="store_true", help="run 时不写正式台账")
    ap.add_argument("--reason", default="", help="归档原因")
    ap.add_argument("--undo", action="store_true", help="归档取出")
    a = ap.parse_args()
    if a.selftest:   # 自测：队列/异常/回单/danju入口全部指到临时区，绝不碰真实数据
        global EXCEPT_DIR, QUEUE_F, DONE_F, RECV_DIR, DANJU_IN
        troot = Path(a.archive_root or (paths.STATE / "intake_selftest"))
        EXCEPT_DIR = troot / "收件异常"
        QUEUE_F = troot / "intake_queue.json"
        DONE_F = troot / "intake_done.json"
        RECV_DIR = troot / "转账回单"
        DANJU_IN = troot / "danju_10_输入"
        a.archive_root = str(troot / "采购项目档案")
    if a.archive_root:
        paths.ARCHIVE = Path(a.archive_root)
    only = [x.strip() for x in a.only.split(",") if x.strip()] or None
    if a.cmd == "scan":
        cmd_scan(Path(a.inbox))
    elif a.cmd == "run":
        cmd_scan(Path(a.inbox), live=True, ledger_mode="copy" if a.selftest else "real")
    elif a.cmd == "flush":
        cmd_flush(only)
    elif a.cmd == "verify":
        cmd_verify(a.project or None)
    elif a.cmd == "reflow":
        cmd_reflow(dry=a.dry_run)
    elif a.cmd == "deliver":
        cmd_reflow(dry=a.dry_run)          # 复用扫描：新文件即交付+回流
    elif a.cmd == "archive":
        sys.exit(cmd_archive(a.project.strip().rstrip("/\\").split("\\")[-1], a.reason, a.undo))
    else:
        cmd_status()


if __name__ == "__main__":
    main()
