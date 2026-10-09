# -*- coding: utf-8 -*-
"""交付区 v2：只收 PDF，按事项合并，日期前缀，附交付台账。

规则（定稿）：
  · 交付/功能/日期前缀+功能（事项名）.pdf  —— 前缀日期=成员里最晚文件日期
  · 同一（功能+事项）多次出单 → 成员清单记录在 运行/state/delivery_manifest.json，
    每次 stage() 整份重建（转PDF+合并），不叠页；日期变了即自然形成新版文件（可区分）
  · 每份成品在 记录/台账/交付台账.xlsx 追加一行记录（时间/功能/事项/文件/成员/页数）
  · 原件链路不变，交付区纯副本，随便打印抽走
用法：deliver.stage("付款单", [Path(..)], "潜污泵采购")
"""
import datetime as dt
import hashlib
import json
import re
from pathlib import Path

from oa_common import paths

MANIFEST = paths.STATE / "delivery_manifest.json"
TMP = paths.STATE / "deliv_tmp"
# 台账双存（用户规矩）：主份在 记录/台账/，副本随功能夹；两处分身同步追记
LEDGERS = [paths.DATA / "记录/台账/06_交付台账.xlsx", paths.DELIVER / "交付台账.xlsx"]
LEDGER_HEAD = ["生成时间", "功能", "事项", "交付文件", "成员数", "页数", "成员清单", "备注"]


def _sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()[:16]


def _load():
    if MANIFEST.exists():
        return json.loads(MANIFEST.read_text(encoding="utf-8"))
    return {}


def _save(m):
    MANIFEST.parent.mkdir(parents=True, exist_ok=True)
    MANIFEST.write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")


# ---------- 格式转 PDF（本机 Office COM） ----------

def _to_pdf(src, outdir):
    src = Path(src)
    if src.suffix.lower() == ".pdf":
        return src if src.exists() else None
    outdir.mkdir(parents=True, exist_ok=True)
    dst = outdir / (src.stem + ".pdf")
    import win32com.client as w
    if src.suffix.lower() in (".docx", ".doc"):
        app = w.Dispatch("Word.Application")
        try:
            doc = app.Documents.Open(str(src.resolve()), ReadOnly=True)
            doc.ExportAsFixedFormat(str(dst.resolve()), 17)   # wdExportFormatPDF
            doc.Close(False)
        finally:
            app.Quit()
    elif src.suffix.lower() in (".xlsx", ".xls"):
        app = w.Dispatch("Excel.Application")
        try:
            wbk = app.Workbooks.Open(str(src.resolve()), ReadOnly=True)
            wbk.ExportAsFixedFormat(0, str(dst.resolve()))    # xlTypePDF
            wbk.Close(False)
        finally:
            app.Quit()
    else:
        return None
    return dst if dst.exists() else None


def _merge(pdf_list, dst):
    import pymupdf as fitz
    out = fitz.open()
    for p in pdf_list:
        with fitz.open(str(p)) as d:
            out.insert_pdf(d)
    out.save(str(dst))
    n = out.page_count
    out.close()
    return n


def _ledger_append(row):
    """同一行记入所有分身为本台账的路径（主份+功能夹副本）。"""
    from openpyxl import Workbook, load_workbook
    for led in LEDGERS:
        if led.exists():
            wb = load_workbook(led)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "交付台账"
            ws.append(LEDGER_HEAD)
        ws.append(row)
        led.parent.mkdir(parents=True, exist_ok=True)
        wb.save(led)


def stage(func, files, matter):
    """登记一批成员文件并重建该（功能+事项）的交付PDF。返回输出路径或None。"""
    files = [Path(f) for f in files if Path(f).exists()]
    matter = re.sub(r'[\\/:*?"<>|\s]', "", str(matter or "")).strip("（）()")
    if not files or not matter:
        return None
    key = f"{func}|{matter}"
    man = _load()
    members = man.setdefault(key, {})
    changed = False
    for f in files:
        h = _sha(f)
        rec = members.get(str(f.resolve()))
        if not rec or rec.get("sha") != h:
            members[str(f.resolve())] = dict(sha=h, mtime=f.stat().st_mtime)
            changed = True
    out_dir = paths.DELIVER / func
    out_dir.mkdir(parents=True, exist_ok=True)
    latest = max(v["mtime"] for v in members.values())
    datep = dt.datetime.fromtimestamp(latest).strftime("%Y%m%d")
    # 统一口径：事项名·金额（成员文件名里的金额求和；识别 金额：X 或 结尾_X 两种）
    amts = []
    for mp in members:
        n = Path(mp).name
        m = re.search(r"金额[：:]\s*(\d+(?:\.\d+)?)", n)
        if not m:
            m2 = re.search(r"_(\d+(?:\.\d+)?)[^_]*$", n)
            # 尾数当金额需≥100，避免把版本副本 _2 / 页码后缀误认成金额
            if m2 and float(m2.group(1)) >= 100:
                m = m2
        if m:
            try:
                amts.append(float(m.group(1)))
            except ValueError:
                pass
    amt_txt = ""
    if amts:
        s = round(sum(amts), 2)
        amt_txt = "·金额：" + (str(int(s)) if s == int(s) else f"{s:.2f}")
    dst = out_dir / f"{datep}{func}（{matter}{amt_txt}）.pdf"
    sig = "|".join(sorted(f"{k.split('/')[-1]}:{v['sha']}" for k, v in members.items()))
    if not changed and man.get("_sig", {}).get(key) == sig and dst.exists():
        _save(man)
        return None                       # 内容没变，不重复出件
    _save(man)
    TMP.mkdir(parents=True, exist_ok=True)
    pdfs = []
    for p in sorted(members, key=lambda x: Path(x).name):
        conv = _to_pdf(p, TMP)
        if conv:
            pdfs.append(conv)
    if not pdfs:
        return None
    n = _merge(pdfs, dst)
    man.setdefault("_sig", {})[key] = sig
    _save(man)
    _ledger_append([dt.datetime.now().strftime("%Y-%m-%d %H:%M"), func, matter, dst.name,
                    len(pdfs), n,
                    "、".join(Path(p).name for p in sorted(members, key=lambda x: Path(x).name)),
                    "重建" if changed else ""])
    print(f"[交付] {func}/{dst.name}（{len(pdfs)}件·{n}页）")
    try:
        for t in TMP.glob("*.pdf"):
            t.unlink()
    except OSError:
        pass
    return dst


# 兼容旧调用
def copy(src, func):
    return stage(func, [src], Path(src).stem.split("（")[-1].rstrip("）") or Path(src).stem)
