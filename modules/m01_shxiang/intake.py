# -*- coding: utf-8 -*-
"""收件箱甄别 -> 台账行生成
用法：
  python intake.py scan [--dir 目录]     扫描分组+渲染扫描件，输出 manifest.json
  python intake.py add --json rows.json  把提取好的行写入台账（状态=待确认）
提取字段由智能体视觉识别完成，本脚本负责分组、渲染、写入。
"""
import argparse
import json
import re
import sys
from pathlib import Path
import sys as _sys
_sys.path.insert(0, r"D:/自动备份/Qwen work/自动化（OA流程）/oa-automation")
from oa_common import paths

BASE = paths.STATE
DEFAULT_INBOX = paths.INBOX
OCR_DIR = BASE / "ocr_intake"
LEDGER = paths.LEDGER_OA
PDF_EXT = {".pdf"}
DOC_EXT = {".doc", ".docx", ".xls", ".xlsx"}

# 台账列(v3): B事由 C类型 D金额 E标的 F规格 G数量 H现状 I费用 J概述覆盖 K内容覆盖 L附件 M状态
TYPE_MAP = {"货物": "货物类采购事项", "服务": "服务类采购事项", "工程": None}  # 工程需人工定



def paren_key(stem):
    """文件名括号内项目名作为分组键"""
    m = re.findall(r"[（(]([^（）()]+)[）)]", stem)
    if m:
        key = re.sub(r"[-_\s]*金额[:：]\s*[\d.]+\s*$", "", m[0])
        return re.sub(r"\s+", "", key)
    return None


def parse_name_hint(files):
    """从文件名提取事项名与金额：申购单（PE配件-金额：999.91）
    返回 {name, amount, tier} 或 None；tier: small=≤2000 事项审批即完结"""
    SMALL_LIMIT = 2000
    for fp in files:
        stem = Path(fp).stem
        m = re.search(r"[（(](.+?)[-\s]*金额[:：]\s*([\d.]+)[）)]", stem)
        if m:
            name = re.sub(r"[-_\s]+$", "", m.group(1)).strip()
            amt = float(m.group(2))
            return {"name": name, "amount": amt,
                    "tier": "small" if amt <= SMALL_LIMIT else "normal"}
    return None


def scan(inbox):
    if not inbox.is_dir():
        print("目录不存在:", inbox)
        return 2
    groups, loose = {}, []
    for f in sorted(inbox.iterdir()):
        if not f.is_file() or f.name.startswith("~"):
            continue
        k = paren_key(f.stem)
        if k:
            groups.setdefault(k, []).append(str(f))
        else:
            loose.append(str(f))
    OCR_DIR.mkdir(parents=True, exist_ok=True)
    manifest = {"inbox": str(inbox), "groups": {}, "loose": loose}
    import fitz
    for k, files in groups.items():
        gdir = OCR_DIR / re.sub(r"[\\/:*?\"<>|]", "_", k)
        gdir.mkdir(exist_ok=True)
        pages, docs = [], []
        for fp in files:
            p = Path(fp)
            if p.suffix.lower() in PDF_EXT:
                try:
                    doc = fitz.open(fp)
                    for i, pg in enumerate(doc):
                        out = gdir / f"{p.stem}_p{i+1}.png"
                        pg.get_pixmap(dpi=200).save(out)
                        pages.append(str(out))
                except Exception as e:
                    docs.append({"file": fp, "error": str(e)[:100]})
            elif p.suffix.lower() in DOC_EXT:
                docs.append({"file": fp, "note": "Office文档，需另行提取文本"})
        manifest["groups"][k] = {"files": files, "pages": pages, "docs": docs,
                                 "hint": parse_name_hint(files)}
        h = manifest["groups"][k]["hint"]
        tag = f" 金额提示={h['amount']}({h['tier']})" if h else ""
        print(f"[group] {k}: {len(files)}个文件 -> {len(pages)}页已渲染{tag}")
    if loose:
        print(f"[loose] {len(loose)}个文件括号内无项目名，需人工指定归属")
    mfp = BASE / "intake_manifest.json"
    mfp.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print("manifest ->", mfp)
    return 0


def parse_amount(s):
    """'约1.6万元'->16000.0 ; '8000元'->8000.0 ; 无则 None"""
    if not s:
        return None
    s = str(s).replace(",", "")
    m = re.search(r"([\d.]+)\s*万\s*元?", s)
    if m:
        return round(float(m.group(1)) * 10000, 2)
    m = re.search(r"([\d.]+)\s*元", s)
    if m:
        return round(float(m.group(1)), 2)
    return None


def add_rows(rows_json):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    rows = json.loads(Path(rows_json).read_text(encoding="utf-8"))
    wb = load_workbook(LEDGER)
    ws = wb["发起台账"]
    # 找第一个空行（C列为空且第2行起）
    r = 2
    while ws.cell(row=r, column=2).value not in (None, ""):
        r += 1
    start = r
    yellow = PatternFill("solid", fgColor="FFF2CC")
    for row in rows:
        if row.get("金额来源") and row.get("金额") in (None, ""):
            row["金额"] = parse_amount(row["金额来源"])
        vals = {2: row.get("事由", ""),
                3: row.get("事项类型", ""), 4: row.get("金额", ""),
                5: row.get("标的", ""), 6: row.get("规格", ""),
                7: row.get("数量", ""), 8: row.get("现状原因", ""),
                9: row.get("费用依据", ""), 13: "待确认",
                16: row.get("备注", "intake自动生成，请核对")}
        for col, v in vals.items():
            c = ws.cell(row=r, column=col, value=v)
            c.font = Font(name="微软雅黑", size=10)
            c.alignment = Alignment(vertical="top", wrap_text=True)
        r += 1
    wb.save(LEDGER)
    print(f"已写入台账第{start}~{r-1}行（状态=待确认），共{r-start}行")
    return 0


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    s1 = sub.add_parser("scan")
    s1.add_argument("--dir", default=None)
    s2 = sub.add_parser("add")
    s2.add_argument("--json", required=True)
    args = ap.parse_args()
    if args.cmd == "scan":
        return scan(Path(args.dir) if args.dir else DEFAULT_INBOX)
    return add_rows(args.json)


if __name__ == "__main__":
    sys.exit(main())
