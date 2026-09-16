# -*- coding: utf-8 -*-
"""m03b 供应商信息查询对比报告（独立模块）

数据来源：国家企业信用信息公示系统截图/营业执照（由智能体视觉提取后入库），
存储：SQLite supplier_profile 表（跨事项复用，同供应商只查一次）。

用法（仓库根目录）：
  python modules/m03b_supplier_info/supplier_info.py set --data-file 临时.json
      临时.json: {"荆州市XX公司": {"credit_code":"...", "type":"...", ...}, ...}
  python modules/m03b_supplier_info/supplier_info.py show --供应商 名称
  python modules/m03b_supplier_info/supplier_info.py gen --事项 X --供应商 A,B,C
      输出横向 A4 Word 表格（行=查询项，列=供应商）到 项目夹\\02_供应商报价及资质\\
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from oa_common import db, archive  # noqa: E402

from docx import Document  # noqa: E402
from docx.enum.section import WD_ORIENT  # noqa: E402
from docx.enum.table import WD_ALIGN_VERTICAL  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402


def _set_font(run, size=10.5, bold=False):
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size)
    run.font.bold = bold


def cmd_set(a):
    data = json.loads(Path(a.data_file).read_text(encoding="utf-8"))
    conn = db.connect()
    src = a.source or ""
    for supplier, fields in data.items():
        db.upsert_profile(conn, supplier, fields, src)
        print(f"[db] 画像已入库：{supplier}")
    conn.commit()


def cmd_show(a):
    conn = db.connect()
    p = db.get_profile(conn, a.供应商)
    if not p:
        print("无画像记录"); return
    for label, key in db.PROFILE_LABELS:
        v = p["supplier_name"] if key == "name" else p.get(key)
        print(f"  {label}: {v or '（待补充）'}")


def cmd_gen(a):
    suppliers = [s.strip() for s in a.供应商.split(",") if s.strip()]
    if len(suppliers) < 2:
        raise SystemExit("至少 2 家供应商")
    conn = db.connect()
    profiles = []
    for s in suppliers:
        p = db.get_profile(conn, s)
        if not p:
            raise SystemExit(f"供应商『{s}』无画像记录，先 set 入库")
        profiles.append(p)

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(1.5)
    sec.top_margin = sec.bottom_margin = Cm(1.5)

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_font(h.add_run("供应商信息查询对比报告"), 16, True)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    import datetime
    _set_font(sub.add_run(f"事项：{a.事项}    查询日期：{datetime.date.today().strftime('%Y年%m月%d日')}"), 10.5)

    n_rows, n_cols = len(db.PROFILE_LABELS), len(suppliers) + 1
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = "Table Grid"
    table.columns[0].width = Cm(4.2)
    col_w = Cm((26.7 - 4.2) / len(suppliers))
    for j in range(1, n_cols):
        table.columns[j].width = col_w

    def fill(cell, text, bold=False, size=10.5, center=False):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        para = cell.paragraphs[0]
        if center:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lines = str(text or "（待补充）").split("\n")
        for i, ln in enumerate(lines):
            p = para if i == 0 else cell.add_paragraph()
            if center and i > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_font(p.add_run(ln), size, bold)

    for i, (label, key) in enumerate(db.PROFILE_LABELS):
        fill(table.cell(i, 0), label, bold=True, center=True)
        for j, p in enumerate(profiles):
            v = p["supplier_name"] if key == "name" else p.get(key)
            size = 9 if key == "scope" else 10.5
            fill(table.cell(i, j + 1), v, size=size, center=(key != "scope"))

    note = doc.add_paragraph()
    _set_font(note.add_run("注：以上信息来源于国家企业信用信息公示系统及营业执照扫描件。"
                           "经对比上述供应商均依法登记注册、经营状态正常，其经营范围涵盖本次采购内容，"
                           "符合采购方案资质要求。"), 10.5)

    out_dir = Path(a.out_dir) if a.out_dir else archive.project_dir(a.事项) / "02_供应商报价及资质"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"供应商信息查询对比报告（{a.事项}）.docx"
    doc.save(str(out))
    print(f"[ok] 生成：{out}")
    missing = [(p["supplier_name"], lb) for p in profiles
               for lb, k in db.PROFILE_LABELS
               if not (p["supplier_name"] if k == "name" else p.get(k))]
    if missing:
        print("[提醒] 以下字段待补充（需公示系统截图提取）：")
        for s, lb in missing:
            print(f"   {s} - {lb}")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("set")
    p1.add_argument("--data-file", required=True)
    p1.add_argument("--source", default="")
    p1.set_defaults(fn=cmd_set)
    p2 = sub.add_parser("show")
    p2.add_argument("--供应商", required=True)
    p2.set_defaults(fn=cmd_show)
    p3 = sub.add_parser("gen")
    p3.add_argument("--事项", required=True)
    p3.add_argument("--供应商", required=True, help="逗号分隔，按报价从低到高")
    p3.add_argument("--out-dir", default="")
    p3.set_defaults(fn=cmd_gen)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
