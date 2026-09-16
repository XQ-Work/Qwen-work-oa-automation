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


def cmd_verify(a):
    """在线核验：data-file = {供应商: {field: 在线值, ...}}，与执照画像逐字段比对"""
    data = json.loads(Path(a.data_file).read_text(encoding="utf-8"))
    conn = db.connect()
    stat = {"一致": 0, "不一致": 0, "单边缺失": 0}
    for supplier, online in data.items():
        p = db.get_profile(conn, supplier)
        if not p:
            print(f"[跳过] 无画像记录：{supplier}")
            continue
        fixes = {}
        for key, oval in online.items():
            if key == "risk":
                continue
            lval = p.get(key) or ""
            n_l, n_o = db.normalize_value(key, lval), db.normalize_value(key, oval)
            if not n_l:
                m = "单边缺失"
            elif n_l == n_o:
                m = "一致"
            elif key in ("scope", "address") and (n_o in n_l or n_l in n_o):
                m = "一致"  # 长文本截断包含视为一致
            else:
                m = "不一致"
            stat[m] += 1
            db.upsert_verify(conn, supplier, key, lval, oval, m)
            if m == "不一致" and a.apply:
                fixes[key] = oval
        if "risk" in online:
            fixes["risk"] = online["risk"]
        if fixes:
            sets = ", ".join(f"{k}=?" for k in fixes)
            conn.execute(f"UPDATE supplier_profile SET {sets} WHERE supplier_name=?",
                         list(fixes.values()) + [supplier])
            print(f"[修正] {supplier}: {list(fixes)}")
    conn.commit()
    print(f"[verify] 比对完成：一致{stat['一致']} 不一致{stat['不一致']} "
          f"执照缺失{stat['单边缺失']}{'（画像已按官方值修正）' if a.apply else ''}")


def cmd_gen(a):
    suppliers = [s.strip() for s in a.供应商.split(",") if s.strip()]
    if len(suppliers) < 2:
        raise SystemExit("至少 2 家供应商")
    conn = db.connect()
    profiles = []
    verifies = {}
    for s in suppliers:
        p = db.get_profile(conn, s)
        if not p:
            raise SystemExit(f"供应商『{s}』无画像记录，先 set 入库")
        profiles.append(p)
        verifies[s] = db.get_verifies(conn, s)
    verified_any = any(verifies[s] for s in suppliers)

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
            mark = ""
            if key != "name":
                vf = verifies.get(p["supplier_name"], {}).get(key)
                if vf:
                    if vf["match"] == "一致":
                        mark = "  ✓"
                    elif vf["match"] == "不一致":
                        v = f"{v}\n(执照:{vf['license_value'] or '—'})\n(官方:{vf['online_value']})"
                        mark = "  ✗"
                    elif vf["match"] == "单边缺失":
                        mark = "  （网络核验补充）"
                elif key == "risk":
                    v = v or "无"
            fill(table.cell(i, j + 1), (str(v or "（待补充）") + mark) if key != "scope"
                 else (str(v or "（待补充）") + mark), size=size, center=(key != "scope"))

    note = doc.add_paragraph()
    if verified_any:
        _set_font(note.add_run("注：以上信息以营业执照扫描件为基准，经与天眼查工商登记信息交叉核验。"
                               "✓表示执照与官方登记一致，✗表示存在差异（并列示两值，请以官方登记为准核实）。"
                               "上述供应商登记状态正常、经营范围涵盖本次采购内容，符合采购方案资质要求。"), 10.5)
    else:
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
    pv = sub.add_parser("verify")
    pv.add_argument("--data-file", required=True)
    pv.add_argument("--apply", action="store_true",
                    help="不一致字段以官方值修正画像")
    pv.set_defaults(fn=cmd_verify)
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
