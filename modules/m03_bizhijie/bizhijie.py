# -*- coding: utf-8 -*-
"""m03 比质比价报告单生成（吸收 TRAE 旧方案的填充逻辑与措辞，数据层改 SQLite）

用法（在仓库根目录）：
  python modules/m03_bizhijie/bizhijie.py add-quote --事项 X --供应商 A --单位 批 --数量 1 --价格 8943.00 [--pdf 路径]
  python modules/m03_bizhijie/bizhijie.py list --事项 X
  python modules/m03_bizhijie/bizhijie.py gen --事项 X [--类别 批量|服务] [--品名 显示名]
      读取 DB 中该事项全部报价（2~3 家，按价升序）-> 套模板 -> 输出 .doc+.docx
      到 采购项目档案\\{自动编号|既有夹}\\03_比质比价报告单\\，并登记 report 表
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from oa_common import paths, db  # noqa: E402

TPL_BATCH = paths.TPL_BIZ / "比质比价报告单（批量）.doc"
TPL_SERVICE = paths.TPL_BIZ / "比质比价报告单（服务）.doc"


def _norm(s):
    return re.sub(r"[（）()【】\[\]\s\-_·、,，]", "", s or "")


def _num(price):
    m = re.search(r"[\d.]+", (price or "").replace(",", ""))
    return float(m.group()) if m else None


def archive_dir_for(matter_name, create=True):
    """命中既有项目夹（名含事项核心词）则复用，否则当年最大序号+1 新建"""
    keys = [matter_name]
    for suf in ("申购单", "申购", "采购", "服务", "及耗材", "明细"):
        for k in list(keys):
            if k.endswith(suf) and len(k) - len(suf) >= 2:
                keys.append(k[:-len(suf)])
    nk = [_norm(k) for k in keys if len(_norm(k)) >= 2]
    if paths.ARCHIVE.is_dir():
        for d in sorted(paths.ARCHIVE.iterdir(), reverse=True):
            if d.is_dir() and any(k in _norm(d.name) for k in nk if len(k) >= 3):
                return d / "03_比质比价报告单"
    year = __import__("datetime").date.today().year
    mx = 0
    if paths.ARCHIVE.is_dir():
        for d in paths.ARCHIVE.iterdir():
            m = re.match(rf"^{year}-(\d{{3}})", d.name)
            if m:
                mx = max(mx, int(m.group(1)))
    if not create:
        return None
    return paths.ARCHIVE / f"{year}-{mx + 1:03d}-{matter_name}" / "03_比质比价报告单"


def cmd_add_quote(a):
    conn = db.connect()
    db.add_quote(conn, a.事项, a.供应商, a.单位, a.数量, a.价格, _num(a.价格), a.pdf or "")
    conn.commit()
    n = len(db.quotes_sorted(conn, a.事项))
    print(f"[db] 报价已记录：{a.事项} <- {a.供应商} {a.价格}（当前共 {n} 家）")


def cmd_list(a):
    conn = db.connect()
    for q in db.quotes_sorted(conn, a.事项):
        print(f"  {q['supplier_name']:30s} 单位:{q['unit'] or '-':4s} 数量:{q['qty'] or '-':4s} 价格:{q['price']}")


def cmd_gen(a):
    conn = db.connect()
    quotes = db.quotes_sorted(conn, a.事项)
    if not (2 <= len(quotes) <= 3):
        raise SystemExit(f"报价需 2~3 家，当前 {len(quotes)} 家")
    matter = conn.execute("SELECT * FROM matter WHERE name=?", (a.事项,)).fetchone()
    cat = a.类别 or ("服务" if (matter and matter["category"] == "服务") else "批量")
    product = a.品名 or a.事项
    is_service = cat == "服务"
    tpl = TPL_SERVICE if is_service else TPL_BATCH
    if not tpl.exists():
        raise SystemExit(f"模板缺失：{tpl}")

    low = quotes[0]
    req_note = a.要求 or (matter["note"] if matter and matter["note"] else "")
    if a.要求 and matter:
        conn.execute("UPDATE matter SET note=? WHERE id=?", (a.要求, matter["id"]))
        conn.commit()
    out_dir = archive_dir_for(a.事项)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_doc = out_dir / f"比质比价报告单（{product}）.doc"
    out_docx = out_dir / f"比质比价报告单（{product}）.docx"

    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(str(tpl))
        table = doc.Tables(1)
        for i, q in enumerate(quotes):
            r = i + 2
            table.Cell(r, 1).Range.Text = q["supplier_name"]
            if is_service:
                # 服务模板：品名第二行带简洁服务要求（书写区域有限，要求务必简短）
                name_text = product + (f"\r{req_note}" if req_note else "")
            else:
                name_text = f"{product}（详情见附件清单）"
            table.Cell(r, 2).Range.Text = name_text
            table.Cell(r, 3).Range.Text = q["unit"] or ""
            table.Cell(r, 4).Range.Text = q["qty"] or ""
            table.Cell(r, 5).Range.Text = q["price"] or ""
        tail = f"单价{low['price']}" if is_service else f"总价{low['price']}元"
        # 说明句按旧规则使用事项名
        comparison = (
            f"根据事项审批及供应商报价单，{a.事项}比质比价详情如下：\n"
            "1.经查询上述三家供应商均符合采购方案资质要求。\n"
            "2.上述报价包含：设备、税费、运杂费、保险费、质保、售后服务等一切费用。\n"
            f"3.综合价格对比建议采纳{low['supplier_name']}，{tail}。")
        table.Cell(5, 2).Range.Text = comparison
        table.Cell(6, 2).Range.Text = f"建议选择{low['supplier_name']}。"
        for f in (out_docx, out_doc):
            f.exists() and f.unlink()
        doc.SaveAs2(str(out_doc))
        doc.SaveAs2(str(out_docx), FileFormat=16)
        doc.Close(False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    conn.execute(
        "INSERT INTO report(matter_id,template_type,recommended,output_path,status,created_at)"
        " VALUES(?,?,?,?,?,?)",
        (matter["id"], cat, low["supplier_name"], str(out_doc), "待确认", db.now()))
    conn.commit()
    try:
        from oa_common import deliver      # 交付区多存一份（docx 优先）
        deliver.stage("比质比价报告单", [out_docx if out_docx.exists() else out_doc], product)
    except Exception as e:
        print("  [交付] 副本失败(不影响生成):", str(e)[:60])
    print(f"[ok] 生成（{cat}）：{out_doc}")
    print(f"[推荐] 最低价 {low['supplier_name']} {low['price']} —— 请核对后确认")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("add-quote")
    p1.add_argument("--事项", required=True)
    p1.add_argument("--供应商", required=True)
    p1.add_argument("--单位", default="批")
    p1.add_argument("--数量", default="1")
    p1.add_argument("--价格", required=True)
    p1.add_argument("--pdf", default="")
    p1.set_defaults(fn=cmd_add_quote)
    p2 = sub.add_parser("list")
    p2.add_argument("--事项", required=True)
    p2.set_defaults(fn=cmd_list)
    p3 = sub.add_parser("gen")
    p3.add_argument("--事项", required=True)
    p3.add_argument("--类别", choices=["批量", "服务"], default="")
    p3.add_argument("--品名", default="")
    p3.add_argument("--要求", default="", help="服务类品名列第二行的简洁服务要求")
    p3.set_defaults(fn=cmd_gen)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
