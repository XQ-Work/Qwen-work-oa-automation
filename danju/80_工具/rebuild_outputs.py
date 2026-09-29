# -*- coding: utf-8 -*-
"""从数据库重建采购台账与验收单 docx。

背景：上一轮跑批存在两处缺陷——台账月份按"第一条记录"归类导致 8 月单据
全部写进 7 月页；验收单漏填"入库单号/供货单位"表头行且漏掉 CGSH202608130006。
本脚本以数据库（已逐张与原图核读一致）为唯一数据源重建全部输出。

用法：python 90_临时\\_rebuild.py
仅生成 台账 xls 与 验收单 docx；PDF 由 80_工具\\to_pdf.py 另行转换。
"""
import os
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import paths
import db
import naming

import xlrd
from xlutils import copy as xcopy
from xlwt import Font, XFStyle
from lxml import etree

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_NS = 'http://www.w3.org/XML/1998/namespace'


def qn_w(tag):
    if ':' in tag:
        prefix, local = tag.split(':', 1)
        return '{%s}%s' % (NS_W, local)
    return '{%s}%s' % (NS_W, tag)


def set_cell_text(tc, text):
    p = tc.find(qn_w('w:p'))
    if p is None:
        return
    runs = p.findall(qn_w('w:r'))
    t_elems = [r.find(qn_w('w:t')) for r in runs]
    t_elems = [t for t in t_elems if t is not None]
    if t_elems:
        t_elems[0].text = text
        t_elems[0].set('{%s}space' % XML_NS, 'preserve')
        r_elem = t_elems[0].getparent()
        rpr = r_elem.find(qn_w('w:rPr'))
        if rpr is None:
            rpr = etree.SubElement(r_elem, qn_w('w:rPr'))
        rf = rpr.find(qn_w('w:rFonts'))
        if rf is None:
            rf = etree.SubElement(rpr, qn_w('w:rFonts'))
        rf.set(qn_w('ascii'), '宋体')
        rf.set(qn_w('hAnsi'), '宋体')
        rf.set(qn_w('eastAsia'), '宋体')
        for extra_t in t_elems[1:]:
            extra_t.text = ''
    else:
        r = p.find(qn_w('w:r'))
        if r is None:
            r = etree.SubElement(p, qn_w('w:r'))
        t = etree.SubElement(r, qn_w('w:t'))
        t.text = text
        t.set('{%s}space' % XML_NS, 'preserve')
        rpr = r.find(qn_w('w:rPr'))
        if rpr is None:
            rpr = etree.SubElement(r, qn_w('w:rPr'))
        rf = rpr.find(qn_w('w:rFonts'))
        if rf is None:
            rf = etree.SubElement(rpr, qn_w('w:rFonts'))
        rf.set(qn_w('ascii'), '宋体')
        rf.set(qn_w('hAnsi'), '宋体')
        rf.set(qn_w('eastAsia'), '宋体')


def main():
    receipt_ids = sorted(db.receipt_ids(),
                         key=lambda rid: (db.get_receipt(rid)["date"], rid))
    print("数据库入库单 %d 单: %s" % (len(receipt_ids), ", ".join(receipt_ids)))

    all_records = []
    all_receipts = []
    for rid in receipt_ids:
        data = db.get_receipt(rid)
        for it in data["items"]:
            all_records.append({
                "code": it.get("code", ""), "name": it["name"],
                "spec": it.get("spec", ""), "unit": it["unit"],
                "qty": it["qty"], "price": it.get("price", 0),
                "amount": it.get("amount", 0), "date": data["date"],
                "receipt": rid, "supplier": data["supplier"], "reason": ""})
        all_receipts.append({
            "id": rid, "date": data["date"],
            "supplier_short": data["supplier_short"],
            "items": data["items"], "total": data["total"]})

    # ---------------- 台账 ----------------
    out_xls = os.path.join(paths.LEDGER_DIR, "采购台账_2026.xls")
    rb = xlrd.open_workbook(paths.TEMPLATE_XLS, formatting_info=True)
    wb = xcopy.copy(rb)
    sheet_names = rb.sheet_names()

    style_song = XFStyle()
    style_song.font = Font()
    style_song.font.name = 'SimSun'
    style_song.font.height = 200
    style_bold = XFStyle()
    style_bold.font = Font()
    style_bold.font.name = 'SimSun'
    style_bold.font.bold = True
    style_bold.font.height = 200

    by_month = {}
    for rec in all_records:
        m = int(rec["date"][5:7])
        by_month.setdefault(m, []).append(rec)

    for m in sorted(by_month):
        label = "%d月" % m
        if label not in sheet_names:
            print("  [WARNING] 模板无工作表 %s，跳过 %d 条" % (label, len(by_month[m])))
            continue
        idx = sheet_names.index(label)
        sheet = rb.sheet_by_index(idx)
        ws = wb.get_sheet(idx)
        recs = by_month[m]
        for i, rec in enumerate(recs):
            vals = [rec["code"], rec["name"], rec["spec"], rec["unit"],
                    rec["qty"], rec["price"], rec["amount"], rec["date"],
                    rec["receipt"], rec["supplier"], rec["reason"]]
            for col, val in enumerate(vals):
                ws.write(1 + i, col, val, style_song)
        total = sum(r["amount"] for r in recs)
        ws.write(1 + len(recs), 6, total, style_bold)
        print("  台账 %s: %d 条, 合计 %.2f" % (label, len(recs), total))

    wb.save(out_xls)
    print("台账已重建: %s" % out_xls)

    # ---------------- 验收单 ----------------
    unpacked = os.path.join(paths.WORK_DIR, "unpacked_tpl")
    if os.path.exists(unpacked):
        shutil.rmtree(unpacked)
    with zipfile.ZipFile(paths.TEMPLATE_DOCX, 'r') as zf:
        zf.extractall(unpacked)

    doc_xml_src = os.path.join(unpacked, 'word', 'document.xml')
    n_docx = 0
    import accept_pages
    for r in all_receipts:
        fname = naming.acceptance_name(r['id'], r['date'], r['supplier_short'], r['total'])
        out_docx = os.path.join(paths.ACCEPT_DIR, fname)
        n_pages = accept_pages.fill_acceptance_docx(unpacked, r, out_docx)
        if n_pages > 1:
            print("  续页: %s 共 %d 项物料，分 %d 页打印" % (r['id'], len(r['items']), n_pages))
        n_docx += 1
        print("  验收单: %s" % fname)
    print("验收单 docx 已重建 %d 个" % n_docx)


if __name__ == "__main__":
    main()
