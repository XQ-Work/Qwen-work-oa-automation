# -*- coding: utf-8 -*-
"""合并验收单生成模块。

把多张入库单的验收单合并成一个 DOCX（单与单之间分页）；一张入库单物料
超过一页容量(cap)时自动续页，每页都复制模板表格并带完整抬头（单号/日期/
供货单位），物料与金额只做内部校验、不印金额。

金额校验：某张入库单 物料总金额(∑item.amount) != 入库单总金额(receipt.total)
时，中止该单（不纳入合并），由调用方在报告里列出。
"""

import os
import shutil
import zipfile
import subprocess
import copy
import datetime
import xml.etree.ElementTree as ET

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_NS = 'http://www.w3.org/XML/1998/namespace'
CAP = 17


def _q(t):
    return '{%s}%s' % (NS_W, t)


def _set_cell_text(tc, text):
    p = tc.find(_q('w:p'))
    if p is None:
        return
    runs = p.findall(_q('w:r'))
    t_elems = [r.find(_q('w:t')) for r in runs]
    t_elems = [t for t in t_elems if t is not None]
    if t_elems:
        t_elems[0].text = text
        t_elems[0].set('{%s}space' % XML_NS, 'preserve')
        r_elem = t_elems[0].getparent()
        rpr = r_elem.find(_q('w:rPr'))
        if rpr is None:
            rpr = ET.SubElement(r_elem, _q('w:rPr'))
        rf = rpr.find(_q('w:rFonts'))
        if rf is None:
            rf = ET.SubElement(rpr, _q('w:rFonts'))
        rf.set(_q('w:ascii'), '宋体')
        rf.set(_q('w:hAnsi'), '宋体')
        rf.set(_q('w:eastAsia'), '宋体')
        for extra_t in t_elems[1:]:
            extra_t.text = ''
    else:
        r = p.find(_q('w:r'))
        if r is None:
            r = ET.SubElement(p, _q('w:r'))
        t = ET.SubElement(r, _q('w:t'))
        t.text = text
        t.set('{%s}space' % XML_NS, 'preserve')
        rpr = r.find(_q('w:rPr'))
        if rpr is None:
            rpr = ET.SubElement(r, _q('w:rPr'))
        rf = rpr.find(_q('w:rFonts'))
        if rf is None:
            rf = ET.SubElement(rpr, _q('w:rFonts'))
        rf.set(_q('w:ascii'), '宋体')
        rf.set(_q('w:hAnsi'), '宋体')
        rf.set(_q('w:eastAsia'), '宋体')


def _render_tbl(tbl, receipt, items, start_seq, is_last):
    rows = tbl.findall(_q('w:tr'))
    date_tcs = rows[2].findall(_q('w:tc'))
    _set_cell_text(date_tcs[3], receipt.get('date', ''))
    head_tcs = rows[3].findall(_q('w:tc'))
    supplier = receipt.get('supplier_short') or receipt.get('supplier') or ''
    _set_cell_text(head_tcs[0], "入库单号：%s  供货单位：%s" % (receipt['id'], supplier))
    for i in range(len(items)):
        it = items[i]
        tcs = rows[5 + i].findall(_q('w:tc'))
        _set_cell_text(tcs[0], str(start_seq + i))
        _set_cell_text(tcs[1], it.get('name', ''))
        _set_cell_text(tcs[2], it.get('spec', '') or '')
        _set_cell_text(tcs[3], it.get('unit', ''))
        qty = it.get('qty', '')
        _set_cell_text(tcs[4], str(int(qty)) if isinstance(qty, float) and qty == int(qty) else str(qty))
    if not is_last:
        tail = rows[-3:] if len(rows) >= 24 else []
        for tr in tail:
            tbl.remove(tr)
    return tbl


def _validate(receipts):
    valid, failed = [], []
    for r in receipts:
        total = sum(float(it.get('amount', 0) or 0) for it in r.get('items', []))
        ref = float(r.get('total', 0) or 0)
        if abs(total - ref) < 0.011:
            valid.append(r)
        else:
            failed.append((r.get('id'), total, ref))
    return valid, failed


def render_acceptance(receipts, template_docx, work_dir, soffice):
    valid, failed = _validate(receipts)
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    out_docx = os.path.join(work_dir, "验收单_%s.docx" % stamp)
    out_pdf = os.path.join(work_dir, "验收单_%s.pdf" % stamp)
    if not valid:
        return (None, None, failed)

    wk = os.path.join(work_dir, "_acc_w")
    if os.path.exists(wk):
        shutil.rmtree(wk)
    os.makedirs(wk, exist_ok=True)
    with zipfile.ZipFile(template_docx) as zf:
        zf.extractall(wk)

    doc_path = os.path.join(wk, 'word', 'document.xml')
    tree = ET.parse(doc_path)
    root = tree.getroot()
    body = root.find(_q('w:body'))
    children = list(body)

    head_ps = [ch for ch in children if ch.tag == _q('w:p')][:2]
    tbl_src = None
    for ch in children:
        if ch.tag == _q('w:tbl'):
            tbl_src = ch
            break
    sectpr = body.find(_q('w:sectPr'))

    new_children = list(head_ps)
    for si, rec in enumerate(valid):
        items = rec.get('items', [])
        for start in range(0, max(len(items), 1), CAP):
            slice_items = items[start:start + CAP]
            is_last = (start + CAP >= len(items))
            page_tbl = _render_tbl(copy.deepcopy(tbl_src), rec, slice_items,
                                   start + 1, is_last)
            new_children.append(page_tbl)
            if not (si == len(valid) - 1 and is_last):
                p = ET.SubElement(body, _q('w:p'))
                r = ET.SubElement(p, _q('w:r'))
                br = ET.SubElement(r, _q('w:br'))
                br.set(_q('w:type'), 'page')
                new_children.append(p)
    if sectpr is not None:
        new_children.append(sectpr)

    for old in list(body):
        body.remove(old)
    for ch in new_children:
        body.append(ch)

    tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(out_docx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for dp, dn, fn in os.walk(wk):
            for f in fn:
                fp = os.path.join(dp, f)
                zf.write(fp, os.path.relpath(fp, wk))
    shutil.rmtree(wk, ignore_errors=True)

    subprocess.run([soffice, '--headless', '--convert-to', 'pdf',
                    '--outdir', os.path.dirname(out_pdf), out_docx],
                   capture_output=True)
    return (out_docx, out_pdf, failed)
