# -*- coding: utf-8 -*-
"""验收单分页渲染（主流程与 rebuild_outputs 共用）。

模板明细表容量 17 行。单张入库单物料超过 17 项时自动续页：
- 每页复制一份完整模板表格，抬头、验收日期、明细表头、签字栏、
  "遗留问题及处理意见"行在**每一页都完整保留**（不裁表尾）
- 序号跨页连续（…17 | 18…）
返回页数（调用方据此打印提示）。
"""
import copy
import os
import shutil
import zipfile
from lxml import etree

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_NS = 'http://www.w3.org/XML/1998/namespace'
CAP = 17


def _q(tag):
    return '{%s}%s' % (NS_W, tag)


def set_cell_text(tc, text):
    p = tc.find(_q('p'))
    if p is None:
        return
    runs = p.findall(_q('r'))
    t_elems = [r.find(_q('t')) for r in runs]
    t_elems = [t for t in t_elems if t is not None]
    if t_elems:
        t_elems[0].text = text
        t_elems[0].set('{%s}space' % XML_NS, 'preserve')
        r_elem = t_elems[0].getparent()
        rpr = r_elem.find(_q('rPr'))
        if rpr is None:
            rpr = etree.SubElement(r_elem, _q('rPr'))
        rf = rpr.find(_q('rFonts'))
        if rf is None:
            rf = etree.SubElement(rpr, _q('rFonts'))
        rf.set(_q('ascii'), '宋体')
        rf.set(_q('hAnsi'), '宋体')
        rf.set(_q('eastAsia'), '宋体')
        for extra_t in t_elems[1:]:
            extra_t.text = ''
    else:
        r = p.find(_q('r'))
        if r is None:
            r = etree.SubElement(p, _q('r'))
        t = etree.SubElement(r, _q('t'))
        t.text = text
        t.set('{%s}space' % XML_NS, 'preserve')


def _render_page(tbl_src, receipt, page_items, start_seq, is_last, cap=CAP):
    """渲染一页：抬头/日期/表头/17行明细/签字栏每页都完整保留。"""
    tbl = copy.deepcopy(tbl_src)
    rows = tbl.findall(_q('tr'))
    set_cell_text(rows[2].findall(_q('tc'))[3], receipt.get("date", ""))
    for i in range(cap):
        tcs = rows[5 + i].findall(_q('tc'))
        if i < len(page_items):
            it = page_items[i]
            set_cell_text(tcs[0], str(start_seq + i))
            set_cell_text(tcs[1], it["name"])
            set_cell_text(tcs[2], it.get("spec", "") or "")
            set_cell_text(tcs[3], it.get("unit", ""))
            q = it["qty"]
            set_cell_text(tcs[4], str(int(q)) if float(q) == int(float(q)) else str(q))
        else:
            for tc in tcs[:5]:
                set_cell_text(tc, '')
    return tbl


def fill_acceptance_docx(template_dir, receipt, output_path, cap=CAP):
    """由解包模板目录渲染一张验收单 docx（物料>cap 自动续页）。返回页数。"""
    items = receipt.get("items", []) or []
    pages = [items[i:i + cap] for i in range(0, max(len(items), 1), cap)]
    n_pages = len(pages)

    work = output_path + ".wip"
    if os.path.exists(work):
        shutil.rmtree(work)
    shutil.copytree(template_dir, work)
    doc_xml = os.path.join(work, "word", "document.xml")
    tree = etree.parse(doc_xml)
    root = tree.getroot()
    body = root.find(_q('body'))
    children = list(body)

    tbl_idx = next((i for i, ch in enumerate(children) if ch.tag == _q('tbl')), None)
    if tbl_idx is None:
        shutil.rmtree(work, ignore_errors=True)
        raise ValueError("模板缺少明细表格，无法渲染验收单")
    tbl_src = children[tbl_idx]
    if len(tbl_src.findall(_q('tr'))) < 5 + cap:
        shutil.rmtree(work, ignore_errors=True)
        raise ValueError("模板表格行数不足 %d，无法容纳 %d 项明细" % (5 + cap, cap))
    head_ps = children[:tbl_idx]
    sectpr = body.find(_q('sectPr'))

    new_children = list(head_ps)
    for pi, page_items in enumerate(pages):
        is_last = (pi == n_pages - 1)
        new_children.append(_render_page(tbl_src, receipt, page_items,
                                         pi * cap + 1, is_last, cap))
        if not is_last:
            p = etree.Element(_q('p'))
            r = etree.SubElement(p, _q('r'))
            br = etree.SubElement(r, _q('br'))
            br.set(_q('type'), 'page')
            new_children.append(p)
    if sectpr is not None:
        new_children.append(sectpr)

    for old in list(body):
        body.remove(old)
    for ch in new_children:
        body.append(ch)

    tree.write(doc_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for dp, dn, fn in os.walk(work):
            for f in fn:
                fp = os.path.join(dp, f)
                zf.write(fp, os.path.relpath(fp, work))
    shutil.rmtree(work)
    return n_pages
