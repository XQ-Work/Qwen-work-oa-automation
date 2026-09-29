# -*- coding: utf-8 -*-
"""付款台账写入模块。

一张付款单按关联到的入库单物料拆成多行展开；每行一条物料，
付款相关字段（付款单号/日期/收款单位/金额/用途/账号/开户行）随行重复。
无法关联物料时写一行、仅填付款字段。
histo 数据：不重复写入同一付款单号（幂等，重跑不重复累积）。
"""

import os
import xlwt
import xlrd

HEADERS = ["付款单号", "付款日期", "收款单位", "付款金额", "用途", "账号", "开户行",
           "物料名称", "规格", "单位", "数量", "单价", "物料金额", "关联入库单号"]
N_COLS = len(HEADERS)


_STYLE_CACHE = {}


def _style(header=False):
    # 复用样式对象，避免逐格新建 XF 触发 xlwt “More than 4094 XFs” 上限
    st = _STYLE_CACHE.get(header)
    if st is None:
        st = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = 'SimSun'
        font.height = 220
        font.bold = header
        st.font = font
        _STYLE_CACHE[header] = st
    return st


def append_payment(path, pay, materials):
    """追加一张付款单到付款台账。

    pay: dict(sign, date, seller, amount, usage, account, bank)
    materials: list[dict(code,name,spec,unit,qty,price,amount,receipt)]，可为空
    """
    os.makedirs(os.path.dirname(path), exist_ok=True)
    exists = os.path.exists(path)
    rs = None
    if exists:
        rb = xlrd.open_workbook(path, formatting_info=False)
        rs = rb.sheet_by_index(0)
        # 若付款单号早已写入，幂等跳过
        sign_col = 0
        for r in range(rs.nrows):
            if str(rs.cell_value(r, sign_col)).strip() == str(pay["sign"]).strip():
                return 0

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('付款台账')
    row = 0
    if exists and rs is not None:
        # 复制已有全部单元格
        for r in range(rs.nrows):
            for c in range(N_COLS):
                v = rs.cell_value(r, c)
                if v not in ("", None):
                    ws.write(r, c, v, _style(r == 0))
            row = rs.nrows
    else:
        for c, h in enumerate(HEADERS):
            ws.write(0, c, h, _style(header=True))
        row = 1

    if materials:
        for m in materials:
            _write_row(ws, row, pay, m)
            row += 1
    else:
        _write_row(ws, row, pay, None)
        row += 1

    wb.save(path)
    return 1


def _write_row(ws, r, pay, m):
    ws.write(r, 0, pay["sign"])
    ws.write(r, 1, pay["date"])
    ws.write(r, 2, pay["seller"])
    ws.write(r, 3, pay["amount"])
    ws.write(r, 4, pay.get("usage", ""))
    ws.write(r, 5, pay.get("account", ""))
    ws.write(r, 6, pay.get("bank", ""))
    if m:
        ws.write(r, 7, m.get("name", ""))
        ws.write(r, 8, m.get("spec", ""))
        ws.write(r, 9, m.get("unit", ""))
        ws.write(r, 10, m.get("qty", ""))
        ws.write(r, 11, m.get("price", ""))
        ws.write(r, 12, m.get("amount", ""))
        ws.write(r, 13, m.get("receipt", ""))


def match_materials(seller, amount, all_records):
    """按供应商匹配物料；若存在金额==发票金额的入库单则仅取该张，否则取该供应商全部记录。"""
    keywords = [k for k in seller.replace("荆州", "").replace("沙市区", "").replace("荆州区", "").replace("市", "").replace("省", "").split(" ") if len(k) >= 2]
    if not keywords:
        return []
    cand = []
    for r in all_records:
        sup = str(r.get("supplier", ""))
        if any(kw in sup for kw in keywords):
            cand.append(r)
    if not cand:
        return []
    # 按入库单号分组求和，找金额==发票金额的单据
    by_receipt = {}
    for r in cand:
        by_receipt.setdefault(r.get("receipt"), []).append(r)
    for rec_no, rows in by_receipt.items():
        total = sum(float(r.get("amount", 0) or 0) for r in rows)
        if abs(total - float(amount)) < 0.01:
            return rows
    return cand