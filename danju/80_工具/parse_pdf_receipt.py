# -*- coding: utf-8 -*-
"""PDF 版入库单解析（数字化 PDF，自带文字层，免 OCR）。

用 pdfplumber 直提文字层，解析成与图片 OCR 链路同构的入库单数据：
    {receipt_id, date, supplier, items:[{code,name,spec,unit,qty,price,amount}], total}

版式假设（与扫描件一致）：
- 每页一张入库单；页内含 "入库单号: CGSHxxxxxxxxxxxx" 即为新单
- 物料行：编号 名称 [规格] 单位 数量(2位小数) 不含税单价(4位小数) 不含税金额 含税单价(4位小数) 含税金额
- 合计行：合 计： <不含税> <含税>（取后者）
- 无单号的页视为上一单的续页，物料并入
-  scanned PDF（图片型、无文字层）解析不出单号/明细，由主流程提示转人工
"""
import re

# 单位候选：1~3 个汉字且其后紧跟两位小数（数量列）
_UNIT_RE = re.compile(r'^[\u4e00-\u9fa5]{1,3}$')
_CODE_RE = re.compile(r'^[A-Za-z0-9][A-Za-z0-9\-]{4,}$')
_QTY_RE = re.compile(r'^-?\d+\.\d{2}$')
_PRICE_RE = re.compile(r'^-?\d+\.\d{4}$')
_AMOUNT_RE = re.compile(r'^-?\d+(?:\.\d{1,2})?$')
_RID_RE = re.compile(r'(CGSH\d{12,})')
_DATE_RE = re.compile(r'(\d{4})年(\d{1,2})月(\d{1,2})')
_SUP_HINT = ("经营部", "公司", "厂", "店", "个体", "商行", "中心", "合作社",
             "一次性供应商", "临时供应商")


def _page_rid(lines):
    for l in lines:
        m = _RID_RE.search(l.replace(" ", ""))
        if m:
            return m.group(1)
    return ""


def _page_date(lines):
    for l in lines:
        m = _DATE_RE.search(l)
        if m:
            return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    return ""


def _page_supplier(lines):
    for i, l in enumerate(lines):
        if "供应商名称" in l:
            for j in range(i + 1, min(i + 4, len(lines))):
                cand = lines[j].strip()
                if not cand or "制单人" in cand and len(cand) <= 10 and not any(h in cand for h in _SUP_HINT):
                    continue
                first = cand.split()[0] if cand.split() else ""
                # 命中机构后缀关键词，或该词足够长（真实供应商名≥5字，制单人姓名仅2~4字）
                if any(h in first for h in _SUP_HINT) or len(first) >= 5:
                    return first
    return ""


def _page_total(lines):
    for i, l in enumerate(lines):
        if re.search(r'合\s*计', l):
            nums = [float(x) for x in re.findall(r'-?\d+(?:\.\d+)?', l.replace("合计：", " ").replace("合 计：", " ")) if _is_num(x)]
            if nums:
                return nums[-1]
    return 0.0


def _is_num(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def _parse_item_line(line):
    """解析一行物料，返回 dict 或 None（非物料行）。"""
    tokens = line.split()
    if len(tokens) < 6 or not _CODE_RE.match(tokens[0]):
        return None
    code = tokens[0]
    # 锚点：1~3 字汉字单位 + 其后紧跟两位小数的数量
    unit_idx = -1
    for i in range(2, len(tokens) - 1):
        if _UNIT_RE.match(tokens[i]) and _QTY_RE.match(tokens[i + 1]):
            unit_idx = i
            break
    if unit_idx < 0:
        return None
    qty = float(tokens[unit_idx + 1])
    name = tokens[1]
    spec = " ".join(tokens[2:unit_idx])
    nums = tokens[unit_idx + 2:]
    prices = [float(v) for v in nums if _PRICE_RE.match(v)]
    amounts = [float(v) for v in nums if _AMOUNT_RE.match(v) and not _PRICE_RE.match(v)]
    price_no_tax = prices[0] if len(prices) >= 1 else None
    price_tax = prices[1] if len(prices) >= 2 else None
    amount_no_tax = amounts[0] if len(amounts) >= 1 else None
    amount_tax = amounts[1] if len(amounts) >= 2 else None
    if amount_tax is None and price_tax is not None:
        amount_tax = round(price_tax * qty, 2)
    if price_tax is None and price_no_tax is not None:
        price_tax = price_no_tax
    if amount_tax is None and amount_no_tax is not None:
        amount_tax = amount_no_tax
    return {
        "code": code, "name": name, "spec": spec, "unit": tokens[unit_idx],
        "qty": qty,
        "price": price_tax if price_tax is not None else 0,
        "amount": amount_tax if amount_tax is not None else 0,
    }


def _page_items(lines):
    items = []
    for l in lines:
        it = _parse_item_line(l)
        if it:
            items.append(it)
    return items


def parse_page_texts(page_texts):
    """输入每页文字，返回入库单列表（同单号跨页自动合并）。"""
    receipts = []
    cur = None
    for text in page_texts:
        lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
        if not lines:
            continue
        rid = _page_rid(lines)
        items = _page_items(lines)
        total = _page_total(lines)
        if rid:
            cur = {"receipt_id": rid, "date": _page_date(lines),
                   "supplier": _page_supplier(lines), "items": [], "total": 0.0}
            receipts.append(cur)
        if cur is None:
            continue
        cur["items"].extend(items)
        if total and not cur["total"]:
            cur["total"] = total
    # 同单号去重合并（重复打印页）
    merged = {}
    ordered = []
    for r in receipts:
        rid = r["receipt_id"]
        if rid in merged:
            m = merged[rid]
            m["items"].extend(r["items"])
            if r["total"] and not m["total"]:
                m["total"] = r["total"]
            if not m["date"]:
                m["date"] = r["date"]
            if not m["supplier"]:
                m["supplier"] = r["supplier"]
        else:
            merged[rid] = r
            ordered.append(r)
    return ordered


def parse_file(path):
    """解析一个 PDF 文件，返回入库单列表。无文字层时返回空列表。"""
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        page_texts = [(p.extract_text() or "") for p in pdf.pages]
    if not any(t.strip() for t in page_texts):
        return []
    return parse_page_texts(page_texts)
