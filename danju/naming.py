# -*- coding: utf-8 -*-
"""统一文件命名（单据类自动化 v4 重构版）。

规范（已拍板）：
- 中文前缀；下划线 `_` 分隔字段；日期 `YYYYMMDD`；金额取整数元、无小数。
- 避免 Windows 非法字符与空格。
脚本内不散写文件名，统一调用本模块。
"""
import re

_ILLEGAL = re.compile(r'[\\/:*?"<>|\s]+')


def safe(text):
    """去除文件名非法字符（保留中文、英文、数字、下划线、连字符、括号）。"""
    if text is None:
        return ""
    return _ILLEGAL.sub("", str(text)).strip("._")


def compact_date(date):
    """YYYY-MM-DD 或 YYYY年M月D日 -> YYYYMMDD。失败返回原串去连字符。"""
    if not date:
        return ""
    d = str(date).strip()
    m = re.search(r'(\d{4})[-年/.]?(\d{1,2})[-月/.]?(\d{1,2})', d)
    if m:
        return "%s%02d%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    return re.sub(r'[-/:.]', '', d)


def int_money(amount):
    """金额取整数元。"""
    try:
        return int(round(float(amount)))
    except (TypeError, ValueError):
        return 0


# ---------- 手动输入文件 ----------
def scan_name(receipt_id, date, supplier_short, ext="jpg"):
    return "入库单扫描件_%s_%s_%s.%s" % (safe(receipt_id), compact_date(date), safe(supplier_short), ext)


def invoice_name(supplier_short, amount, date=""):
    n = "发票_%s_%s" % (safe(supplier_short), int_money(amount))
    cd = compact_date(date)
    return ("%s_%s.pdf" % (n, cd)) if cd else ("%s.pdf" % n)


def purchase_name(supplier_short, amount):
    return "申购单_%s_%s.pdf" % (safe(supplier_short), int_money(amount))


def approval_name(supplier_short, amount):
    return "事项审批_%s_%s.pdf" % (safe(supplier_short), int_money(amount))


# ---------- 输出文件 ----------
def acceptance_name(receipt_id, date, supplier_short, amount, ext="docx"):
    base = "验收单_%s_%s_%s_%s" % (safe(receipt_id), compact_date(date), safe(supplier_short), int_money(amount))
    return "%s.%s" % (base, (ext or "docx").lstrip(".").lower())


def payment_name(receipt_id, date, supplier_short, amount, ext="xlsx"):
    base = "付款单_%s_%s_%s_%s" % (safe(receipt_id), compact_date(date), safe(supplier_short), int_money(amount))
    return "%s.%s" % (base, (ext or "xlsx").lstrip(".").lower())


def invoice_copy_name(supplier_short, amount, date=""):
    return invoice_name(supplier_short, amount, date)


def invoice_payment_name(supplier_short, amount, ext="xlsx"):
    """发票驱动的付款单（无入库单号，按供应商+金额命名）。"""
    base = "付款单_%s_%s" % (safe(supplier_short), int_money(amount))
    return "%s.%s" % (base, (ext or "xlsx").lstrip(".").lower())


def ledger_name(year=2026):
    return "采购台账_%s.xls" % year


# ---------- 归档 / 日志 ----------
def archive_name(doc_type, date):
    return "归档_%s_%s.pdf" % (safe(doc_type), compact_date(date))


def run_report_name(date, time_compact):
    return "运行报告_%s_%s.log" % (compact_date(date), safe(time_compact))