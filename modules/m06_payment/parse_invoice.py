# -*- coding: utf-8 -*-
"""电子发票解析器：从 dzfp_*.pdf 抽 发票号码/开票日期/价税合计/销售方(收款方)/
不含税金额/税额/销方开户行+账号。供 m06 付款台账与 fill_payment 使用。

用法：python parse_invoice.py <发票.pdf>
"""
import re
import sys
from pathlib import Path


def read_text(path):
    import pymupdf
    d = pymupdf.open(path)
    return "\n".join(d[p].get_text() for p in range(d.page_count))


def _norm_date(s):
    m = re.search(r"(\d{4})\D{0,2}(\d{1,2})\D{0,2}(\d{1,2})", s)
    if m:
        return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    return ""


def parse(path):
    t = read_text(path)
    flat = re.sub(r"\s+", "", t)
    out = {"发票号码": "", "开票日期": "", "价税合计": "", "不含税金额": "",
           "税额": "", "销售方": "", "购买方": "", "销方开户行": "", "销方账号": ""}
    # 发票号码：20 位数字（数电票）或传统号码
    m = re.search(r"发票号码[：:]?(\d{8,20})", flat) or re.search(r"(\d{20})", flat)
    if m:
        out["发票号码"] = m.group(1)
    # 开票日期（标签与值在 PDF 里分块不相邻，直接全局找日期串）
    m = re.search(r"(\d{4}年\d{1,2}月\d{1,2}日)", flat)
    if m:
        out["开票日期"] = _norm_date(m.group(1))
    # 价税合计：优先“价税合计…（小写）¥X”，否则取所有 ¥ 金额里最大的
    m = re.search(r"价税合计[^¥￥\d]*[¥￥]?\s*([\d,]+\.\d{2})", flat)
    if not m:
        ys = [float(x.replace(",", "")) for x in re.findall(r"[¥￥]([\d,]+\.\d{2})", t)]
        if ys:
            out["价税合计"] = "%.2f" % max(ys)
    if m:
        out["价税合计"] = "%.2f" % float(m.group(1).replace(",", ""))
    # 不含税金额 / 税额（合计行）
    m = re.search(r"合\s*计[¥￥]([\d,]+\.\d{2})\s*[¥￥]([\d,]+\.\d{2})", t)
    if m:
        out["不含税金额"] = "%.2f" % float(m.group(1).replace(",", ""))
        out["税额"] = "%.2f" % float(m.group(2).replace(",", ""))
    # 购买方/销售方名称：公司名（排除本公司）
    ORG = r"[\u4e00-\u9fa5A-Za-z0-9（）()]{2,30}?(?:有限公司|有限责任公司|股份公司|经营部|商行|商店|中心|加工厂|厂)"
    names = re.findall(ORG, t)
    names = [n for n in names if "银行" not in n and "支行" not in n]
    our = "荆州浦华荆清水务有限公司"
    uniq = []
    for n in names:
        if n not in uniq:
            uniq.append(n)
    for n in uniq:
        if n == our and not out["购买方"]:
            out["购买方"] = n
        elif n != our and not out["销售方"]:
            out["销售方"] = n
    # 销方开户行 + 账号（备注里，用分号分隔）
    m = re.search(r"销方开户银行[:：]\s*([^;；\n]+)", t)
    if m:
        out["销方开户行"] = m.group(1).strip()
    m = re.search(r"银行账号[:：]?\s*(\d{10,30})", t)
    if m:
        out["销方账号"] = m.group(1)
    return out


def main():
    if len(sys.argv) < 2:
        print("用法：python parse_invoice.py <发票.pdf>"); return
    r = parse(Path(sys.argv[1]))
    print("=" * 50)
    for k, v in r.items():
        print(f"{k}: {v}")


if __name__ == "__main__":
    main()
