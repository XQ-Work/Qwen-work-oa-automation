# -*- coding: utf-8 -*-
r"""单据库维护工具 manage_db.py（D6：缺单号补录 / 查库 / 补发票用途）

用法（在运行环境根目录运行，其余路径带 80_工具\ 前缀亦可）：
  python manage_db.py list                          # 列出全部入库单
  python manage_db.py list -t invoice               # 列出全部发票（含用途/是否已出付款单）
  python manage_db.py show <receipt_id>             # 查看单张入库单含物料明细
  python manage_db.py add <receipt_id>              # 缺单号补录（交互式逐项录入，最稳妥）
  python manage_db.py add <receipt_id> --date 2026-08-13 --supplier XX --short XX --total 1234 \
         --item "物料|规格|单位|数量|单价|金额" --item ...
                                                    # 非交互式补录（--item 可多次）
  python manage_db.py usage <发票filename> <用途>    # 补/改发票用途（申购单未匹配到时用）
  python manage_db.py total <receipt_id> <金额>      # 修正入库单合计金额

说明：补录数据需来自入库单扫描件真实内容，严禁编造。单号缺失或物料无法识别时，请先看
扫描件并在此手工核对后录入。
"""
import sys, os, argparse
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import db


def fl(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def cmd_list(args):
    if args.t == "invoice":
        rows = db.all_invoices()
        if not rows:
            print("（无发票记录）")
            return
        print("%-3s %-9s %-9s %-10s %s" % ("id", "已付款", "金额", "用途", "seller / filename"))
        for r in rows:
            print("%-3d %-9s %-9.0f %-10s %s  %s" % (
                r["id"], r["processed"], r["amount"] or 0, (r["usage"] or ""), r["seller"] or "", r["filename"]))
    else:
        ids = db.receipt_ids()
        if not ids:
            print("（无入库单记录）")
            return
        for rid in sorted(ids):
            r = db.get_receipt(rid)
            print("%-20s %-10s %s(%s)  合计%.0f  物料%d" % (
                rid, r["date"] or "", r["supplier"] or "", r["supplier_short"] or "", r["total"] or 0, len(r["items"])))


def cmd_show(args):
    r = db.get_receipt(args.id)
    if not r:
        print("未找到单号：%s" % args.id)
        return
    print("单号:%s 日期:%s 供应商:%s 简称:%s 合计:%s  状态:%s" % (
        r["receipt_id"], r["date"] or "", r["supplier"] or "", r["supplier_short"] or "", r["total"], r["status"]))
    if not r["items"]:
        print("（无物料明细）")
        return
    print("%-14s %-12s %-10s %-4s %8s %8s %8s" % ("名称", "规格", "单位", "数量", "单价", "金额", "编码"))
    for it in r["items"]:
        print("%-14s %-12s %-10s %-6g %8g %8g  %s" % (
            it["name"] or "", it["spec"] or "", it["unit"] or "", fl(it["qty"]),
            fl(it["price"]), fl(it["amount"]), it["code"] or ""))


def _guard_total(receipt_id, items, header_total):
    """补录闸门：明细合计必须与入库单表头合计对上（±1 元容差）。
    主流程的 cross_validate_items 只长在 一键生成.py 里，补录侧门必须自带，
    否则看图漏行这类错误会无声入库（CGSH202609220003 漏"统一重负荷齿轮油"即此教训）。"""
    s = round(sum(float(i.get("amount") or 0) for i in items), 2)
    if header_total is None:
        raise SystemExit("[拒写] %s：补录必须提供表头合计 --total（以原图为准），"
                         "防止漏行无声入库" % receipt_id)
    if abs(s - float(header_total)) > 1:
        raise SystemExit("[拒写] %s：明细合计 %.2f ≠ 表头合计 %.2f（差 %.2f），"
                         "请对照原图补行或改 --total" % (
                             receipt_id, s, float(header_total), float(header_total) - s))
    for i, it in enumerate(items, 1):
        q, p, a = (float(it.get(k) or 0) for k in ("qty", "price", "amount"))
        if q > 0 and p > 0 and abs(q * p - a) > 0.5:
            print("  WARNING 第%d项 %s: 数量%g×单价%.2f≠金额%.2f" % (i, it.get("name", "?"), q, p, a))
    print("  闸门通过: %d 项合计 %.2f = 表头 %.2f" % (len(items), s, float(header_total)))


def _interactive_add(receipt_id, args):
    print("交互式补录单号 %s （输入 q 结束物料录入）" % receipt_id)
    date = input("日期(YYYY-MM-DD) [%s]: " % (args.date or "")).strip() or (args.date or "")
    supplier = input("供应商全称 [%s]: " % (args.supplier or "")).strip() or (args.supplier or "")
    short = input("供应商简称 [%s]: " % (args.short or "")).strip() or (args.short or "")
    header_total = input("入库单表头合计金额(必填，以原图为准): ").strip()
    if not header_total:
        raise SystemExit("[拒写] 未提供表头合计，无法核对漏行")
    items = []
    while True:
        line = input("物料(名称|规格|单位|数量|单价): ").strip()
        if line.lower() == "q":
            break
        p = [x.strip() for x in line.split("|")]
        if len(p) < 4:
            print("  (格式: 名称|规格|单位|数量|单价)")
            continue
        name, spec, unit, qty = p[0], p[1], p[2], fl(p[3])
        price = fl(p[4]) if len(p) > 4 else 0
        items.append(dict(code=p[5] if len(p) > 5 else "", name=name, spec=spec, unit=unit,
                          qty=qty, price=price, amount=round(qty * price, 2)))
    total = sum(i["amount"] for i in items)
    data = dict(date=date, supplier=supplier, supplier_short=short,
                total=float(header_total), items=items)
    _guard_total(receipt_id, items, float(header_total))
    db.add_receipt(receipt_id, data)
    print("已写入单号 %s（物料 %d 条，合计 %.2f）" % (receipt_id, len(items), data["total"]))


def cmd_add(args):
    if args.item:
        data = dict(date=args.date or "", supplier=args.supplier or "",
                    supplier_short=args.short or "", total=args.total if args.total is not None else 0, items=[])
        for s in args.item:
            p = [x.strip() for x in s.split("|")]
            name = p[0] if len(p) > 0 else ""
            spec = p[1] if len(p) > 1 else ""
            unit = p[2] if len(p) > 2 else ""
            qty = fl(p[3]) if len(p) > 3 else 0
            price = fl(p[4]) if len(p) > 4 else 0
            code = p[5] if len(p) > 5 else ""
            data["items"].append(dict(code=code, name=name, spec=spec, unit=unit,
                                      qty=qty, price=price, amount=round(qty * price, 2)))
        _guard_total(args.id, data["items"], args.total)
        data["total"] = float(args.total)
        db.add_receipt(args.id, data)
        print("已写入单号 %s（物料 %d 条，合计 %.2f）" % (args.id, len(data["items"]), data["total"]))
    else:
        _interactive_add(args.id, args)


def cmd_usage(args):
    db.update_invoice_usage(args.filename, args.usage)
    print("已为发票 %s 设置用途: %s" % (args.filename, args.usage))


def cmd_total(args):
    conn = db.connect()
    conn.execute("UPDATE receipts SET total=? WHERE receipt_id=?", (args.amount, args.id))
    conn.commit()
    conn.close()
    print("已将单号 %s 合计金额改为 %s" % (args.id, args.amount))


def cmd_audit(args):
    """全库对账：每张入库单 表头合计 vs 明细合计。漏行/差一分钱都揪出来。"""
    bad = 0
    for rid in db.receipt_ids():
        r = db.get_receipt(rid)
        items = r.get("items") or []
        s = round(sum(float(i.get("amount") or 0) for i in items), 2)
        t = float(r.get("total") or 0)
        if abs(s - t) > 1:
            bad += 1
            print("[不符] %s 表头%.2f vs 明细%d行合计%.2f（差%.2f）" % (rid, t, len(items), s, t - s))
    print("审计完成：%d 张入库单，%d 张不符" % (len(db.receipt_ids()), bad))
    raise SystemExit(1 if bad else 0)


def main():
    ap = argparse.ArgumentParser(description="单据库维护工具")
    sub = ap.add_subparsers(dest="cmd")

    p = sub.add_parser("list", help="列出入库单；-t invoice 列发票")
    p.add_argument("-t", dest="t", default="receipt", help="receipt|invoice")
    p.set_defaults(fn=cmd_list)

    p = sub.add_parser("show", help="查看单张入库单")
    p.add_argument("id")
    p.set_defaults(fn=cmd_show)

    p = sub.add_parser("add", help="缺单号补录（交互式或 --item，须带 --total 表头合计对账）")
    p.add_argument("id")
    p.add_argument("--date")
    p.add_argument("--supplier")
    p.add_argument("--short")
    p.add_argument("--total", type=fl)
    p.add_argument("--item", action="append")
    p.set_defaults(fn=cmd_add)

    p = sub.add_parser("audit", help="全库对账：表头合计 vs 明细合计")
    p.set_defaults(fn=cmd_audit)

    p = sub.add_parser("usage", help="补/改发票用途")
    p.add_argument("filename")
    p.add_argument("usage")
    p.set_defaults(fn=cmd_usage)

    p = sub.add_parser("total", help="修正入库单合计金额")
    p.add_argument("id")
    p.add_argument("amount", type=fl)
    p.set_defaults(fn=cmd_total)

    args = ap.parse_args()
    if not hasattr(args, "fn"):
        ap.print_help()
        return
    args.fn(args)


if __name__ == "__main__":
    main()