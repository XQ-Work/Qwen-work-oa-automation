# -*- coding: utf-8 -*-
"""把 build_ocr_cache 中逐张核对过的入库单数据写入 SQLite（receipts / receipt_items）。

用途：本机首次部署时数据库为空，而主流程步骤2 的闸门是“库中已有入库单记录”。
数据来源为对照原图人工核验后的结果，字段与票面一致，不做任何推断补全。
幂等：db.add_receipt 为 INSERT OR REPLACE，重复运行不会产生脏数据。
"""
import importlib.util
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
_root = os.path.dirname(_here)
sys.path.insert(0, _root)
import db

_spec = importlib.util.spec_from_file_location("build_ocr_cache", os.path.join(_here, "build_ocr_cache.py"))
cache_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cache_mod)


def main():
    db.init_db()
    existing = set(db.receipt_ids())
    added = 0
    for r in cache_mod.RECEIPTS:
        rid = r["receipt_id"]
        data = {
            "date": r["date"],
            "supplier": r["supplier"],
            "supplier_short": r["supplier"],
            "items": r["items"],
            "total": r["total"],
        }
        for prefix in ["荆州市", "荆州区", "沙市区", "荆州经济技术开发区"]:
            if data["supplier_short"].startswith(prefix):
                data["supplier_short"] = data["supplier_short"][len(prefix):]
                break
        db.add_receipt(rid, data)
        print("%s %s: %d 项 / %s 元" % ("更新" if rid in existing else "新增", rid, len(r["items"]), r["total"]))
        added += 1
    print("\n完成：%d 张入库单已入库，当前库中共 %d 条" % (added, len(db.receipt_ids())))


if __name__ == "__main__":
    main()
