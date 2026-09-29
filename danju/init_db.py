# -*- coding: utf-8 -*-
"""单据类自动化 v4 -> SQLite 建库与一次性迁移脚本。

数据源（运行时 单据类 下的现有 JSON）：
    - 入库单数据.json  -> receipts + receipt_items
    - 打印文件夹\发票\发票记录.json -> invoices
目标库：本项目目录下 60_数据\单据.db

幂等：每次运行重建表并重新导入（重建式）。
"""
import json
import os
import sqlite3
import sys

BASE = os.environ.get("DJ_BASE") or os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "60_数据")
DB_PATH = os.path.join(DATA_DIR, "单据.db")

RECEIPTS_JSON = os.path.join(BASE, "入库单数据.json")
INVOICES_JSON = os.path.join(BASE, "打印文件夹", "发票", "发票记录.json")

SCHEMA = """
CREATE TABLE receipts (
    receipt_id     TEXT PRIMARY KEY,
    date           TEXT,
    supplier       TEXT,
    supplier_short TEXT,
    total          REAL,
    status         TEXT DEFAULT 'matched',
    created_at     TEXT DEFAULT (datetime('now','localtime')),
    updated_at     TEXT
);
CREATE TABLE receipt_items (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_id TEXT NOT NULL,
    code       TEXT,
    name       TEXT,
    spec       TEXT,
    unit       TEXT,
    qty        REAL,
    price      REAL,
    amount     REAL,
    FOREIGN KEY (receipt_id) REFERENCES receipts(receipt_id) ON DELETE CASCADE
);
CREATE TABLE invoices (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    filename   TEXT UNIQUE,
    seller     TEXT,
    amount     REAL,
    usage      TEXT,
    processed  INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
);
CREATE TABLE meta (
    key   TEXT PRIMARY KEY,
    value TEXT
);
"""


def connect():
    os.makedirs(DATA_DIR, exist_ok=True)
    return sqlite3.connect(DB_PATH)


def migrate(config):
    conn = connect()
    cur = conn.cursor()
    cur.executescript("DROP TABLE IF EXISTS receipt_items;")
    cur.executescript("DROP TABLE IF EXISTS receipts;")
    cur.executescript("DROP TABLE IF EXISTS invoices;")
    cur.executescript("DROP TABLE IF EXISTS meta;")
    cur.executescript(SCHEMA)

    n_receipts = 0
    n_items = 0

    # 1) 入库单数据.json
    if os.path.exists(config["receipts_json"]):
        with open(config["receipts_json"], "r", encoding="utf-8") as f:
            receipts = json.load(f)
        for rid, r in receipts.items():
            cur.execute(
                "INSERT OR REPLACE INTO receipts (receipt_id, date, supplier, supplier_short, total) "
                "VALUES (?,?,?,?,?)",
                (rid, r.get("date", ""), r.get("supplier", ""),
                 r.get("supplier_short", ""), r.get("total", 0)),
            )
            for it in r.get("items", []):
                cur.execute(
                    "INSERT INTO receipt_items (receipt_id, code, name, spec, unit, qty, price, amount) "
                    "VALUES (?,?,?,?,?,?,?,?)",
                    (rid, it.get("code", ""), it.get("name", ""), it.get("spec", "") or "",
                     it.get("unit", ""), it.get("qty", 0), it.get("price", 0), it.get("amount", 0)),
                )
                n_items += 1
            n_receipts += 1
        print("入库单: %d 条, 物料 %d 条" % (n_receipts, n_items))
    else:
        print("跳过: 找不到 %s" % config["receipts_json"])

    # 2) 发票记录.json
    n_invoices = 0
    if os.path.exists(config["invoices_json"]):
        with open(config["invoices_json"], "r", encoding="utf-8") as f:
            invoices = json.load(f)
        for inv in invoices:
            cur.execute(
                "INSERT OR REPLACE INTO invoices (filename, seller, amount) VALUES (?,?,?)",
                (inv.get("文件名", ""), inv.get("销售方名称", ""), inv.get("价税合计", 0)),
            )
            n_invoices += 1
        print("发票: %d 条" % n_invoices)
    else:
        print("跳过: 找不到 %s" % config["invoices_json"])

    cur.execute("INSERT OR REPLACE INTO meta (key, value) VALUES ('schema_version', '1')")
    cur.execute("INSERT OR REPLACE INTO meta (key, value) VALUES ('migrated_at', datetime('now','localtime'))")
    conn.commit()
    conn.close()
    print("迁移完成 -> %s" % DB_PATH)


if __name__ == "__main__":
    has_src = os.path.exists(RECEIPTS_JSON) or os.path.exists(INVOICES_JSON)
    if not has_src and "--force" not in sys.argv:
        n = 0
        if os.path.exists(DB_PATH):
            c = sqlite3.connect(DB_PATH)
            try:
                n = c.execute("SELECT COUNT(*) FROM receipts").fetchone()[0]
            except sqlite3.Error:
                n = 0
            c.close()
        if n > 0:
            print("中止：找不到历史 JSON 数据源（入库单数据.json / 发票记录.json），")
            print("      而库中已有 %d 张入库单。本脚本会重建并清空这些记录。" % n)
            print("      确认为空库重建请加 --force。")
            sys.exit(1)
    cfg = {
        "receipts_json": RECEIPTS_JSON,
        "invoices_json": INVOICES_JSON,
    }
    migrate(cfg)