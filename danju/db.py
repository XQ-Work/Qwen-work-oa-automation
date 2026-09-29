# -*- coding: utf-8 -*-
"""SQLite 数据访问层（单据类自动化 v4 重构版）。

替换原散落的 .json 读写。供 一键生成.py 等使用。
表：
  receipts       入库单主表
  receipt_items  物料明细表
  invoices       发票表（含用途、是否已出付款单）
  meta          schema 版本等
"""
import sqlite3

from paths import DB_PATH

SCHEMA = """
CREATE TABLE IF NOT EXISTS receipts (
    receipt_id     TEXT PRIMARY KEY,
    date           TEXT,
    supplier       TEXT,
    supplier_short TEXT,
    total          REAL,
    status         TEXT DEFAULT 'matched',
    created_at     TEXT DEFAULT (datetime('now','localtime')),
    updated_at     TEXT
);
CREATE TABLE IF NOT EXISTS receipt_items (
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
CREATE TABLE IF NOT EXISTS invoices (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    filename   TEXT UNIQUE,
    seller     TEXT,
    amount     REAL,
    usage      TEXT,
    processed  INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
);
CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT
);
"""


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = connect()
    conn.executescript(SCHEMA)
    conn.commit()
    conn.close()


# ---------- receipts ----------
def get_receipt(receipt_id):
    conn = connect()
    row = conn.execute("SELECT * FROM receipts WHERE receipt_id=?", (receipt_id,)).fetchone()
    if row is None:
        conn.close()
        return None
    r = dict(row)
    r["items"] = [dict(x) for x in conn.execute(
        "SELECT code,name,spec,unit,qty,price,amount FROM receipt_items WHERE receipt_id=?",
        (receipt_id,)).fetchall()]
    conn.close()
    return r


def unpaid_receipts():
    """未生成付款单的入库单（含明细），按日期排序。"""
    conn = connect()
    ids = [r["receipt_id"] for r in conn.execute(
        "SELECT receipt_id FROM receipts WHERE IFNULL(status,'matched')!='paid' "
        "ORDER BY date, receipt_id").fetchall()]
    out = []
    for rid in ids:
        r = get_receipt(rid)
        if r:
            out.append(r)
    conn.close()
    return out


def receipt_ids():
    conn = connect()
    rows = conn.execute("SELECT receipt_id FROM receipts").fetchall()
    conn.close()
    return [r["receipt_id"] for r in rows]


def add_receipt(receipt_id, data):
    """新增/更新一张入库单（含物料）。重复入库时保留 status（paid 不回退）。"""
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO receipts (receipt_id, date, supplier, supplier_short, total) "
        "VALUES (?,?,?,?,?) "
        "ON CONFLICT(receipt_id) DO UPDATE SET "
        "date=excluded.date, supplier=excluded.supplier, "
        "supplier_short=excluded.supplier_short, total=excluded.total, "
        "updated_at=datetime('now','localtime')",
        (receipt_id, data.get("date", ""), data.get("supplier", ""),
         data.get("supplier_short", ""), data.get("total", 0)))
    cur.execute("DELETE FROM receipt_items WHERE receipt_id=?", (receipt_id,))
    for it in data.get("items", []):
        cur.execute(
            "INSERT INTO receipt_items (receipt_id, code, name, spec, unit, qty, price, amount) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (receipt_id, it.get("code", ""), it.get("name", ""), it.get("spec", "") or "",
             it.get("unit", ""), it.get("qty", 0), it.get("price", 0), it.get("amount", 0)))
    conn.commit()
    conn.close()
    return True


def mark_receipt_paid(receipt_id):
    """标记入库单已生成付款单（防重复出单）。"""
    conn = connect()
    conn.execute("UPDATE receipts SET status='paid', updated_at=datetime('now','localtime') "
                 "WHERE receipt_id=?", (receipt_id,))
    conn.commit()
    conn.close()


# ---------- invoices ----------
def all_invoices():
    conn = connect()
    rows = conn.execute("SELECT * FROM invoices ORDER BY id").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_invoice(filename):
    conn = connect()
    row = conn.execute("SELECT * FROM invoices WHERE filename=?", (filename,)).fetchone()
    conn.close()
    return dict(row) if row else None


def add_invoice(filename, seller, amount, usage=""):
    conn = connect()
    conn.execute("INSERT OR REPLACE INTO invoices (filename, seller, amount, usage) VALUES (?,?,?,?)",
                 (filename, seller, amount, usage))
    conn.commit()
    conn.close()


def update_invoice_usage(filename, usage):
    conn = connect()
    conn.execute("UPDATE invoices SET usage=? WHERE filename=?", (usage, filename))
    conn.commit()
    conn.close()


def mark_invoice_processed(filename):
    conn = connect()
    conn.execute("UPDATE invoices SET processed=1 WHERE filename=?", (filename,))
    conn.commit()
    conn.close()


def find_invoice(amount=None, seller=None):
    conn = connect()
    q = "SELECT * FROM invoices WHERE 1=1"
    args = []
    if amount is not None:
        q += " AND amount=?"
        args.append(amount)
    if seller:
        q += " AND seller LIKE ?"
        args.append("%" + seller + "%")
    rows = conn.execute(q, args).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def meta_value(key):
    conn = connect()
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else None


def set_meta(key, value):
    conn = connect()
    conn.execute("INSERT OR REPLACE INTO meta (key, value) VALUES (?,?)", (key, value))
    conn.commit()
    conn.close()