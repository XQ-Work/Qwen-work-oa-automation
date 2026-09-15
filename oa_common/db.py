# -*- coding: utf-8 -*-
"""SQLite 共享数据层（全链主数据）：事项/供应商/报价/报告/流程记录。
库文件位于数据目录（config.json data_root），永不入 Git 仓库。"""
import datetime
import sqlite3

from oa_common import paths

DB_PATH = paths.DATA / "oa_automation.db"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS matter(
  id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, category TEXT,
  note TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS supplier(
  id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, note TEXT,
  oa_archived INTEGER DEFAULT 0, created_at TEXT);
CREATE TABLE IF NOT EXISTS quotation(
  id INTEGER PRIMARY KEY, matter_id INT NOT NULL, supplier_name TEXT NOT NULL,
  unit TEXT, qty TEXT, price TEXT, price_num REAL, source_pdf TEXT,
  extracted_at TEXT, UNIQUE(matter_id, supplier_name));
CREATE TABLE IF NOT EXISTS report(
  id INTEGER PRIMARY KEY, matter_id INT NOT NULL, template_type TEXT,
  recommended TEXT, output_path TEXT, status TEXT DEFAULT '待确认',
  created_at TEXT);
CREATE TABLE IF NOT EXISTS oa_flow(
  id INTEGER PRIMARY KEY, matter_id INT, flow_type TEXT NOT NULL,
  requestid TEXT, title TEXT, status TEXT, created_at TEXT);
"""


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def connect():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.executescript(_SCHEMA)
    return conn


def get_or_create_matter(conn, name, category="", note=""):
    row = conn.execute("SELECT * FROM matter WHERE name=?", (name,)).fetchone()
    if row:
        if category and not row["category"]:
            conn.execute("UPDATE matter SET category=? WHERE id=?", (category, row["id"]))
        return row["id"]
    cur = conn.execute("INSERT INTO matter(name,category,note,created_at) VALUES(?,?,?,?)",
                       (name, category, note, now()))
    return cur.lastrowid


def add_quote(conn, matter_name, supplier, unit, qty, price, price_num=None,
              source_pdf=""):
    mid = get_or_create_matter(conn, matter_name)
    conn.execute("INSERT OR IGNORE INTO supplier(name,created_at) VALUES(?,?)",
                 (supplier, now()))
    conn.execute(
        "INSERT INTO quotation(matter_id,supplier_name,unit,qty,price,price_num,"
        "source_pdf,extracted_at) VALUES(?,?,?,?,?,?,?,?) "
        "ON CONFLICT(matter_id,supplier_name) DO UPDATE SET unit=excluded.unit,"
        "qty=excluded.qty,price=excluded.price,price_num=excluded.price_num,"
        "source_pdf=excluded.source_pdf,extracted_at=excluded.extracted_at",
        (mid, supplier, unit, qty, price, price_num, source_pdf, now()))
    return mid


def quotes_sorted(conn, matter_name):
    """按价格升序返回该事项报价（price_num 为空时从 price 文本提取首个数字）"""
    rows = conn.execute(
        "SELECT q.*, m.category FROM quotation q JOIN matter m ON m.id=q.matter_id "
        "WHERE m.name=? ORDER BY COALESCE(q.price_num, -1)", (matter_name,)).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        if d.get("price_num") is None:
            import re
            m = re.search(r"[\d.]+", (d.get("price") or "").replace(",", ""))
            d["_sort"] = float(m.group()) if m else float("inf")
        else:
            d["_sort"] = d["price_num"]
        out.append(d)
    return sorted(out, key=lambda x: x["_sort"])
