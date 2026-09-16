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
CREATE TABLE IF NOT EXISTS supplier_profile(
  supplier_name TEXT PRIMARY KEY, credit_code TEXT, type TEXT,
  legal_person TEXT, reg_capital TEXT, reg_date TEXT, approval_date TEXT,
  status TEXT, address TEXT, scope TEXT, shareholders TEXT,
  risk TEXT, source TEXT, updated_at TEXT);
CREATE TABLE IF NOT EXISTS supplier_verify(
  id INTEGER PRIMARY KEY, supplier_name TEXT NOT NULL, field TEXT NOT NULL,
  license_value TEXT, online_value TEXT,
  match TEXT CHECK(match IN ('一致','不一致','单边缺失','待核')),
  online_source TEXT, verified_at TEXT,
  UNIQUE(supplier_name, field));
"""

PROFILE_FIELDS = ["credit_code", "type", "legal_person", "reg_capital",
                  "reg_date", "approval_date", "status", "address",
                  "scope", "shareholders", "risk"]

# 报告行定义：(显示名, 字段key)；名称列用供应商名本身
PROFILE_LABELS = [
    ("统一社会信用代码", "credit_code"), ("名称", "name"), ("类型", "type"),
    ("法定代表人（经营者）", "legal_person"), ("注册资本", "reg_capital"),
    ("成立日期（注册日期）", "reg_date"), ("核准日期", "approval_date"),
    ("经营状态", "status"), ("经营场所", "address"), ("经营范围", "scope"),
    ("股东（出资情况）", "shareholders"), ("风险线索", "risk"),
]


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def connect():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.executescript(_SCHEMA)
    try:  # 旧库补列
        conn.execute("ALTER TABLE supplier_profile ADD COLUMN risk TEXT")
    except sqlite3.OperationalError:
        pass
    return conn


def normalize_value(key, v):
    """比对前归一化：统一日期、金额、空白与全半角括号"""
    import re
    s = str(v or "").strip().replace("（", "(").replace("）", ")")
    if not s:
        return ""
    if key in ("reg_date", "approval_date"):
        m = re.search(r"(\d{4})[年./\-]?(\d{1,2})[月./\-]?(\d{1,2})", s)
        if m:
            return f"{int(m.group(1)):04d}{int(m.group(2)):02d}{int(m.group(3)):02d}"
        return s
    if key == "reg_capital":
        m = re.search(r"([0-9.]+)\s*万", s)
        if m:
            return f"{float(m.group(1)):.0f}万"
        cn = {"壹":1,"贰":2,"叁":3,"肆":4,"伍":5,"陆":6,"柒":7,"捌":8,"玖":9,"拾":10,
              "佰":100,"仟":1000,"万":10000,"亿":100000000,"零":0}
        digits = re.findall(r"[壹贰叁肆伍陆柒捌玖拾佰仟万亿零]", s)
        if digits:
            total, section, num = 0, 0, 0
            for ch in digits:
                v2 = cn[ch]
                if v2 >= 10000:
                    section = (section + num) * v2
                    total += section
                    section = num = 0
                elif v2 >= 10:
                    section += num * v2
                    num = 0
                else:
                    num = v2
            return f"{(total + section + num)/10000:.0f}万"
        return s
    return re.sub(r"\s+", "", s)


def upsert_verify(conn, supplier, field, license_value, online_value,
                  match, source="天眼查"):
    conn.execute(
        "INSERT INTO supplier_verify(supplier_name,field,license_value,online_value,"
        "match,online_source,verified_at) VALUES(?,?,?,?,?,?,?) "
        "ON CONFLICT(supplier_name,field) DO UPDATE SET license_value=excluded.license_value,"
        "online_value=excluded.online_value,match=excluded.match,"
        "online_source=excluded.online_source,verified_at=excluded.verified_at",
        (supplier, field, license_value, online_value, match, source, now()))


def get_verifies(conn, supplier):
    rows = conn.execute("SELECT * FROM supplier_verify WHERE supplier_name=?",
                        (supplier,)).fetchall()
    return {r["field"]: dict(r) for r in rows}


def get_or_create_matter(conn, name, category="", note=""):
    row = conn.execute("SELECT * FROM matter WHERE name=?", (name,)).fetchone()
    if row:
        if category and not row["category"]:
            conn.execute("UPDATE matter SET category=? WHERE id=?", (category, row["id"]))
        return row["id"]
    cur = conn.execute("INSERT INTO matter(name,category,note,created_at) VALUES(?,?,?,?)",
                       (name, category, note, now()))
    return cur.lastrowid


def upsert_profile(conn, supplier, data, source=""):
    cols = ["supplier_name"] + PROFILE_FIELDS
    vals = [supplier] + [data.get(f) or "" for f in PROFILE_FIELDS]
    updates = ", ".join(f"{f}=excluded.{f}" for f in PROFILE_FIELDS)
    conn.execute(
        "INSERT INTO supplier_profile(" + ", ".join(cols) + ",source,updated_at) "
        "VALUES(" + ", ".join(["?"] * (len(cols) + 2)) + ") "
        "ON CONFLICT(supplier_name) DO UPDATE SET " + updates +
        ",source=excluded.source,updated_at=excluded.updated_at",
        vals + [source, now()])


def get_profile(conn, supplier):
    row = conn.execute("SELECT * FROM supplier_profile WHERE supplier_name=?",
                       (supplier,)).fetchone()
    return dict(row) if row else None


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
