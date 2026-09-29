# -*- coding: utf-8 -*-
"""金额链主档（方案A · 第1步）：projects / payments / amount_history 三张表 + 对账。

口径（用户确认）：
  · 金额为"阶段属性"，项目编号(2026-NNN)为全链主键；
  · 付款与合同金额不一致是常态（明细未完全执行）：校验=累计已付≤合同金额，
    差额记"未执行差额"，仅超付算异常；
  · 付款性质按关键字甄别：预付款/质保金/合同款（含简写 预付/质保），取词顺序
    事由或摘要 → 回单摘要 → 文件名，缺省合同款；
  · ≥2000 分流道以合同金额改判（A=有合同，B=无合同）。

挂接点：m05 extract（合同金额）/ m06 fill（付款行）/ m07 match（已付确认）。
CLI：python -m oa_common.moneyline report | show <code>
"""
import datetime as dt
import json
import re
import sqlite3

from oa_common import paths

DB_PATH = paths.DB
SMALL_LIMIT = 2000.0

_SCHEMA = """
CREATE TABLE IF NOT EXISTS projects(
  code TEXT PRIMARY KEY, matter TEXT, category TEXT,
  est_amount REAL, appr_amount REAL, contract_amount REAL,
  paid_total REAL DEFAULT 0, gap REAL, lane TEXT, note TEXT,
  updated_at TEXT);
CREATE TABLE IF NOT EXISTS payments(
  id INTEGER PRIMARY KEY, code TEXT NOT NULL, nature TEXT NOT NULL,
  seq INTEGER DEFAULT 1, amount REAL, payee TEXT,
  ledger TEXT, row INT, requestid TEXT,
  status TEXT DEFAULT '待付款', note TEXT, created_at TEXT,
  UNIQUE(code, nature, seq));
CREATE TABLE IF NOT EXISTS amount_history(
  id INTEGER PRIMARY KEY, code TEXT, stage TEXT,
  old_value REAL, new_value REAL, source TEXT, created_at TEXT);
"""

NATURES = (("预付款", ("预付款", "预付")),
           ("质保金", ("质保金", "质保")),
           ("合同款", ("合同款", "合同")))


def now():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def connect():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.executescript(_SCHEMA)
    try:
        conn.execute("ALTER TABLE projects ADD COLUMN tax TEXT")
    except sqlite3.OperationalError:
        pass
    return conn


def extract_tax(text):
    """从合同正文抓税率：税率13% / 税率（1%） / 税率13 等写法。"""
    import re
    m = re.search(r"税率[\s（(是为:：]*?(\d{1,2}(?:\.\d)?)\s*[%％]?", str(text))
    return (m.group(1) + "%") if m else ""


def set_tax(code, tax):
    if not (code and tax):
        return
    conn = connect()
    if conn.execute("SELECT 1 FROM projects WHERE code=?", (code,)).fetchone():
        conn.execute("UPDATE projects SET tax=? WHERE code=?", (str(tax), code))
        conn.commit()
    conn.close()


# ---------- 项目编号解析 ----------

def code_from_path(path):
    """从 '采购项目档案/2026-005-事项/...' 之类的路径反解 (code, matter)。"""
    m = re.search(r"[\\/](\d{4}-\d{3})-([^\\/]+)[\\/]", str(path) + "\\")
    if m:
        return m.group(1), m.group(2)
    m = re.match(r"^(\d{4})-(\d{3})-(.+)$", str(path).strip())
    return (f"{m.group(1)}-{m.group(2)}", m.group(3)) if m else (None, None)


def resolve_code(matter_kw="", path=""):
    """先按归档路径反解，再按事项关键词找项目夹；返回 (code, matter)。"""
    code, matter = code_from_path(path)
    if code:
        return code, matter
    if matter_kw:
        from oa_common import archive
        d = archive.project_dir(str(matter_kw), create=False)
        if d:
            return code_from_path(d.name)
    return None, matter_kw or ""


def nature_of(*texts):
    """按取词顺序 事由/摘要→回单摘要→文件名 识别付款性质，缺省合同款。"""
    for t in texts:
        s = re.sub(r"\s+", "", str(t or ""))
        for canon, kws in NATURES:
            if any(k in s for k in kws):
                return canon
    return "合同款"


# ---------- 主档写入 ----------

def upsert_project(code, matter=None, **fields):
    conn = connect()
    row = conn.execute("SELECT * FROM projects WHERE code=?", (code,)).fetchone()
    if row is None:
        conn.execute("INSERT INTO projects(code,matter,paid_total,updated_at) VALUES(?,?,0,?)",
                     (code, matter or "", now()))
        row = conn.execute("SELECT * FROM projects WHERE code=?", (code,)).fetchone()
    sets = {}
    if matter:
        sets["matter"] = matter
    for k in ("category", "note"):
        if fields.get(k):
            sets[k] = fields[k]
    if sets:
        sets["updated_at"] = now()
        conn.execute("UPDATE projects SET " + ",".join(f"{k}=?" for k in sets)
                     + " WHERE code=?", tuple(sets.values()) + (code,))
    conn.commit()
    conn.close()


def _recompute(conn, code):
    p = conn.execute("SELECT * FROM projects WHERE code=?", (code,)).fetchone()
    paid = conn.execute("SELECT IFNULL(SUM(amount),0) s FROM payments WHERE code=?",
                        (code,)).fetchone()["s"]
    gap = round((p["contract_amount"] or 0) - paid, 2)
    conn.execute("UPDATE projects SET paid_total=?, gap=?, updated_at=? WHERE code=?",
                 (round(paid, 2), gap, now(), code))


def _num(v):
    return round(float(str(v).replace(",", "").strip()), 2)


def set_amount(code, matter, stage, value, source=""):
    """stage ∈ est(申购预估)/appr(事项审批)/contract(合同)。留痕+改判分流道。"""
    if value is None or str(value).strip() == "":
        return
    value = _num(value)
    assert stage in ("est", "appr", "contract"), stage
    upsert_project(code, matter)
    conn = connect()
    col = {"est": "est_amount", "appr": "appr_amount", "contract": "contract_amount"}[stage]
    old = conn.execute(f"SELECT {col} v FROM projects WHERE code=?", (code,)).fetchone()["v"]
    if old is not None and round(float(old), 2) == value:
        conn.close()
        return
    conn.execute(f"UPDATE projects SET {col}=?, updated_at=? WHERE code=?", (value, now(), code))
    conn.execute("INSERT INTO amount_history(code,stage,old_value,new_value,source,created_at)"
                 " VALUES(?,?,?,?,?,?)", (code, stage, old, value, str(source)[:200], now()))
    if stage == "contract":
        lane = "A" if value >= SMALL_LIMIT else "B"
        conn.execute("UPDATE projects SET lane=? WHERE code=?", (lane, code))
    _recompute(conn, code)
    conn.commit()
    conn.close()


def add_payment(code, matter, amount, payee="", nature_src=(), ledger="", row=0,
                requestid="", status="待付款", note=""):
    """付款行入账。nature_src=[事由/摘要, 回单摘要, 文件名] 依次取词。"""
    if code is None or amount is None or str(amount).strip() == "":
        return None
    amount = _num(amount)
    nature = nature_of(*nature_src)
    upsert_project(code, matter)
    conn = connect()
    seq = conn.execute("SELECT COUNT(*) c FROM payments WHERE code=? AND nature=?",
                       (code, nature)).fetchone()["c"] + 1
    # 同 台账+行号 再来一次视为更新（回写requestid/状态）
    dup = conn.execute("SELECT id FROM payments WHERE code=? AND ledger=? AND row=?",
                       (code, ledger, row)).fetchone() if ledger else None
    if dup:
        conn.execute("UPDATE payments SET amount=?, payee=?, status=?, requestid=?, note=?"
                     " WHERE id=?",
                     (amount, payee, status, requestid, note, dup["id"]))
        pid = dup["id"]
    else:
        cur = conn.execute(
            "INSERT INTO payments(code,nature,seq,amount,payee,ledger,row,requestid,status,note,created_at)"
            " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (code, nature, seq, amount, payee, ledger, row, requestid, status, note, now()))
        pid = cur.lastrowid
    _recompute(conn, code)
    conn.commit()
    conn.close()
    return pid


def mark_paid(code, amount, payee="", receipt_note="", receipt_src=()):
    """M07 回单命中：该项目内金额相符的'待付款'行 → '已付款'；无则补记（Lane B）。"""
    amount = _num(amount)
    conn = connect()
    row = conn.execute("SELECT * FROM payments WHERE code=? AND status='待付款' "
                       "ORDER BY ABS(amount-?) LIMIT 1", (code, amount)).fetchone()
    if row and abs(row["amount"] - amount) <= 5.0:
        new_note = ((row["note"] or "") + f"；回单:{receipt_note}")[:200]
        conn.execute("UPDATE payments SET status='已付款', note=? WHERE id=?",
                     (new_note, row["id"]))
    else:
        nature = nature_of(*receipt_src, receipt_note)
        seq = conn.execute("SELECT COUNT(*) c FROM payments WHERE code=? AND nature=?",
                           (code, nature)).fetchone()["c"] + 1
        conn.execute("INSERT INTO payments(code,nature,seq,amount,payee,status,note,created_at)"
                     " VALUES(?,?,?,?,?,?,?,?)",
                     (code, nature, seq, amount, payee, "已付款", f"回单:{receipt_note}", now()))
    _recompute(conn, code)
    conn.commit()
    conn.close()


# ---------- 查询 / 对账 ----------

def get(code):
    conn = connect()
    p = conn.execute("SELECT * FROM projects WHERE code=?", (code,)).fetchone()
    out = dict(p) if p else None
    if out:
        out["payments"] = [dict(r) for r in conn.execute(
            "SELECT * FROM payments WHERE code=? ORDER BY nature, seq", (code,))]
    conn.close()
    return out


def flags(code):
    """异常标记：超付 / 未执行差额待说明（差额>0 且无在途待付款）。"""
    p = get(code)
    if not p:
        return []
    f = []
    if p["contract_amount"] and p["paid_total"] and p["paid_total"] > p["contract_amount"] + 0.01:
        f.append("超付")
    elif (p["contract_amount"] and p["paid_total"]
          and abs(p["gap"]) > 0.01 and not any(
              x["status"] == "待付款" for x in p["payments"])):
        f.append("未执行差额")
    return f


def report():
    conn = connect()
    ps = [dict(r) for r in conn.execute(
        "SELECT * FROM projects ORDER BY code").fetchall()]
    conn.close()
    if not ps:
        print("（金额链主档为空：合同/付款/回单跑过后自动积累）")
        return
    print(f"{'编号':10} {'事项':14} {'合同金额':>10} {'已付':>10} {'差额':>10}  性质拆分")
    for p in ps:
        kinds = " ".join(f"{x['nature']}{'' if x['status']=='已付款' else '*'}{x['amount']:g}"
                         for x in get(p["code"])["payments"])
        fl = "/".join(flags(p["code"]))
        print(f"{p['code']:10} {(p['matter'] or '')[:14]:14} "
              f"{p['contract_amount'] or 0:>10.2f} {p['paid_total'] or 0:>10.2f} "
              f"{p['gap'] if p['gap'] is not None else 0:>10.2f}  {kinds}"
              + (f"  [{fl}]" if fl else ""))
    print("（* = 待付款）")


if __name__ == "__main__":
    import sys
    utf8 = getattr(sys.stdout, "reconfigure", None)
    if utf8:
        utf8(encoding="utf-8", errors="replace")
    cmd = sys.argv[1] if len(sys.argv) > 1 else "report"
    if cmd == "report":
        report()
    elif cmd == "show" and len(sys.argv) > 2:
        print(json.dumps(get(sys.argv[2]), ensure_ascii=False, indent=2))
