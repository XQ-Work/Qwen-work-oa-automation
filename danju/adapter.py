# -*- coding: utf-8 -*-
"""单据类适配器（并入 oa-automation 主包后的统一入口）。

原则"合而不重构"：不改 一键生成.py 等内部逻辑，本文件只做两件事——
  1) 转调单据类自带 CLI（run.py 会自动挑选装好依赖的解释器）；
  2) 直读 60_数据/单据.db（sqlite 只读），供 OA 侧统一归档/付款预填取数。

用法：
  python adapter.py status            概况（入库单/发票/输入目录待处理数）
  python adapter.py status --json     同上，JSON 输出（供其他模块调用）
  python adapter.py receipts [--all]  未生成付款单[--全部]的入库单清单
  python adapter.py invoices          未处理/用途为空的发票清单
  python adapter.py run [--report]    执行单据类一键生成（[--报告链路]）

数据根解析与 danju/paths.py 同序：环境变量 DJ_BASE > 主套件
config.json 的 danju_base（相对 data_root）> 本目录。
"""
import argparse
import json
import os
import sqlite3
import subprocess
import sys

# 强制 UTF-8 输出（git-bash/cmd 皆不乱码）：Python 流按 UTF-8 写，
# Windows 下同时把控制台输出代码页切到 65001，退出时不残留（CP 变化无碍）。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
if os.name == "nt":
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE_CONFIG = os.path.normpath(os.path.join(HERE, os.pardir, "config.json"))


def base_dir():
    b = os.environ.get("DJ_BASE")
    if b:
        return b
    try:
        with open(SUITE_CONFIG, encoding="utf-8") as f:
            cfg = json.load(f)
        db_ = cfg.get("danju_base")
        if db_:
            return db_ if os.path.isabs(db_) else os.path.join(cfg.get("data_root", ""), db_)
    except Exception:
        pass
    return HERE


def db_path():
    return os.path.join(base_dir(), "60_数据", "单据.db")


def _rows(sql, args=()):
    path = db_path()
    if not os.path.exists(path):
        return []
    try:
        conn = sqlite3.connect("file:%s?mode=ro" % path.replace("\\", "/"), uri=True)
        conn.row_factory = sqlite3.Row
        return [dict(r) for r in conn.execute(sql, args).fetchall()]
    except sqlite3.OperationalError:
        return []          # 空库/未初始化：无表即视为零记录
    finally:
        try:
            conn.close()
        except Exception:
            pass


def _inout_counts():
    """10_输入 各分拣目录中待处理文件数（不算子目录里的归档件）。"""
    root = base_dir()
    out = {}
    for sub in ["", "入库单扫描件", "发票", "申购单", "事项审批"]:
        d = os.path.join(root, "10_输入", sub)
        out[sub or "根目录"] = (len([f for f in os.listdir(d)
                                    if os.path.isfile(os.path.join(d, f))])
                               if os.path.isdir(d) else 0)
    return out


def cmd_status(as_json=False):
    rec = _rows("SELECT status, COUNT(*) n FROM receipts GROUP BY status")
    inv = _rows("SELECT COUNT(*) total, "
                "SUM(processed) done, "
                "SUM(CASE WHEN IFNULL(usage,'')='' THEN 1 ELSE 0 END) no_usage "
                "FROM invoices")
    s = {
        "base": base_dir(),
        "receipts": {r["status"] or "matched": r["n"] for r in rec},
        "invoices": inv[0] if inv else {"total": 0, "done": 0, "no_usage": 0},
        "inbox": _inout_counts(),
    }
    if as_json:
        print(json.dumps(s, ensure_ascii=False, indent=2))
        return s
    print("数据根: %s" % s["base"])
    print("入库单: %s" % (s["receipts"] or "（空）"))
    iv = s["invoices"]
    print("发票: 共%s张，已出付款单%s张，用途为空%s张"
          % (iv["total"], iv["done"] or 0, iv["no_usage"] or 0))
    print("10_输入待处理: %s" % s["inbox"])
    return s


def cmd_receipts(all=False):
    if all:
        rows = _rows("SELECT receipt_id,date,supplier,total,status FROM receipts "
                     "ORDER BY date, receipt_id")
    else:
        rows = _rows("SELECT receipt_id,date,supplier,total,status FROM receipts "
                     "WHERE IFNULL(status,'matched')!='paid' ORDER BY date, receipt_id")
    for r in rows:
        print("%s  %s  %-24s  %10.2f  [%s]" % (r["receipt_id"], r["date"] or "",
                                               r["supplier"] or "", r["total"] or 0,
                                               r["status"] or "matched"))
    print("共 %d 张%s" % (len(rows), "" if all else "（未出付款单）"))


def cmd_invoices():
    rows = _rows("SELECT filename,seller,amount,usage,processed,created_at "
                 "FROM invoices ORDER BY id")
    for r in rows:
        flag = "已处理" if r["processed"] else "未处理"
        print("%s  %-12s %10.2f  用途=%-10s [%s]  %s"
              % (r["created_at"] or "", r["seller"] or "", r["amount"] or 0,
                 r["usage"] or "（空）", flag, r["filename"]))
    print("共 %d 张" % len(rows))


def cmd_run(report=False):
    cmd = [sys.executable, os.path.join(HERE, "run.py")]
    if report:
        cmd.append("report")
    print("$ %s" % " ".join(cmd))
    return subprocess.call(cmd, cwd=HERE)


def main():
    ap = argparse.ArgumentParser(description="单据类适配器")
    ap.add_argument("cmd", choices=["status", "receipts", "invoices", "run"])
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--report", action="store_true")
    a = ap.parse_args()
    if a.cmd == "status":
        cmd_status(as_json=a.json)
    elif a.cmd == "receipts":
        cmd_receipts(all=a.all)
    elif a.cmd == "invoices":
        cmd_invoices()
    elif a.cmd == "run":
        sys.exit(cmd_run(report=a.report))


if __name__ == "__main__":
    main()
