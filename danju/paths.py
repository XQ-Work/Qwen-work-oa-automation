# -*- coding: utf-8 -*-
"""统一路径常量（单据类自动化 v4 重构版）。

说明：本模块为"开发沙盒 + 运行时指向单据类"模式。所有路径集中在此，
脚本内不再散写绝对路径。目标采用 10 档分层结构。
"""
import os

# ============ 根目录 ============
# 数据根解析优先级：环境变量 DJ_BASE > 主套件 config.json 的 danju_base
# （相对路径按 data_root 解析）> 本文件所在目录（默认，整体搬家无需改代码）。
# 2026-09 并入 oa-automation 主包：内部逻辑不变，仅新增套件级挂钩。
BASE = os.environ.get("DJ_BASE")
if not BASE:
    try:
        import json
        _suite_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  os.pardir, "config.json")
        if os.path.exists(_suite_cfg):
            with open(_suite_cfg, encoding="utf-8") as _f:
                _cfg = json.load(_f)
            _b = _cfg.get("danju_base")
            if _b:
                BASE = _b if os.path.isabs(_b) else os.path.join(_cfg.get("data_root", ""), _b)
            else:
                BASE = None
    except Exception:
        BASE = None
BASE = BASE or os.path.dirname(os.path.abspath(__file__))

# ============ 10_输入（待处理单据直接丢根目录，程序自动分拣归位）============
INPUT_DIR     = os.path.join(BASE, "10_输入")
SCAN_DIR      = os.path.join(BASE, "10_输入", "入库单扫描件")
INVOICE_IN_DIR = os.path.join(BASE, "10_输入", "发票")
PURCHASE_IN_DIR = os.path.join(BASE, "10_输入", "申购单")
APPROVAL_IN_DIR = os.path.join(BASE, "10_输入", "事项审批")

# ============ 20_输出（脚本生成成品，待打印）============
OUTPUT_DIR    = os.path.join(BASE, "20_输出")
LEDGER_DIR    = os.path.join(OUTPUT_DIR, "台账")
ACCEPT_DIR    = os.path.join(OUTPUT_DIR, "验收单")   # docx 与 pdf 成品
PAY_DIR       = os.path.join(OUTPUT_DIR, "付款单")   # xlsx 与 pdf 成品
INVOICE_OUT_DIR = os.path.join(OUTPUT_DIR, "发票")   # 发票归档副本

# ============ 30_待打印 / 40_归档 ============
PRINT_DIR    = os.path.join(BASE, "30_待打印")
ARCHIVE_DIR  = os.path.join(BASE, "40_归档")

# ============ 50_模板（只读）============
TEMPLATE_DIR   = os.path.join(BASE, "50_模板")
TEMPLATE_XLS   = os.path.join(TEMPLATE_DIR, "采购物资台帐.xls")
TEMPLATE_DOCX  = os.path.join(TEMPLATE_DIR, "验收单模板.docx")
TEMPLATE_PAY_XLSX = os.path.join(TEMPLATE_DIR, "付款单模板.xlsx")
DATABASE_XLSX  = os.path.join(TEMPLATE_DIR, "付款单数据库.xlsx")

# ============ 60_数据（数据库）============
DATA_DIR = os.path.join(BASE, "60_数据")
DB_PATH  = os.path.join(DATA_DIR, "单据.db")

# ============ 70_日志 ============
LOG_DIR = os.path.join(BASE, "70_日志")

# ============ OCR 等外部依赖 ============
WORK_DIR   = os.path.join(BASE, "90_临时")
# 识别结果缓存：每张扫描件对应 "<图片MD5>.txt"，人工校对后的行式明细优先于自动识别
OCR_CACHE_DIR = os.path.join(WORK_DIR, "ocr_cache")
# 本地 OCR 脚本（Windows 内置 OCR 引擎，zh-Hans-CN）
OCR_SCRIPT = os.path.join(BASE, "80_工具", "ocr_local.ps1")


def ensure_dirs():
    """创建所有需要存在的目录（幂等）。"""
    for d in [SCAN_DIR, INVOICE_IN_DIR, PURCHASE_IN_DIR, APPROVAL_IN_DIR,
              LEDGER_DIR, ACCEPT_DIR, PAY_DIR, INVOICE_OUT_DIR,
              PRINT_DIR, ARCHIVE_DIR, TEMPLATE_DIR, DATA_DIR, LOG_DIR]:
        os.makedirs(d, exist_ok=True)


def ledger_path(year=2026):
    """采购台账（按年）。"""
    return os.path.join(LEDGER_DIR, "采购台账_%d.xls" % year)


def payment_ledger_path(year=2026):
    """付款台账（按年）。"""
    return os.path.join(LEDGER_DIR, "付款台账_%d.xls" % year)