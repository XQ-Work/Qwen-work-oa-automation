# -*- coding: utf-8 -*-
"""
付款回单 OCR → 字段提取 → 台账生成 + 重命名归档
================================================
输入   : 工作区「输入」目录下的付款回单 PNG（财务微信图片）
处理   : 调用本地 NPU OCR (run.ps1) 逐张识别 → 正则提取 12 个台账字段
输出   : 付款回单台账.xlsx（12 列，追加式；幂等去重依据 = 会计流水号）
归档   : 识别成功且字段完整的 PNG，重命名为
         「收款人名称_金额两位小数_摘要_记账日期.png」后移入
         已完成目录下以「收款人名称」命名的子文件夹（按供应商分类）
         重名自动加 _2/_3 后缀；字段缺失的留在「输入」并标"待人工核对"

用法   :
  python 扫描付款回单生成台账.py                 # 正式运行
  python 扫描付款回单生成台账.py --dry-run       # 只预览不写盘不移动
  python 扫描付款回单生成台账.py --input <dir> --output <xlsx> --ocr-script <ps1>
"""
from __future__ import annotations

import argparse
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from decimal import Decimal, InvalidOperation

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

# ============================================================
# 路径配置（可被命令行参数覆盖）
# ============================================================
import sys as _sys
_sys.path.insert(0, str(__import__("pathlib").Path(__file__).resolve().parents[2]))
from oa_common import paths as _paths
BASE_DIR = str(_paths.DATA / "转账回单")   # 收编进套件:指向新工作区
INPUT_DIR = os.path.join(BASE_DIR, "输入")
DONE_DIR = os.path.join(BASE_DIR, "已完成")
OUTPUT_XLSX = os.path.join(BASE_DIR, "付款回单台账.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "扫描日志.txt")
OCR_SCRIPT = r"c:\Users\JQZH\.trae-cn\skills\local-ocr-npu\scripts\run.ps1"

# ============================================================
# 日志（控制台 + 文件双路）
# ============================================================
log = logging.getLogger("pay_slip_scan")


def setup_logging(log_path: str):
    log.setLevel(logging.INFO)
    if log.handlers:
        log.handlers.clear()
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    log.addHandler(ch)


# ============================================================
# OCR 调用（复用 一键生成_优化版.py 的成熟范式）
# ============================================================
def ocr_one(img_path: str, index: int, ocr_script: str, tmp_root: str, timeout=300) -> list[str]:
    """使用 NPU OCR 识别单张图片（强制 CPU 模式）。
    ppocr.exe 不支持中文路径/文件名，先复制到纯英文临时文件。"""
    os.makedirs(tmp_root, exist_ok=True)
    tmp_img = os.path.join(tmp_root, "img_%04d.jpg" % index)
    try:
        shutil.copy2(img_path, tmp_img)
    except Exception as e:
        log.warning("复制临时文件失败 %s: %s", img_path, e)
        return []
    try:
        r = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", ocr_script, tmp_img, "-Device", "cpu"],
            capture_output=True, timeout=timeout,
        )
        raw = r.stdout
        try:
            out = raw.decode("utf-8")
        except UnicodeDecodeError:
            out = raw.decode("gbk", errors="replace")
        if r.returncode != 0:
            err = [l for l in out.splitlines() if "[ERROR]" in l]
            msg = "; ".join(err) if err else ("exit code %d" % r.returncode)
            log.warning("OCR失败 %s: %s", os.path.basename(img_path), msg)
            return []
        lines: list[str] = []
        inside = False
        for line in out.splitlines():
            s = line.strip()
            if s.startswith("---") and len(s) >= 40:
                inside = not inside
                continue
            if inside and s:
                lines.append(s)
        return lines
    except subprocess.TimeoutExpired:
        log.warning("OCR超时(>%ss) %s", timeout, os.path.basename(img_path))
        return []
    except Exception as e:
        log.warning("OCR异常 %s: %s", os.path.basename(img_path), e)
        return []
    finally:
        if os.path.exists(tmp_img):
            try:
                os.remove(tmp_img)
            except OSError:
                pass


# ============================================================
# 字段提取正则（table-driven，加字段只需在此加一行）
# 注：2026-09-09 改造 — 删除「付款人名称/开户行/币种」三列（不再入台账）；
#     保留「金额大写」用于内校验（不存台账也参与比对）
# 修：所有 \s 改为 [ \t]，避免跨行匹配污染字段
# ============================================================
FIELDS = {
    "记账日期":  r"记账日期[ \t]*[:：]?[ \t]*(\d{4})[ \t]*(\d{2})[ \t]*(\d{2})",
    "收款人名称": r"收款人名称[ \t]*[:：]?[ \t]*([^\n]+)",
    "收款人账号": r"收款人账号[ \t]*[:：]?[ \t]*(\d[\d\s]*)",
    "金额":     r"(?<!大写)金额[ \t]*[:：]?[ \t]*[¥￥]?([\d,，][\d,，\.]*)",
    "金额大写":  r"金额大写[ \t]*[:：][ \t]*([^\n]+)",
    "摘要":     r"摘要[ \t]*[:：]?[ \t]*([^\n]+)",
    "会计流水号": r"会计流水号[ \t]*[:：]?[ \t]*([A-Za-z0-9]+)",
    "批次元号":  r"批次(?:元|原)?号[ \t]*[:：]?[ \t]*(\d+)",
}

# 台账表头列顺序（9 列；去重靠第 8 列会计流水号，兜底靠内容指纹 = 收款人|金额|摘要|日期）
HEADERS = ["序号", "记账日期", "收款人名称", "收款人账号",
           "金额", "金额大写", "摘要", "会计流水号", "批次元号"]
MONEY_COL = 5    # 金额列（1 起）
CAP_COL   = 6    # 金额大写列（保留存档）
ID_COL    = 8    # 会计流水号列（去重键）

# 必须项（缺任一 → 行标"待人工核对"、文件留输入目录）；其余字段均为可选项
REQUIRED_FIELDS = ("收款人名称", "金额", "摘要")


def normalize_colon(text: str) -> str:
    return text.replace("：", ":").replace("（", "(").replace("）", ")")


def extract_fields(lines: list[str]) -> dict:
    """从 OCR 行文本提取台账字段。缺失字段填 '待人工核对'。"""
    text = normalize_colon("\n".join(lines))
    out: dict[str, object] = {}
    missing: list[str] = []
    for key, pat in FIELDS.items():
        if key == "记账日期":
            m = re.search(pat, text)
            out[key] = ("%s-%s-%s" % (m.group(1), m.group(2), m.group(3))) if m else ""
        elif key == "金额":
            m = re.search(pat, text)
            if m:
                digits = re.sub(r"[,\s，]", "", m.group(1))
                # 仅保留最多两位小数的数字（防 OCR 粘入多余点号）
                mm = re.match(r"^(\d+)(?:\.(\d{1,2}))?", digits)
                out[key] = (mm.group(1) + ("." + mm.group(2) if mm.group(2) else "")) if mm else ""
            else:
                out[key] = ""
        else:
            m = re.search(pat, text)
            val = m.group(1).strip() if m else ""
            if key == "收款人账号":
                val = re.sub(r"\s+", "", val)
            out[key] = val
    # 必须项缺失 → 标"待人工核对"；可选字段（记账日期/流水号/批次号等）缺失 → 留空
    for k in out:
        if not out[k]:
            out[k] = "待人工核对" if k in REQUIRED_FIELDS else ""
    missing = [k for k in REQUIRED_FIELDS if out[k] == "待人工核对"]
    out["__missing__"] = missing
    return out


# ============================================================
# 中文大写金额 → Decimal（用于金额交叉校验）
# 例："人民币壹万叁仟玖佰贰拾元整" → Decimal("13920.00")
#     "捌仟叁佰玖拾元伍角"        → Decimal("8390.50")
#     "壹拾元整"                  → Decimal("10.00")
# 解析失败返回 None（OCR 大写乱码场景，避免误伤）
# ============================================================
_CN_NUM = {"零":0,"壹":1,"贰":2,"叁":3,"肆":4,"伍":5,"陆":6,"柒":7,"捌":8,"玖":9,
           "〇":0,"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9}
_CN_SMALL = {"拾":10,"佰":100,"仟":1000}
_CN_BIG = {"万":10000,"亿":100000000}
_CN_YUAN_RE = re.compile(r"[元圆]")


def _parse_int_section(s: str) -> int:
    """解析 ≤ 4 位的一节(无亿/万单位)，如 '叁仟玖佰贰拾' → 3920；'壹拾' → 10。"""
    if not s: return 0
    s = s.replace("零", "")
    if not s: return 0
    val, cur = 0, 0
    for ch in s:
        if ch in _CN_NUM:
            cur = _CN_NUM[ch]
        elif ch in _CN_SMALL:
            val += (cur if cur else 1) * _CN_SMALL[ch]
            cur = 0
        else:
            # 未知字符，跳过
            pass
    return val + cur


def _parse_int_part(s: str) -> int:
    """处理 '亿/万' 切节，例 '壹万叁仟玖佰贰拾' → 13920。"""
    if not s: return 0
    total = 0
    if "亿" in s:
        yi, rest = s.split("亿", 1)
        total += _parse_int_section(yi) * 100000000
        s = rest
    if "万" in s:
        wan, rest = s.split("万", 1)
        total += _parse_int_section(wan) * 10000
        s = rest
    total += _parse_int_section(s)
    return total


def cn_upper_to_decimal(text: str):
    """中文大写金额 → Decimal(两位小数)；解析失败返回 None。"""
    if not text: return None
    s = str(text).strip()
    s = re.sub(r"(人民币|RMB|￥|¥)", "", s)
    s = re.sub(r"[\s()（）]", "", s)
    if not s: return None
    s = re.sub(r"[整正]$", "", s)   # 尾缀「整/正」去除
    if not s: return None
    # 拆元前后
    m = _CN_YUAN_RE.search(s)
    if m:
        int_part = s[:m.start()]
        tail = s[m.end():]
    else:
        int_part, tail = s, ""
    # 关键校验：int_part 中必须含至少一个有效数字字符，否则视为 OCR 乱码返回 None
    if not any(ch in _CN_NUM for ch in int_part):
        return None
    try:
        yuan = _parse_int_part(int_part)
    except Exception:
        return None
    frac = 0
    m_jiao = re.search(r"([零〇壹贰叁肆伍陆柒捌玖一二三四五六七八九])角", tail)
    if m_jiao: frac += _CN_NUM[m_jiao.group(1)] * 10
    m_fen = re.search(r"([零〇壹贰叁肆伍陆柒捌玖一二三四五六七八九])分", tail)
    if m_fen: frac += _CN_NUM[m_fen.group(1)]
    return Decimal(yuan) + Decimal(frac) / Decimal(100)


def content_fingerprint(fields: dict):
    """单据内容指纹 = 收款人|金额|摘要|日期(可空)，四元组相同视为重复。"""
    payee = str(fields.get("收款人名称", "")).strip()
    amt = str(fields.get("金额", "")).strip()
    memo = str(fields.get("摘要", "")).strip()
    date = str(fields.get("记账日期", "")).strip()
    if any(x in ("", "待人工核对") for x in (payee, amt, memo)):
        return None
    return "%s|%s|%s|%s" % (payee, amt, memo, date)


# ============================================================
# 归档重命名：收款人名称_金额(两位小数)_摘要_记账日期
# ============================================================
ILLEGAL_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')


def sanitize(s: str) -> str:
    s = ILLEGAL_CHARS.sub("_", str(s))
    s = re.sub(r"\s+", "", s)
    return s[:40] or "待人工核对"


def build_new_name(fields: dict) -> str:
    """收款人_金额两位小数_摘要_记账日期.png；记账日期(可选)缺失时省略日期段。"""
    payee = sanitize(fields["收款人名称"])
    memo = sanitize(fields["摘要"])
    amt_raw = str(fields["金额"]).replace("待人工核对", "0") or "0"
    try:
        amt = "%.2f" % Decimal(amt_raw)   # 两位小数：13920.00 / 3927.17
    except (InvalidOperation, ValueError):
        amt = "0.00"
    parts = [payee, amt, memo]
    date = sanitize(fields["记账日期"])
    if date and date != "待人工核对":
        parts.append(date)
    return "_".join(parts) + ".png"


def unique_dest(done_dir: str, new_name: str) -> str:
    """重名自动加 _2/_3 后缀，返回可用目标路径。"""
    stem, ext = os.path.splitext(new_name)
    cand = os.path.join(done_dir, new_name)
    n = 2
    while os.path.exists(cand):
        cand = os.path.join(done_dir, "%s_%d%s" % (stem, n, ext))
        n += 1
    return cand


# ============================================================
# Excel 台账：表头 / 追加 / 去重 / 合计
# ============================================================
def ensure_headers(xlsx_path: str):
    wb = Workbook() if not os.path.exists(xlsx_path) else load_workbook(xlsx_path)
    ws = wb.worksheets[0]   # 总表固定为第一个工作表（汇总/供应商分表在其后）
    if ws.max_row == 0 or ws.cell(1, 1).value != HEADERS[0]:
        # 已存在旧合计行(空行起始)时先清空重建
        for c in range(1, len(HEADERS) + 1):
            cell = ws.cell(1, c)
            cell.value = HEADERS[c - 1]
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22
        # 列宽（9 列：序号/记账日期/收款人名称/收款人账号/金额/金额大写/摘要/会计流水号/批次元号）
        widths = [6, 12, 32, 22, 14, 34, 34, 22, 22]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w
    wb.save(xlsx_path)


def load_seen_ids(xlsx_path: str) -> tuple[set[str], set[str]]:
    """读取台账已有 (会计流水号, 内容指纹) 集合（用于幂等去重）。
    指纹 = 收款人|金额|摘要|日期  四元组全等视为同一单据。"""
    if not os.path.exists(xlsx_path):
        return set(), set()
    try:
        wb = load_workbook(xlsx_path)
        ws = wb.worksheets[0]   # 总表固定为第一个工作表（汇总/供应商分表在其后）
        seen_ids: set[str] = set()
        seen_fps: set[str] = set()
        for r in range(2, ws.max_row + 1):
            v1 = ws.cell(r, 1).value
            if not isinstance(v1, (int, float)):    # 跳过表头/合计行
                continue
            # 1) 会计流水号（主去重键）
            sid = ws.cell(r, ID_COL).value
            if sid and str(sid).strip() not in ("", "待人工核对"):
                seen_ids.add(str(sid).strip())
            # 2) 内容指纹（兜底：流水号缺失时仍能识别重发）
            payee = str(ws.cell(r, 3).value or "").strip()
            amt = str(ws.cell(r, 5).value or "").strip()
            memo = str(ws.cell(r, 7).value or "").strip()
            date = str(ws.cell(r, 2).value or "").strip()
            if payee and amt and memo and amt not in ("", "待人工核对"):
                seen_fps.add("%s|%s|%s|%s" % (payee, amt, memo, date))
        return seen_ids, seen_fps
    except Exception as e:
        log.warning("读取台账去重失败: %s", e)
        return set(), set()


def append_and_total(xlsx_path: str, row_data: list):
    """单次加载：删除旧合计 → 追加新数据行 → 重建合计行（Decimal 求和）。

    避免"追加与合计分两次读写"导致的孤儿行/行错位问题。"""
    wb = load_workbook(xlsx_path)
    ws = wb.worksheets[0]   # 总表固定为第一个工作表（汇总/供应商分表在其后）

    # 1) 删除已有合计行（可能位于任意位置，一律清掉）
    del_rows = [r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == "合计"]
    for r in reversed(del_rows):
        ws.delete_rows(r)

    # 2) 真实数据末行（按第 1 列序号为数字判断）
    last_data = 1
    for r in range(2, ws.max_row + 1):
        v1 = ws.cell(r, 1).value
        if v1 is not None and isinstance(v1, (int, float)):
            last_data = r

    # 3) 追加新数据行并写序号（先重排已有序号保证连续）
    seq = 1
    for r in range(2, ws.max_row + 1):
        v1 = ws.cell(r, 1).value
        if v1 is not None and isinstance(v1, (int, float)):
            ws.cell(r, 1).value = seq
            seq += 1
    row_data[0] = seq
    ws.append(row_data)
    new_r = ws.max_row
    ws.cell(new_r, MONEY_COL).number_format = "#,##0.00"

    # 4) 重建合计行
    total = Decimal(0)
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == "合计":
            continue
        v = ws.cell(r, MONEY_COL).value
        if v is None or str(v).strip() in ("", "待人工核对"):
            continue
        try:
            total += Decimal(str(v))
        except (InvalidOperation, ValueError):
            continue
    total_row = ws.max_row + 1
    ws.cell(total_row, 1).value = "合计"
    ws.cell(total_row, MONEY_COL).value = float(total)
    ws.cell(total_row, MONEY_COL).number_format = "#,##0.00"
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(total_row, c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FCE4D6")
    wb.save(xlsx_path)


# ============================================================
# 汇总表 + 按供应商明细表（每次运行后自动重建；总表内容不动）
# ============================================================
MAIN_SHEET_NAME = "付款回单台账"      # 总表标签名
SUMMARY_SHEET = "汇总"               # 汇总表标签名
_HDR_FILL = PatternFill("solid", fgColor="DDEBF7")
_TOT_FILL = PatternFill("solid", fgColor="FCE4D6")
_TITLE_FONT = Font(bold=True, size=13, color="1F4E79")
_SUB_FONT = Font(size=9, color="808080")
_CENTER = Alignment(horizontal="center", vertical="center")
_SHEET_ILLEGAL = re.compile(r"[\[\]:*?/\\]")   # Excel 工作表名禁用字符


def fmt_money(d) -> str:
    """Decimal/数字 → 千分位两位小数字符串。"""
    try:
        return "{:,.2f}".format(Decimal(str(d)))
    except (InvalidOperation, ValueError):
        return "0.00"


def _read_main_rows(ws) -> list[dict]:
    """读取总表全部数据行（跳过表头与合计行），返回结构化列表。"""
    rows: list[dict] = []
    for r in range(2, ws.max_row + 1):
        if not isinstance(ws.cell(r, 1).value, (int, float)):
            continue
        raw = ws.cell(r, MONEY_COL).value
        txt = "" if raw is None else str(raw).strip()
        try:
            amt = Decimal(txt) if txt not in ("", "None", "待人工核对") else Decimal(0)
        except (InvalidOperation, ValueError):
            amt = Decimal(0)
        rows.append({
            "seq": ws.cell(r, 1).value,
            "date": ws.cell(r, 2).value or "",
            "payee": str(ws.cell(r, 3).value or "").strip() or "未识别收款人",
            "amt": amt,
            "amt_txt": txt,
            "memo": ws.cell(r, 7).value or "",
            "sid": ws.cell(r, 8).value or "",
            "batch": ws.cell(r, 9).value or "",
        })
    return rows


def _sheet_title(payee: str, used: set) -> str:
    """收款人名 → 合法且不重名的工作表名（限 31 字符，禁用 [ ] : * ? / \\）。"""
    s = _SHEET_ILLEGAL.sub("_", str(payee)).strip().strip("'")
    s = (s or "未命名")[:31]
    base, n = s, 2
    while s.lower() in used:
        suffix = "(%d)" % n
        s = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(s.lower())
    return s


def _style_row(ws, row: int, ncol: int, fill):
    for i in range(1, ncol + 1):
        c = ws.cell(row, i)
        c.font = Font(bold=True)
        c.fill = fill
        c.alignment = _CENTER


def _put_title(ws, ncol: int, title: str, subtitle: str):
    """写两行表头（大标题 + 灰色副标题）。"""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    t = ws.cell(1, 1, title)
    t.font = _TITLE_FONT
    t.alignment = _CENTER
    ws.row_dimensions[1].height = 26
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    s = ws.cell(2, 1, subtitle)
    s.font = _SUB_FONT
    s.alignment = Alignment(horizontal="center")


def rebuild_summary(xlsx_path: str) -> dict:
    """重建「汇总」表与每个供应商的独立明细工作表。

    总表（第一个工作表）仅作数据源，内容与列结构均不改动；
    除总表外的工作表一律视为本流程生成，重建时先清空再写入，
    因此供应商增减、改名后分表始终与总表一致。
    """
    if not os.path.exists(xlsx_path):
        return {}
    wb = load_workbook(xlsx_path)
    main = wb.worksheets[0]
    if main.title != MAIN_SHEET_NAME:
        main.title = MAIN_SHEET_NAME        # 仅改标签名，表格内容不动
    rows = _read_main_rows(main)

    # 1) 清掉上次生成的表（总表保留）
    for name in list(wb.sheetnames):
        if name != main.title:
            del wb[name]

    # 2) 按收款人聚合
    agg: dict[str, dict] = {}
    for x in rows:
        a = agg.setdefault(x["payee"], {"cnt": 0, "amt": Decimal(0), "rows": []})
        a["cnt"] += 1
        a["amt"] += x["amt"]
        a["rows"].append(x)
    order = sorted(agg.items(), key=lambda kv: (-kv[1]["amt"], kv[0]))
    total_amt = sum((a["amt"] for a in agg.values()), Decimal(0))
    total_cnt = len(rows)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 3) 「汇总」表（紧跟总表）
    ws = wb.create_sheet(SUMMARY_SHEET, 1)
    _put_title(ws, 5, "付款回单汇总（按收款人）",
               "数据来源：总表（%s）　|　共 %d 家收款人 / %d 笔　|　合计 ¥%s　|　统计时间 %s"
               % (main.title, len(order), total_cnt, fmt_money(total_amt), now))
    heads = ["序号", "收款人名称", "笔数", "金额合计", "占比"]
    for i, h in enumerate(heads, start=1):
        ws.cell(3, i, h)
    _style_row(ws, 3, len(heads), _HDR_FILL)
    r = 4
    for i, (payee, a) in enumerate(order, start=1):
        ws.cell(r, 1, i).alignment = _CENTER
        ws.cell(r, 2, payee)
        ws.cell(r, 3, a["cnt"]).alignment = _CENTER
        ws.cell(r, 4, float(a["amt"])).number_format = "#,##0.00"
        ws.cell(r, 5, float(a["amt"] / total_amt) if total_amt else 0.0).number_format = "0.00%"
        r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.cell(r, 1, "合计")
    ws.cell(r, 3, total_cnt)
    ws.cell(r, 4, float(total_amt)).number_format = "#,##0.00"
    ws.cell(r, 5, 1.0 if total_amt else 0.0).number_format = "0.00%"
    _style_row(ws, r, len(heads), _TOT_FILL)
    for i, w in enumerate([6, 34, 8, 16, 10], start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A4"

    # 4) 每家供应商一张明细表
    used = {SUMMARY_SHEET.lower(), main.title.lower()}
    for payee, a in order:
        sws = wb.create_sheet(_sheet_title(payee, used))
        _put_title(sws, 6, payee,
                   "共 %d 笔　|　合计 ¥%s　|　数据来源：总表（%s）"
                   % (a["cnt"], fmt_money(a["amt"]), main.title))
        headers = ["序号", "记账日期", "金额", "摘要", "会计流水号", "批次元号"]
        for i, h in enumerate(headers, start=1):
            sws.cell(3, i, h)
        _style_row(sws, 3, len(headers), _HDR_FILL)
        rr = 4
        for i, x in enumerate(a["rows"], start=1):
            sws.cell(rr, 1, i).alignment = _CENTER
            sws.cell(rr, 2, x["date"]).alignment = _CENTER
            sws.cell(rr, 3, float(x["amt"])).number_format = "#,##0.00"
            sws.cell(rr, 4, x["memo"])
            sws.cell(rr, 5, x["sid"])
            sws.cell(rr, 6, x["batch"])
            rr += 1
        sws.cell(rr, 1, "小计")
        sws.cell(rr, 3, float(a["amt"])).number_format = "#,##0.00"
        sws.cell(rr, 4, "%d 笔" % a["cnt"])
        _style_row(sws, rr, len(headers), _TOT_FILL)
        for i, w in enumerate([6, 12, 14, 34, 22, 22], start=1):
            sws.column_dimensions[chr(64 + i)].width = w
        sws.freeze_panes = "A4"

    wb.active = 0       # 打开文件默认停在总表
    wb.save(xlsx_path)
    return {"suppliers": len(order), "rows": total_cnt, "total": fmt_money(total_amt)}


def move_dup(src: str, fname: str, dup_dir: str) -> str:
    """把重复单据移出输入目录（保留原图，不删除）。返回目标路径，失败返回 ''。"""
    try:
        os.makedirs(dup_dir, exist_ok=True)
        dest = unique_dest(dup_dir, fname)
        shutil.move(src, dest)
        return dest
    except Exception as e:
        log.warning("  → 重复单据移出失败: %s", e)
        return ""


# ============================================================
# 主流程
# ============================================================
def main():
    ap = argparse.ArgumentParser(description="付款回单 PNG OCR → 台账 + 重命名归档")
    ap.add_argument("--input", default=INPUT_DIR, help="输入目录（默认: 转账回单/输入）")
    ap.add_argument("--output", default=OUTPUT_XLSX, help="台账 xlsx 路径")
    ap.add_argument("--done", default=DONE_DIR, help="已完成归档目录")
    ap.add_argument("--ocr-script", default=OCR_SCRIPT, help="本地 OCR run.ps1 路径")
    ap.add_argument("--dry-run", action="store_true", help="只预览，不写台账、不移动文件")
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 张（测试用，0=全部）")
    ap.add_argument("--no-move-dup", action="store_true",
                    help="重复单据保留在输入目录（默认移入 待确认\\<日期>\\重复单据）")
    ap.add_argument("--rebuild-only", action="store_true",
                    help="只重建「汇总」表与供应商明细表，不做 OCR")
    args = ap.parse_args()

    input_dir = os.path.abspath(args.input)
    done_dir = os.path.abspath(args.done)
    xlsx_path = os.path.abspath(args.output)
    log_path = os.path.join(os.path.dirname(xlsx_path), "扫描日志.txt")
    dup_dir = os.path.join(os.path.dirname(xlsx_path), "待确认",
                           datetime.now().strftime("%Y-%m-%d"), "重复单据")

    setup_logging(log_path)

    # 只重建汇总/供应商明细表（不做 OCR，便于随时刷新）
    if args.rebuild_only:
        if not os.path.exists(xlsx_path):
            log.error("台账不存在: %s", xlsx_path)
            sys.exit(1)
        info = rebuild_summary(xlsx_path)
        log.info("已重建「汇总」表与供应商明细表：%d 家收款人 | %d 笔 | 合计 ¥%s",
                 info.get("suppliers", 0), info.get("rows", 0), info.get("total", "0.00"))
        return

    if args.dry_run:
        log.info("=" * 60)
        log.info("DRY-RUN 模式：不写台账、不移动文件，仅预览解析结果")
        log.info("=" * 60)

    if not os.path.isdir(input_dir):
        log.error("输入目录不存在: %s", input_dir)
        sys.exit(1)
    if not os.path.exists(args.ocr_script):
        log.error("OCR 脚本不存在: %s", args.ocr_script)
        sys.exit(1)
    os.makedirs(done_dir, exist_ok=True)

    png_files = sorted(f for f in os.listdir(input_dir)
                       if f.lower().endswith((".png", ".jpg", ".jpeg")))
    if args.limit and args.limit > 0:
        png_files = png_files[:args.limit]
    if not png_files:
        log.info("输入目录没有待处理的图片: %s", input_dir)
        if not args.dry_run and os.path.exists(xlsx_path):
            info = rebuild_summary(xlsx_path)
            log.info("已刷新「汇总」表与供应商明细表：%d 家收款人 | %d 笔 | 合计 ¥%s",
                     info.get("suppliers", 0), info.get("rows", 0), info.get("total", "0.00"))
        return

    seen_ids, seen_fps = (set(), set()) if args.dry_run else load_seen_ids(xlsx_path)
    if not args.dry_run and os.path.exists(xlsx_path):
        ensure_headers(xlsx_path)

    log.info("待处理 %d 张图片（台账已有 %d 个流水号 / %d 个内容指纹）",
             len(png_files), len(seen_ids), len(seen_fps))
    tmp_root = os.path.join(tempfile.gettempdir(), "pay_slip_ocr_tmp")
    stats = {"ok": 0, "missing": 0, "fail": 0, "skip": 0, "dup": 0}

    for idx, fname in enumerate(png_files):
        src = os.path.join(input_dir, fname)
        log.info("OCR(%d/%d): %s", idx + 1, len(png_files), fname)
        lines = ocr_one(src, idx, args.ocr_script, tmp_root)
        if not lines:
            log.warning("  → OCR 未返回文本，文件留在输入目录待人工处理")
            stats["fail"] += 1
            continue
        fields = extract_fields(lines)
        missing = fields["__missing__"]
        sid = str(fields["会计流水号"]).strip()
        fingerprint = content_fingerprint(fields)
        # 幂等去重：① 会计流水号 命中 ② 内容指纹 命中 → 跳过
        dup_reason = ""
        if sid not in ("待人工核对", "") and sid in seen_ids:
            dup_reason = "会计流水号已在台账[%s]" % sid
        elif fingerprint and fingerprint in seen_fps:
            dup_reason = "内容指纹重复[%s]" % fingerprint
        if dup_reason:
            log.info("跳过(重复): %s | %s", fname, dup_reason)
            stats["skip"] += 1
            # 移出输入目录，避免每次重跑都对同一张重复图重复 OCR
            if not args.dry_run and not args.no_move_dup:
                moved = move_dup(src, fname, dup_dir)
                if moved:
                    log.info("  ↳ 已移出输入目录: 待确认\\%s\\重复单据\\%s",
                             os.path.basename(os.path.dirname(dup_dir)), os.path.basename(moved))
                    stats["dup"] += 1
            continue
        # 大写金额交叉校验：金额大写存在且与小写不一致 → 视为缺必须字段拦截
        upper_txt = str(fields.get("金额大写", "")).strip()
        if upper_txt and upper_txt != "待人工核对":
            cn_val = cn_upper_to_decimal(upper_txt)
            if cn_val is not None:
                try:
                    amt_val = Decimal(str(fields["金额"]))
                except (InvalidOperation, ValueError):
                    amt_val = None
                if amt_val is not None and cn_val != amt_val:
                    log.warning("  ✗ 金额大小写不一致: 小写=%s, 大写='%s' → 解析=%s, 拦截", 
                                fields["金额"], upper_txt, cn_val)
                    fields["金额"] = "待人工核对"
                    if "金额" not in missing:
                        missing.append("金额")
        new_name = build_new_name(fields)

        if args.dry_run:
            ok_mark = "✓完整" if not missing else "✗缺%s" % ",".join(missing)
            log.info("  [%s] %s", ok_mark, fname)
            log.info("        收款人=%s | 金额=%s | 摘要=%s | 日期=%s | 流水号=%s",
                     fields["收款人名称"], fields["金额"], fields["摘要"], fields["记账日期"], sid)
            log.info("        新文件名: %s", new_name)
            if missing:
                stats["missing"] += 1
            else:
                stats["ok"] += 1
            continue

        # ===== 正式写入 =====
        ensure_headers(xlsx_path)
        c = lambda v: None if v == "" else v   # 可选字段缺失 → 空单元格
        row_data = [0,
                    c(fields["记账日期"]), fields["收款人名称"], c(fields["收款人账号"]),
                    fields["金额"], c(fields["金额大写"]), fields["摘要"],
                    c(fields["会计流水号"]), c(fields["批次元号"])]
        append_and_total(xlsx_path, row_data)
        # 把本单加入去重集合，防同批内重发
        if sid not in ("待人工核对", ""):
            seen_ids.add(sid)
        if fingerprint:
            seen_fps.add(fingerprint)

        if missing:
            log.warning("  → 已写入台账，但缺必须字段(%s) → 文件留输入目录待人工核对",
                        ",".join(missing))
            stats["missing"] += 1
        else:
            # 归档：已完成\<收款人>\ 子目录（按供应商分类）
            payee_dir = sanitize(fields["收款人名称"])
            dest_dir = os.path.join(done_dir, payee_dir)
            os.makedirs(dest_dir, exist_ok=True)
            dest = unique_dest(dest_dir, new_name)
            try:
                shutil.move(src, dest)
                log.info("  ✓ 归档: %s → %s\\%s", fname, payee_dir, os.path.basename(dest))
                stats["ok"] += 1
            except Exception as e:
                log.warning("  → 台账已写入但移动失败: %s", e)
                stats["missing"] += 1

    # 无论本次是否新增，运行结束都重建汇总/供应商分表
    if not args.dry_run and os.path.exists(xlsx_path):
        info = rebuild_summary(xlsx_path)
        log.info("「汇总」表与供应商明细表已刷新：%d 家收款人 | %d 笔 | 合计 ¥%s",
                 info.get("suppliers", 0), info.get("rows", 0), info.get("total", "0.00"))

    log.info("=" * 60)
    log.info("处理完成: 完整归档=%d | 待人工核对=%d | OCR失败=%d | 重复跳过=%d(已移出%d)",
             stats["ok"], stats["missing"], stats["fail"], stats["skip"], stats["dup"])
    if args.dry_run:
        log.info("(DRY-RUN 未写入任何文件)")
    else:
        log.info("台账: %s", xlsx_path)
        log.info("归档目录: %s", done_dir)
        log.info("日志: %s", log_path)


if __name__ == "__main__":
    main()
