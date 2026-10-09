# -*- coding: utf-8 -*-
"""采购项目档案 项目夹定位 + 双轨归档（查阅轨定稿册组装 / 台账快照）。

双轨：
  · 任务轨 = paths.ARCHIVE（记录/项目档案）—— 过程全量，办结后原地不动，供机器重跑/对账。
  · 查阅轨 = paths.ARCHIVED_SMALL / ARCHIVED_LARGE（归档/日常采购·采购合同）——
             从交付区(paths.DELIVER)按“功能”组装一份定稿册，一项目一夹、①②③业务顺序、
             夹名带〔办结日期〕；同时把整套台账快照刷进 归档/台账快照。
归档=组装副本(零删除)，取出=把定稿册移入 运行/state/归档取出回收，记录轨始终不受影响。
"""
import datetime
import re
import shutil

from oa_common import paths


def norm(s):
    return re.sub(r"[（）()【】\[\]\s\-_·、,，:：]", "", s or "")


# 交付区“功能” -> 查阅册里的浏览名（顺序=人看档案的业务直觉顺序，非流程编号）。
# 合同口径：档案放“用印合同”（已盖章生效版），未盖章原件属过程件留在任务轨。
BROWSE_ORDER = [
    ("事项审批", "事项审批"),
    ("比质比价报告单", "比质比价报告"),
    ("报价单", "报价单"),
    ("__合同__", "采购合同"),          # 特殊：优先用印合同，回退未盖章采购合同
    ("流程会签", "流程会签"),
    ("发票", "发票"),
    ("__回单__", "银行回单"),      # 真实付款凭证，从回单区取
    ("其他", "其他"),
]
# 生成物料（送签/做账表单，非真实凭证）：单列一组，进册末尾「生成物料」子夹并标注
GEN_ITEMS = [
    ("验收单", "验收单"),
    ("付款单", "付款单"),
]
GEN_SUBDIR = "生成物料（非真实凭证）"


def _ord(i):
    return f"{i + 1:02d}"       # 册内序号：01 起两位数，不用符号


def _ledger_dir():
    return paths.DATA / "记录/台账"


def _code_matter(folder_name):
    """从 '2026-009-高效沉淀池遮阳板' 或 '2026-009-高效沉淀池遮阳板〔办结…〕' 解析 code + matter。"""
    m = re.match(r"^(\d{4}-\d{3})-(.*?)(?:〔.*)?$", folder_name)
    if m:
        return m.group(1), m.group(2).strip()
    return None, folder_name


def _latest_deliver(func, matter):
    """交付/func 里匹配该事项、日期前缀最新的那份 PDF。"""
    if func == "__合同__":
        return _latest_deliver("用印合同", matter) or _latest_deliver("采购合同", matter)
    d = paths.DELIVER / func
    if not d.is_dir():
        return None
    return _pick_latest(d.glob("*.pdf"), matter)


def _pick_latest(files, matter):
    nm = norm(matter)
    cands = [(re.match(r"^(\d{8})", f.name).group(1) if re.match(r"^(\d{8})", f.name) else "00000000", f)
             for f in files if (nm and nm in norm(f.name))]
    return max(cands)[1] if cands else None


def _latest_receipt(matter):
    """回单区（工作/银行回单，含各泳道子夹）按事项名尽力匹配，取最新一张。
    精确归属口径（金额/台账）待真实回单到位后再固化。"""
    root = paths.DATA / "工作/银行回单"
    if not root.is_dir():
        return None
    exts = {".pdf", ".png", ".jpg", ".jpeg"}
    return _pick_latest([f for f in root.rglob("*") if f.suffix.lower() in exts], matter)


def _resolve(item, matter):
    func, label = item
    if func == "__回单__":
        return _latest_receipt(matter), label
    return _latest_deliver(func, matter), label


def build_browse_folder(code, matter, amt, big, done_date=None, dry=False):
    """从交付区/回单区组装一册定稿 → 归档/架/code-matter〔办结〕。
    主件按业务顺序进册根；验收单/付款单等生成物料单列「生成物料」子夹并标注。"""
    done_date = done_date or datetime.date.today().strftime("%Y%m%d")
    shelf = paths.ARCHIVED_LARGE if big else paths.ARCHIVED_SMALL
    try:                                   # 金额美化：整数去掉 .0
        amt_disp = int(float(amt)) if amt not in (None, "") else ""
    except (TypeError, ValueError):
        amt_disp = amt
    amt_tag = f"·金额：{amt_disp}" if amt_disp not in (None, "") else ""
    dest = shelf / f"{code}-{matter}〔办结{done_date}〕"
    picked = []

    def emit(items, target_dir):
        avail = [r for r in (_resolve(it, matter) for it in items) if r[0]]
        for k, (src, label) in enumerate(avail):
            out = target_dir / f"{_ord(k)}{label}（{matter}{amt_tag}）{src.suffix}"
            if not dry:
                target_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(str(src), str(out))
            picked.append(src)

    emit(BROWSE_ORDER, dest)                # 主件（真实/生效凭证）
    emit(GEN_ITEMS, dest / GEN_SUBDIR)      # 生成物料：单列子夹
    return dest, picked


def remove_browse_folder(code):
    """取出：把该编号的定稿册移进 运行/state/归档取出回收（零删除，可再找回）。
    回收目标名带时分秒 + 撞名自动加序号，避免同日重复归档/取出互相覆盖。"""
    trash = paths.STATE / "归档取出回收"
    moved = []
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    for shelf in (paths.ARCHIVED_SMALL, paths.ARCHIVED_LARGE):
        if not shelf.is_dir():
            continue
        for d in shelf.iterdir():
            if d.is_dir() and _code_matter(d.name)[0] == code:
                trash.mkdir(parents=True, exist_ok=True)
                base = f"{stamp}-{d.name}"
                dst = trash / base
                n = 1
                while dst.exists():
                    n += 1
                    dst = trash / f"{base}#{n}"
                d.rename(dst)
                moved.append(dst)
    return moved


# ---------- 台账快照：整套记录型台账随档留一份 ----------
# 通配“记录/台账”里所有含“台账”的 .xlsx（自动适配 01_~06_ 前缀），归档台账另计不重复。
_SNAP_EXCLUDE = "归档台账"        # 归档台账已是 归档/归档台账.xlsx 分身，不再进快照
_SNAP_SUBDIRS = ["单据流水"]      # 采购台账_YYYY / 付款台账_YYYY（真实账号原样随档，用户本机）


def refresh_snapshots(dry=False):
    """把 记录/台账 的整套记录型台账 + 进度总览刷新进 归档/台账快照。
    改名/换前缀后残留的旧快照副本移入 state/快照旧版回收（零删除），保持快照与源一致。"""
    snap = paths.DATA / "归档/台账快照（随档留一份）"
    src_led = _ledger_dir()
    src_files = [f for f in sorted(src_led.glob("*.xlsx"))
                 if "台账" in f.name and _SNAP_EXCLUDE not in f.name and not f.name.startswith("~")]
    want = {f.name for f in src_files}
    copied = []
    if not dry:
        snap.mkdir(parents=True, exist_ok=True)
        # 清掉源里已不存在的旧快照（改名/删除导致）
        stale = paths.STATE / "快照旧版回收"
        for g in snap.glob("*.xlsx"):
            if g.name not in want and not g.name.startswith("~"):
                stale.mkdir(parents=True, exist_ok=True)
                g.rename(stale / f"{datetime.datetime.now():%Y%m%d-%H%M%S}-{g.name}")
        for f in src_files:
            shutil.copy2(str(f), str(snap / f.name))
            copied.append(f.name)
    else:
        copied = [f.name for f in src_files]
    for sub in _SNAP_SUBDIRS:
        sdir = src_led / sub
        if sdir.is_dir():
            t = snap / sub
            if not dry:
                t.mkdir(parents=True, exist_ok=True)
            for f in sorted(sdir.iterdir()):
                if f.suffix.lower() in (".xls", ".xlsx") and not f.name.startswith("~"):
                    if not dry:
                        shutil.copy2(str(f), str(t / f.name))
                    copied.append(f"{sub}/{f.name}")
    if paths.BOARD.exists():
        if not dry:
            shutil.copy2(str(paths.BOARD), str(paths.DATA / "归档/进度总览（快照）.xlsx"))
        copied.append("进度总览（快照）.xlsx")
    return copied


def project_dir(matter_name, create=True):
    """返回项目夹 Path（不建子目录）。命中既有夹（名含事项核心词）复用，否则续号。"""
    keys = [matter_name]
    for suf in ("申购单", "申购", "采购", "服务", "及耗材", "明细"):
        for k in list(keys):
            if k.endswith(suf) and len(k) - len(suf) >= 2:
                keys.append(k[:-len(suf)])
    nk = [norm(k) for k in keys if len(norm(k)) >= 2]
    if paths.ARCHIVE.is_dir():
        for d in sorted(paths.ARCHIVE.iterdir(), reverse=True):
            if d.is_dir() and any(k in norm(d.name) for k in nk if len(k) >= 3):
                return d
    if not create:
        return None
    year = datetime.date.today().year
    mx = 0
    if paths.ARCHIVE.is_dir():
        for d in paths.ARCHIVE.iterdir():
            m = re.match(rf"^{year}-(\d{{3}})", d.name)
            if m:
                mx = max(mx, int(m.group(1)))
    return paths.ARCHIVE / f"{year}-{mx + 1:03d}-{matter_name}"
