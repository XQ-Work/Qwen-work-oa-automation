# -*- coding: utf-8 -*-
"""OA 070801 事项审批 填单脚本

用法：python fill_oa.py --row 4 --save   # 正式：填表->自动保存草稿->归档收件箱文件->回写台账
      python fill_oa.py --row 2          # 干跑：只填不存，浏览器停留供人工核对
      python fill_oa.py --row 4 --archive [--flow-no XXX]  # 仅归档回写
"""
import argparse
import datetime
import json
import time
from pathlib import Path
import sys as _sys
_sys.path.insert(0, r"D:/自动备份/Qwen work/自动化（OA流程）/系统/oa-automation")
from oa_common import paths

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from oa_common import auth as oa_auth

BASE = paths.STATE
LEDGER = paths.LEDGER_OA
ADD_PATH = ("/workflow/request/AddRequest.jsp"
            "?workflowid=475401&isagent=0&beagenter=0&f_weaver_belongto_userid=")
COMPANY_SHORT = "荆州荆清"
KEEP_OPEN_SECONDS = 15 * 60


def read_row(rownum):
    wb = load_workbook(LEDGER, data_only=True)
    ws = wb["发起台账"]
    rec = {}
    for col, key in COLS.items():
        v = ws.cell(row=rownum, column=col).value
        rec[key] = ("" if v is None else str(v)).strip()
    wb.close()
    if not rec["事由"]:
        raise SystemExit(f"台账第{rownum}行『事由』为空，无法干跑")
    return rec


SMALL_AMOUNT_LIMIT = 2000

# 申购缘由固定三选一（2026-09-12 用户约定，扩充须由用户告知）
REASONS = ("维修需要", "日常生产需要", "化验室日常化验需要")


def normalize_reason(text):
    """把现状原因/申购说明归类到三种缘由；无法归类返回 None"""
    t = (text or "").strip()
    if not t:
        return None
    for r in REASONS:
        if r in t:
            return r
    if any(k in t for k in ("化验", "检测", "试剂")):
        return "化验室日常化验需要"
    if any(k in t for k in ("生产", "运行", "加药", "脱水", "排污")):
        return "日常生产需要"
    if any(k in t for k in ("维修", "更换", "坏损", "老化", "故障")):
        return "维修需要"
    return None


def gen_texts(rec):
    """按写法约定生成 概述/内容；人工覆盖列非空则原样使用。
    金额≤2000：通用简单句式（事项审批即完结）；>2000：详述句式。"""
    try:
        amt = float(rec["金额"]) if rec["金额"] else None
    except ValueError:
        amt = None
    subject = rec["标的"] or rec["事由"]
    overview = rec["概述覆盖"]
    content = rec["内容覆盖"]
    if amt is not None and amt <= SMALL_AMOUNT_LIMIT:
        if not overview:
            reason = normalize_reason(rec["现状原因"]) or "维修需要"
            overview = f"荆州荆清因{reason}申购{subject}一批"
        if not content:
            content = f"{overview}，金额{amt:.2f}元。"
    else:
        if not overview:
            parts = [COMPANY_SHORT, subject]
            if rec["数量"]:
                parts.append(rec["数量"])
            parts.append("申购")
            overview = "".join(parts)
        if not content:
            sents = [s for s in (rec["现状原因"], rec["规格"], rec["费用依据"]) if s]
            content = "。".join(s.rstrip("。") for s in sents) + "。" if sents else overview
    title = f"070801水务事项审批-{COMPANY_SHORT}-肖桥-{rec['事由']}"
    return title, overview, content


ANNEX_DIR = paths.INBOX.parent
ANNEX_EXTS = {".pdf", ".doc", ".docx", ".xls", ".xlsx",
              ".jpg", ".jpeg", ".png", ".zip", ".rar"}


COLS = {  # 台账列号 -> 键（v3：无项目编号列）
    2: "事由", 3: "事项类型", 4: "金额", 5: "标的", 6: "规格",
    7: "数量", 8: "现状原因", 9: "费用依据", 10: "概述覆盖",
    11: "内容覆盖", 12: "附件", 13: "状态",
}

STAGING_DIR = None  # 待发起区已退役（2026-09-28）：附件走 输入→项目档案 链路
ARCHIVE_DIR = paths.ARCHIVE


def _files_in(d):
    return sorted(str(x) for x in d.iterdir()
                  if x.is_file() and x.suffix.lower() in ANNEX_EXTS)


def _find_by_prefix(root, prefix):
    if prefix and root.is_dir():
        for proj in sorted(root.iterdir()):
            if proj.is_dir() and proj.name.startswith(prefix):
                return proj
    return None


def match_annexes(subject):
    """档案兜底（收件箱在 resolve_annexes 里更优先）：
    采购项目档案\\*\\{事由核心词}*\\01_事项审批\\ 内文件
    """
    if ARCHIVE_DIR.is_dir():
        keys = [_norm(k) for k in subject_keys(subject)]
        for proj in sorted(ARCHIVE_DIR.iterdir()):
            pn = _norm(proj.name)
            if proj.is_dir() and any(k in pn for k in keys if len(k) >= 3):
                stage = proj / "01_事项审批"
                if stage.is_dir():
                    return _files_in(stage)
    return []


INBOX = paths.INBOX
_STRIP_SUFFIXES = ("申购单", "申购", "申请", "采购", "服务", "事项", "整改")


def _norm(s):
    import re as _re
    return _re.sub(r"[（）()【】\[\]\s\-_·]", "", s or "")


def subject_keys(subject):
    """事由 -> 关键词集合：原词 + 逐层剥离业务后缀的核心词"""
    ks = {subject}
    changed = True
    while changed:
        changed = False
        for suf in _STRIP_SUFFIXES:
            for k in list(ks):
                if k.endswith(suf) and len(k) - len(suf) >= 2:
                    core = k[:-len(suf)]
                    if core not in ks:
                        ks.add(core)
                        changed = True
    return {k for k in ks if len(_norm(k)) >= 2}


def inbox_match(subject):
    """扫描收件箱：文件名(去括号规范化)或括号内项目名 与事由关键词双向包含即命中"""
    import re as _re
    if not INBOX.is_dir():
        return []
    keys = [_norm(k) for k in subject_keys(subject)]
    out = []
    for x in sorted(INBOX.iterdir()):
        if not (x.is_file() and x.suffix.lower() in ANNEX_EXTS):
            continue
        stem = _norm(x.stem)
        parens = [_norm(p) for p in _re.findall(r"[（(]([^（）()]+)[）)]", x.stem)]
        cands = [stem] + [p for p in parens if p]
        hit = any(
            (k in c) or (len(c) >= 2 and c in k)
            for k in keys for c in cands if c
        )
        if hit:
            out.append(str(x))
    return out


def resolve_annexes(rec):
    """统一附件解析，返回 (文件列表, 来源说明)。优先级：手填>收件箱>项目档案兜底"""
    if rec.get("附件"):
        return prepare_annexes(rec["附件"]), "手填路径"
    files = inbox_match(rec["事由"])
    if files:
        return files, f"收件箱({INBOX})"
    files = match_annexes(rec["事由"])
    if files:
        return files, "项目档案兜底"
    return [], "无匹配"


def sanitize_upload_names(files):
    """上传副本改名：去掉文件名中的“金额：X”段（本地留痕原名，OA 显示干净名）"""
    import re as _re
    import shutil
    up = paths.STATE / "upload_tmp"   # 机器房临时件，不留在业务视野
    up.mkdir(exist_ok=True)
    outs = []
    for f in files:
        p = Path(f)
        clean = _re.sub(r"[\s\-－]*金额[:：][\d.]+", "", p.stem)
        clean = _re.sub(r"[（(]\s*[)）]", "", clean).strip()
        dst = up / (clean or p.stem)
        dst = str(dst) + p.suffix
        shutil.copy2(f, dst)
        outs.append(dst)
    return outs


def prepare_annexes(raw):
    """附件路径处理：不存在则用同名占位PDF替代（干跑不碰真实资料）"""
    paths = [p.strip() for p in raw.replace("；", ";").split(";") if p.strip()]
    out = []
    dummy_dir = BASE / "dryrun"
    dummy_dir.mkdir(exist_ok=True)
    for p in paths:
        if Path(p).exists():
            out.append(p)
        else:
            d = dummy_dir / Path(p).name
            if not d.exists():
                d.write_bytes(
                    b"%PDF-1.4\n1 0 obj<</Type/Catalog>>endobj\ntrailer<</Root 1 0 R>>\n"
                    b"%%EOF")
            out.append(str(d))
    return out


def find_form_frame(page):
    """定位含表单控件的 frame（URL 特征优先，内容兜底）"""
    for _ in range(20):
        for f in page.frames:
            if "addrequestiframe" in f.url.lower():
                return f
        for f in page.frames:
            try:
                if f.locator("input#requestname").count() > 0:
                    return f
            except Exception:
                continue
        time.sleep(2)
    return None


def archive_project_name(rec):
    """项目文件夹简称：事由剥掉一次申购/申请类后缀"""
    s = rec["事由"]
    for suf in ("申购单", "申购", "申请", "事项"):
        if s.endswith(suf) and len(s) - len(suf) >= 2:
            return s[:-len(suf)]
    return s


def next_archive_folder(name):
    """档案自动续号：扫描现有 2026-0XX-xxx 取当年最大序号+1"""
    import re as _re
    year = datetime.date.today().year
    mx = 0
    if ARCHIVE_DIR.is_dir():
        for d in ARCHIVE_DIR.iterdir():
            m = _re.match(rf"^{year}-(\d{{3}})", d.name)
            if m:
                mx = max(mx, int(m.group(1)))
    return f"{year}-{mx + 1:03d}-{name}"


def do_archive(rec, rownum, flow_no="", dry=False, status="已归档"):
    """归档：收件箱匹配文件 -> 采购项目档案\\{自动编号}-{简称}\\01_事项审批；回写台账"""
    import shutil
    files = inbox_match(rec["事由"])
    name = archive_project_name(rec)
    # 优先复用已存在的项目夹（名含事由核心词），否则自动续新号
    dest_root = None
    if ARCHIVE_DIR.is_dir():
        keys = [_norm(k) for k in subject_keys(rec["事由"]) if len(k) >= 3]
        for d in sorted(ARCHIVE_DIR.iterdir(), reverse=True):
            if d.is_dir() and any(k in _norm(d.name) for k in keys):
                dest_root = d
                break
    if dest_root is None:
        dest_root = ARCHIVE_DIR / next_archive_folder(name)
    dest_dir = dest_root / "01_事项审批"
    print(f"[archive] 行{rownum} 事由={rec['事由']} -> {dest_dir.name}")
    moved = 0
    for f in files:
        dest = dest_dir / Path(f).name
        if dest.exists():
            dest = dest_dir / (dest.stem + "_" +
                               datetime.datetime.now().strftime("%H%M%S") + dest.suffix)
        print(f"  {'DRY ' if dry else ''}move {Path(f).name}")
        if not dry:
            dest_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(f, str(dest))
        moved += 1
    if not dry:
        wb = load_workbook(LEDGER)
        ws = wb["发起台账"]
        ws.cell(row=rownum, column=13, value=status)
        if flow_no:
            ws.cell(row=rownum, column=14, value=flow_no)
        ws.cell(row=rownum, column=15,
                value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        # 金额链：项目编号写17列 + 审批金额入主档
        try:
            from oa_common import moneyline
            code, matter = moneyline.resolve_code(rec["事由"], dest_root.name)
            if code:
                ws.cell(row=rownum, column=17, value=code)
                wb.save(LEDGER)
                if rec.get("金额"):
                    moneyline.set_amount(code, matter, "appr", rec["金额"],
                                         source=f"事项审批台账row{rownum}")
                    print(f"[moneyline] {code} 审批金额={rec['金额']} 已入主档")
        except Exception as e:
            print(f"[moneyline] 主档登记失败（不影响归档）: {e}")
        wb.save(LEDGER)
        print(f"[archive] 台账第{rownum}行回写：状态={status}"
              + (f"，编号={flow_no}" if flow_no else "") + f"，移动{moved}个文件")
    return moved


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--row", type=int, required=True)
    ap.add_argument("--no-annex", action="store_true")
    ap.add_argument("--archive", action="store_true",
                    help="办结归档模式：不打开浏览器，移动收件箱文件并回写台账")
    ap.add_argument("--flow-no", default="", help="OA流程编号（归档回写用）")
    ap.add_argument("--dry-archive", action="store_true", help="归档预演，不动文件")
    ap.add_argument("--save", action="store_true",
                    help="填完后自动点保存生成草稿并归档（提交仍由人工）")
    args = ap.parse_args()

    rec = read_row(args.row)
    if args.archive or args.dry_archive:
        do_archive(rec, args.row, args.flow_no, dry=args.dry_archive)
        return
    title, overview, content = gen_texts(rec)
    amount = f"{float(rec['金额']):.2f}" if rec["金额"] else ""
    report = {"row": args.row, "record": rec, "title": title,
              "overview": overview, "content": content, "amount": amount,
              "steps": []}

    def step(name, ok, err=""):
        report["steps"].append({"step": name, "ok": bool(ok), "err": str(err)[:200]})
        print(f"[{'ok ' if ok else 'ERR'}] {name} {err}", flush=True)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            base = oa_auth.login(page)
            step("登录", True, base)
            page.goto(base + ADD_PATH, wait_until="domcontentloaded", timeout=40000)
            frame = find_form_frame(page)
            if frame is None:
                step("定位表单frame", False)
                raise RuntimeError("未找到表单frame")
            step("定位表单frame", True, frame.url[:80])

            # 标题
            frame.locator("input#requestname").first.fill(title)
            step("填标题", True, title)

            # 金额
            frame.locator("input[name=field493135]").first.fill(amount)
            step("填金额", True, amount)

            # 事项类型（等级联选项加载）
            sel = frame.locator("select[name=field493136]").first
            for _ in range(15):
                n = frame.locator("select[name=field493136] option").count()
                if n > 1:
                    break
                time.sleep(1)
            sel.select_option(label=rec["事项类型"])
            step("选事项类型", True, rec["事项类型"])

            # 概述 / 内容（注意编号：概述=496131 内容=496130）
            frame.locator("textarea[name=field496131]").first.fill(overview)
            step("填事项概述", True, overview)
            frame.locator("textarea[name=field496130]").first.fill(content)
            step("填事项内容", True, content[:60])

            # 附件：手填路径 > 收件箱 > 项目档案兜底
            if args.no_annex:
                step("上传附件", None, "跳过")
            else:
                files, src = resolve_annexes(rec)
                report["annex_files"] = files
                report["annex_source"] = src
                if not files:
                    step("上传附件", None,
                         f"未匹配到附件（事由={rec['事由']}），请把「文档类型（项目名）.pdf」放入 {INBOX}")
                else:
                    print(f"[annex] {src}: " + "; ".join(Path(f).name for f in files),
                          flush=True)
                    btn = frame.get_by_text("选取多个文件").first
                    try:
                        upload_files = sanitize_upload_names(files)
                        print("[annex] 上传名: " + "; ".join(
                            Path(u).name for u in upload_files), flush=True)
                        with page.expect_file_chooser(timeout=15000) as fc:
                            btn.click()
                        fc.value.set_files(upload_files)
                        step("上传附件", True, f"{len(upload_files)}个")
                        time.sleep(4)  # 等上传完成
                    except Exception as e:
                        step("上传附件", False, e)

            time.sleep(2)
            page.screenshot(path=str(BASE / "dryrun_filled.png"), full_page=True)
            step("整页截图", True, "dryrun_filled.png")

            if args.save:
                # 点击主frame“保存”（页面自身JS doSave_nNew 处理校验与落库）
                signals = []
                page.on("dialog", lambda d: (signals.append(d.message or ""), d.accept()))
                try:
                    page.locator("input.e8_btn_top[value=保存]").first.click(timeout=8000)
                    step("点保存", True)
                except Exception as e:
                    step("点保存", False, e)
                reqid, saved = "", False
                for _ in range(20):
                    time.sleep(1.5)
                    u = page.url or ""
                    import re as _re
                    m = _re.search(r"requestid=(\d+)", u, _re.I)
                    if m:
                        reqid, saved = m.group(1), True
                        break
                    if any(("成功" in s) or ("success" in s.lower()) for s in signals):
                        saved = True
                        break
                    try:
                        body = page.inner_text("body")[:2000]
                    except Exception:
                        body = ""
                    if "保存成功" in body:
                        saved = True
                        break
                time.sleep(2)
                page.screenshot(path=str(BASE / "saved.png"), full_page=True)
                step("保存草稿", saved,
                     f"requestid={reqid}" if reqid else
                     ("弹窗确认成功" if saved else "未确认成功，请看 saved.png"))
                report["requestid"] = reqid
                report["finished_at"] = datetime.datetime.now().isoformat()
                (BASE / "dryrun_report.json").write_text(
                    json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
                if saved:
                    do_archive(rec, args.row, flow_no=reqid, status="已保存草稿")
                time.sleep(2)
                ctx.close()
                return 0 if saved else 2

            report["finished_at"] = datetime.datetime.now().isoformat()
            (BASE / "dryrun_report.json").write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"[hold] 浏览器保持打开 {KEEP_OPEN_SECONDS//60} 分钟供人工核对；"
                  f"确认无误可自行点提交；不会自动提交/保存。", flush=True)
            time.sleep(KEEP_OPEN_SECONDS)
        except Exception as e:
            step("异常", False, e)
            try:
                page.screenshot(path=str(BASE / "dryrun_error.png"))
            except Exception:
                pass
        finally:
            (BASE / "dryrun_report.json").write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            try:
                ctx.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()
