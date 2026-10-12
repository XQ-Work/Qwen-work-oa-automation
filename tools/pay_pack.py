# -*- coding: utf-8 -*-
"""付款资料装配（流程固定步）：按“付款必备资料清单”从项目档案挑齐 → 交付/<事项>付款/。
缺的点名提示、不硬凑。用法：
    python tools/pay_pack.py 2026-003                      # 自动判定类别
    python tools/pay_pack.py 2026-003 --pay <本次付款单.pdf>
清单（合同类）：付款单(本次) · 采购合同(已用印) · 流程会签 · 发票 · 验收资料(货物/工程) · 前次付款凭证(分期)
清单（日常类<2000无合同）：付款单 · 发票 · 验收资料 · 事项审批/申购单
"""
import argparse, re, shutil, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from oa_common import paths, moneyline

MONEY = lambda x: f"{x:,.2f}"


def find(files, must, avoid=(), prefer=()):
    """按关键词挑一份：must 全含、avoid 不含、prefer 命中者优先（如“已用印”）。"""
    cand = [f for f in files if all(k in f.name for k in must) and not any(a in f.name for a in avoid)]
    if not cand:
        return None
    cand.sort(key=lambda f: (0 if any(p in f.name for p in prefer) else 1, len(f.name)))
    return cand[0]


def build(code, pay_pdf=None):
    proj = next((d for d in sorted(paths.ARCHIVE.iterdir())
                 if d.is_dir() and d.name.startswith(code + "-")), None)
    if not proj:
        print("项目档案里没有:", code); return 1
    matter = re.sub(r"^\d{4}-\d{3}-", "", proj.name)
    files = [f for f in proj.rglob("*") if f.is_file()
             and not f.name.startswith(("~", "项目文件目录")) and f.suffix.lower() in (".pdf", ".doc", ".docx", ".jpg", ".png")]
    has_contract = any("合同" in f.name for f in files)
    cat = (moneyline.get(code) or {}).get("category") or ""
    paid = (moneyline.get(code) or {}).get("paid_total") or 0
    kind = "合同类" if has_contract else "日常类"

    # (清单项, 取法)
    plan = []
    if kind == "合同类":
        plan.append(("01 付款单（本次）", pay_pdf or _latest_pay(matter)))
        plan.append(("02 采购合同（已用印）", find(files, ["合同"], prefer=["已用印"])))
        plan.append(("03 流程会签", find(files, ["流程会签"])))
        plan.append(("04 发票", find(files, ["发票"])))
        if any(k in cat for k in ("货物", "工程")):
            plan.append(("05 验收资料", find(files, ["验收资料"]) or find(files, ["验收"])))
        if paid:
            plan.append(("06 前次付款凭证", find(files, ["付款单"], avoid=["本次"])))
    else:
        plan.append(("01 付款单（本次）", pay_pdf or _latest_pay(matter)))
        plan.append(("02 发票", find(files, ["发票"])))
        plan.append(("03 验收资料", find(files, ["验收资料"]) or find(files, ["验收"])))
        plan.append(("04 事项审批/申购单", find(files, ["事项审批"]) or find(files, ["申购单"])))

    out = paths.DELIVER / f"{matter}付款"
    out.mkdir(parents=True, exist_ok=True)
    print(f"\n【{code} {matter}】类别={kind}/{cat or '?'} 已付={MONEY(paid)} → {out.relative_to(paths.DATA)}")
    miss = []
    for label, src in plan:
        if src and Path(src).exists():
            dst = out / f"{label.split(' ', 1)[0]}{Path(src).name}"
            shutil.copy2(str(src), str(dst))
            print(f"  ✓ {label:<16} ← {Path(src).name}")
        else:
            print(f"  ✗ {label:<16} 缺（需先备好/生成）")
            miss.append(label)
    if miss:
        print("  ⚠ 缺项:", "、".join(miss), "——按约定可先不处理，但送签前须补齐")
    return 0


def _latest_pay(matter):
    d = paths.DELIVER / "付款单"
    if not d.is_dir():
        return None
    hits = [f for f in d.glob("*.pdf") if matter[:3] in f.name]
    return sorted(hits)[-1] if hits else None


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("code"); ap.add_argument("--pay", default=None, help="本次付款单PDF路径")
    a = ap.parse_args()
    raise SystemExit(build(a.code.strip().rstrip("/\\").split("\\")[-1], a.pay))
