# -*- coding: utf-8 -*-
"""流程收尾一步：把「用过的投料源件」归档进 记录/已用投料（带日期前缀，零删除、只移动）。

典型用法——从一张汇总《申购单》Excel 解析出需求台账并生成申购单后，把源表退役：
    python tools/retire_input.py "输入/申购单.xlsx"
    python tools/retire_input.py --all      # 退役 输入/ 里所有非临时文件（谨慎）

同名冲突自动加时间戳，绝不覆盖、绝不删除。"""
import argparse
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from oa_common import paths


def retire(src: Path) -> Path:
    dst_dir = paths.USED_INBOX
    dst_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.date.today().strftime("%Y%m%d")
    dst = dst_dir / f"{stamp}-{src.name}"
    n = 1
    while dst.exists():
        n += 1
        dst = dst_dir / f"{stamp}-{src.stem}#{n}{src.suffix}"
    src.rename(dst)
    return dst


def main():
    ap = argparse.ArgumentParser(description="退役已用投料源件到 记录/已用投料")
    ap.add_argument("files", nargs="*", help="要退役的文件（绝对或相对 data_root）")
    ap.add_argument("--all", action="store_true", help="退役 输入/ 下所有文件（投料口清空，谨慎）")
    a = ap.parse_args()
    targets = []
    if a.all:
        if paths.INBOX.is_dir():
            targets = [f for f in paths.INBOX.iterdir() if f.is_file() and not f.name.startswith("~")]
    for s in a.files:
        p = Path(s)
        if not p.is_absolute():
            p = paths.DATA / s
        targets.append(p)
    if not targets:
        print("没有要退役的文件（给路径或 --all）")
        return 1
    for p in targets:
        if not p.exists():
            print("跳过（不存在）:", p); continue
        d = retire(p)
        print(f"[退役] {p.name} -> 记录/已用投料/{d.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
