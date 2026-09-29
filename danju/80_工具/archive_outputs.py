# -*- coding: utf-8 -*-
"""归档脚本 archive_outputs.py：把 20_输出 的旧一批产物移入 40_归档。

用途：每次运行前调用，把上一批自动生成的验收单/付款单/发票/申购单归档，
让 20_输出 只保留"最新一批"；台账为持续累积文件，不做归档。

用法（被 run_report.py 调用；也可单独运行）：
  python 80_工具\\archive_outputs.py
"""
import os, sys, shutil, datetime
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import paths

# 持续累积、不随批次归档的目录名
SKIP = {"台账"}
# 参与"待打印"暂存的成品目录（台账不打印）
PRINT_CATEGORIES = ("验收单", "付款单", "发票", "申购单")


def archive():
    """把 20_输出 各成品子目录（除 SKIP）内的文件移到 40_归档\\<类型>\\<生成日期>，返回移动数。"""
    if not os.path.isdir(paths.OUTPUT_DIR):
        return 0
    today = datetime.date.today().strftime("%Y%m%d")
    moved = 0
    for name in sorted(os.listdir(paths.OUTPUT_DIR)):
        if name in SKIP:
            continue
        src = os.path.join(paths.OUTPUT_DIR, name)
        if not os.path.isdir(src):
            continue
        dest_root = os.path.join(paths.ARCHIVE_DIR, name, today)
        os.makedirs(dest_root, exist_ok=True)
        for f in sorted(os.listdir(src)):
            fp = os.path.join(src, f)
            if not os.path.isfile(fp):
                continue
            target = os.path.join(dest_root, f)
            stem, ext = os.path.splitext(f)
            i = 2
            while os.path.exists(target):
                target = os.path.join(dest_root, "%s_%d%s" % (stem, i, ext))
                i += 1
            shutil.move(fp, target)
            moved += 1
    return moved


def stage_for_print():
    """把 20_输出 最新一批的成品复制到 30_待打印（已存在则跳过，幂等）。

    30_待打印 作为待打印暂存区，按批累积去重；打印后由用户手动移走或清空。
    台账为累积文件，不进入待打印区。
    """
    if not os.path.isdir(paths.OUTPUT_DIR):
        return 0
    os.makedirs(paths.PRINT_DIR, exist_ok=True)
    staged = 0
    for name in PRINT_CATEGORIES:
        srcdir = os.path.join(paths.OUTPUT_DIR, name)
        if not os.path.isdir(srcdir):
            continue
        for f in sorted(os.listdir(srcdir)):
            fp = os.path.join(srcdir, f)
            if not os.path.isfile(fp):
                continue
            target = os.path.join(paths.PRINT_DIR, f)
            if os.path.exists(target):
                continue
            shutil.copy2(fp, target)
            staged += 1
    return staged


if __name__ == "__main__":
    n = archive()
    print("已归档 %d 个文件到 %s" % (n, paths.ARCHIVE_DIR))