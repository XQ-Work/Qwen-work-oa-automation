# -*- coding: utf-8 -*-
"""把 20_输出\\验收单 下的 docx 批量转成 PDF（用本机 Word 组件）。

口径与主流程一致：docx 留在 20_输出\\验收单，PDF 统一输出到 30_待打印。
主流程在沙箱/受限环境下 Word 组件无法启动时先出 docx，本脚本负责补齐 PDF 环节。

用法：
    python 80_工具\\to_pdf.py            # 转换所有尚无最新 PDF 的验收单
    python 80_工具\\to_pdf.py <文件或目录>  # 指定单个 docx 或目录
"""
import os
import re
import subprocess
import sys
import zipfile

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
import paths

_main_src = open(os.path.join(_ROOT, "一键生成.py"), encoding="utf-8").read()
_ns = {"os": os, "re": re, "zipfile": zipfile, "subprocess": subprocess}
exec(_main_src[_main_src.index("def _strip_revisions"):_main_src.index("if docx_files:")], _ns)
_docx_to_pdf = _ns["_docx_to_pdf"]


def collect(target):
    if os.path.isdir(target):
        return [os.path.join(target, f) for f in sorted(os.listdir(target))
                if f.lower().endswith(".docx") and not f.startswith("~$")]
    return [target]


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else paths.ACCEPT_DIR
    out_dir = paths.PRINT_DIR if os.path.isdir(arg) else os.path.dirname(os.path.abspath(arg))
    os.makedirs(out_dir, exist_ok=True)
    todo = []
    for p in collect(arg):
        pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(p))[0] + ".pdf")
        if os.path.exists(pdf) and os.path.getmtime(pdf) >= os.path.getmtime(p):
            print("跳过（PDF 已是最新）: %s" % os.path.basename(pdf))
            continue
        todo.append(p)

    if not todo:
        print("没有需要转换的文件。")
        return

    print("待转换 %d 个文件（Word 组件逐个导出，请勿关闭 Word）" % len(todo))
    ok, fail = 0, []
    for i, p in enumerate(todo, 1):
        name = os.path.basename(p)
        try:
            done = _docx_to_pdf(p, out_dir)
        except Exception as e:
            done = False
            print("  [%d/%d] %s -> 异常: %s" % (i, len(todo), name, e))
        if done:
            print("  [%d/%d] %s -> PDF 完成" % (i, len(todo), name))
            ok += 1
        else:
            print("  [%d/%d] %s -> [WARNING] 转换失败" % (i, len(todo), name))
            fail.append(name)
    print("\n完成：%d 成功 / %d 失败" % (ok, len(fail)))
    if fail:
        print("失败清单：" + ", ".join(fail))


if __name__ == "__main__":
    main()
