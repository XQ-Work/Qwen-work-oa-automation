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
    # 短名副本放系统临时目录且带本次运行 pid：30_待打印 里旧的 _tN.docx
    # 可能被僵死 Word 占着，撞名会让 copy2 直接崩
    import shutil, tempfile
    tdir = os.path.join(tempfile.gettempdir(), "danju_pdf_%d" % os.getpid())
    os.makedirs(tdir, exist_ok=True)
    for i, p in enumerate(todo, 1):
        name = os.path.basename(p)
        # 长文件名+全角括号会让 Word 拒写输出：复制为短临时名转换，再改回目标名
        ext = os.path.splitext(p)[1]
        short_p = os.path.join(tdir, "_t%d%s" % (i, ext))
        shutil.copy2(p, short_p)
        try:
            # 转换在临时目录完成短名 PDF，再改名落到目标目录（长名直写会被 Word 拒）
            done = _docx_to_pdf(short_p, tdir)
            produced = os.path.join(tdir, "_t%d.pdf" % i)
            target = os.path.join(out_dir, os.path.splitext(name)[0] + ".pdf")
            if done and os.path.exists(produced):
                shutil.move(produced, target)   # 跨盘：C 临时目录 → D 工作区
        except Exception as e:
            done = False
            print("  [%d/%d] %s -> 异常: %s" % (i, len(todo), name, e))
        finally:
            for junk in (short_p, os.path.join(tdir, "_t%d.pdf" % i)):
                try:
                    os.remove(junk)
                except OSError:
                    pass
        if done:
            print("  [%d/%d] %s -> PDF 完成" % (i, len(todo), name))
            ok += 1
        else:
            print("  [%d/%d] %s -> [WARNING] 转换失败" % (i, len(todo), name))
            fail.append(name)
    print("\n完成：%d 成功 / %d 失败" % (ok, len(fail)))
    try:
        shutil.rmtree(tdir, ignore_errors=True)
    except OSError:
        pass
    if fail:
        print("失败清单：" + ", ".join(fail))


if __name__ == "__main__":
    main()
