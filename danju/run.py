# -*- coding: utf-8 -*-
"""一键运行入口：自动挑选装有依赖的 Python 解释器执行主流程。

背景：本机 PATH 中的 python 可能解析到无依赖的内置运行时；
本项目依赖装在用户 Python 下。本脚本探测各解释器的依赖可用性，
选定其一运行 一键生成.py。也可以直接双击 一键生成.bat。
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
PROBE = ("import xlrd, xlwt, xlutils.copy, lxml.etree, pdfplumber, "
         "win32print, pandas, openpyxl")

CANDIDATES = [
    r"C:\Users\JQZH\AppData\Local\Programs\Python\Python313\python.exe",
    r"C:\Users\XQ\AppData\Local\Programs\Python\Python314\python.exe",
    sys.executable,
]

chosen = None
for py in CANDIDATES:
    if not os.path.exists(py):
        continue
    try:
        r = subprocess.run([py, "-c", PROBE], capture_output=True, timeout=60)
    except OSError:
        continue
    if r.returncode == 0:
        chosen = py
        break

if chosen is None:
    print("未找到已安装依赖的 Python。请先执行：")
    print("  python -m pip install -r requirements.txt")
    sys.exit(1)

print("使用解释器: %s" % chosen)
os.chdir(ROOT)
if len(sys.argv) > 1 and sys.argv[1] == "report":
    # 报告链路：批归档(40_) + 运行日志(70_) + 待打印暂存(30_) + 主流程
    script = os.path.join(ROOT, "80_工具", "run_report.py")
else:
    script = os.path.join(ROOT, "一键生成.py")
sys.exit(subprocess.call([chosen, script]))
