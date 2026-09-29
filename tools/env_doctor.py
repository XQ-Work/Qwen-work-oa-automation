# -*- coding: utf-8 -*-
"""环境自检（迁移第5步跑它；--fix 可重建 junction）。

检查项：Python 依赖 / config 可解析 / 路径常量存在 / 三个功能 junction /
        OA 凭据在凭据管理器 / Edge 浏览器 / Word+Excel COM（交付转换用）。
用法：python tools/env_doctor.py [--fix]
退出码 0=全绿 1=有红项（会逐条给出补救命令）。
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
ok_n = red_n = yellow_n = 0
FIXES = []


def report(state, name, hint=""):
    global ok_n, red_n, yellow_n
    {"ok": (setattr, None)}
    mark = {"ok": "[绿]", "red": "[红]", "yellow": "[黄]"}[state]
    if state == "ok":
        ok_n += 1
    elif state == "red":
        red_n += 1
    else:
        yellow_n += 1
    print(f"{mark} {name}" + (f"   → {hint}" if hint and state != "ok" else ""))


def check_imports():
    mods = [("openpyxl", "openpyxl"), ("playwright", "playwright"), ("keyring", "keyring"),
            ("fitz/pymupdf", "pymupdf"), ("docx", "docx"), ("win32com", "win32com.client"),
            ("pdfplumber", "pdfplumber"), ("lxml", "lxml.etree"), ("xlrd", "xlrd"),
            ("xlwt", "xlwt"), ("pandas", "pandas")]
    for name, mod in mods:
        try:
            __import__(mod)
            report("ok", f"依赖 {name}")
        except Exception:
            report("red", f"依赖 {name}", f"pip install -r {BASE/'requirements.txt'}")


def check_paths():
    try:
        from oa_common import paths
    except Exception as e:
        report("red", "config.json 解析", f"检查 {BASE/'config.json'}：{e}")
        return
    miss = [k for k in ["DB", "DOCS_DB", "LEDGER_OA", "LEDGER_SG", "LEDGER_SUPPLIER",
                        "LEDGER_CONTRACT", "LEDGER_PAY", "RECEIPT_DIR", "DANJU_BASE",
                        "INBOX", "TPL", "ARCHIVE", "DELIVER", "PROFILE", "BOARD"]
            if not getattr(paths, k, Path("x")).exists()]
    if miss:
        report("yellow", "数据路径齐备", f"缺失(用到会自动建): {miss}")
    else:
        report("ok", "数据路径齐备（data_root=" + str(paths.DATA) + "）")
    return paths


def check_junctions(paths):
    links = {
        paths.DANJU_BASE / "80_工具": (paths.ROOT / "danju" / "80_工具", "ocr_local.ps1"),
        paths.DANJU_BASE / "50_模板": (paths.DATA / "系统/模板/03_单据类", "付款单模板.xlsx"),
    }
    for link, (target, probe) in links.items():
        if (link / probe).exists():
            report("ok", f"junction {link.name} → {target}")
        else:
            report("red", f"junction 失效/缺失: {link}", "python tools/env_doctor.py --fix 重建")
            FIXES.append((str(link), str(target)))


def check_auth():
    try:
        import keyring
        from oa_common import auth
        acc = auth._load_cfg().get("account")
        if not acc:
            report("yellow", "OA 凭据(凭据管理器)",
                   "无已存账号：首次发草稿(fill_*)会弹窗要求输入一次账号密码，之后自动保存")
        elif keyring.get_password(auth.SERVICE, acc):
            report("ok", f"OA 凭据已存（账号 {acc}）")
        else:
            report("yellow", "OA 凭据(有账号无密码)", "下次运行 fill_* 弹窗补一次密码即可")
    except Exception as e:
        report("red", "OA 凭据读取", str(e)[:60])


def check_office_edge():
    try:
        import win32com.client  # noqa
        win32com.client.Dispatch("Word.Application").Quit()
        win32com.client.Dispatch("Excel.Application").Quit()
        report("ok", "Word/Excel COM（交付PDF转换）")
    except Exception as e:
        report("red", "Word/Excel COM", "需装桌面版 Office：" + str(e)[:60])
    edge = Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
    edge2 = Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe")
    report("ok" if (edge.exists() or edge2.exists()) else "red", "Edge 浏览器", "安装 Microsoft Edge")


def check_ocr():
    ps = ("[void][Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]; "
          "[void][Windows.Media.Ocr.OcrEngine,Windows.Media,ContentType=WindowsRuntime]; "
          "$l=[Windows.Globalization.Language]::new('zh-Hans-CN'); "
          "[bool]([Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($l))")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                           capture_output=True, text=True, timeout=30,
                           encoding="utf-8", errors="replace")
        report("ok" if "True" in (r.stdout or "") else "yellow",
               "Windows 中文OCR引擎", "设置→时间和语言→语言→添加中文(中国) 以启用OCR")
    except Exception:
        report("yellow", "Windows 中文OCR引擎", "手动检查 powershell")


def do_fix():
    for link, target in FIXES:
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        f"if(Test-Path '{link}'){{ }}else{{New-Item -ItemType Junction -Path '{link}' -Target '{target}' | Out-Null}}"])
        print("重建 junction:", link, "->", target)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fix", action="store_true")
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    print("=== OA 自动化环境自检 ===")
    check_imports()
    paths = check_paths()
    if paths:
        check_junctions(paths)
    check_auth()
    check_office_edge()
    check_ocr()
    print(f"\n绿 {ok_n} ｜ 黄 {yellow_n} ｜ 红 {red_n}")
    if a.fix and FIXES:
        do_fix()
    sys.exit(1 if red_n else 0)


if __name__ == "__main__":
    main()
