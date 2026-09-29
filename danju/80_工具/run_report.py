# -*- coding: utf-8 -*-
"""运行报告 run_report.py（D3）：运行主脚本并落盘分类汇总日志。

用法（在运行环境根目录运行）：
  python 80_工具\\run_report.py

按 D3 决定：每次运行把 一键生成.py 的完整输出 + 成功/跳过/失败分类汇总
写入 70_日志\\运行报告_YYYYMMDD_HHMMSS.log，不再让产出"静默消失"。
"""
import sys, os, subprocess, datetime, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths, archive_outputs

LOG_DIR = os.path.join(paths.BASE, "70_日志")
os.makedirs(LOG_DIR, exist_ok=True)
SCRIPT = os.path.join(paths.BASE, "一键生成.py")

# 运行前先归档上一批，让 20_输出 只保留最新一批（台账除外）
archived = 0
try:
    archived = archive_outputs.archive()
except Exception as e:
    print("归档旧批失败（继续运行主流程）: %s" % e)

res = subprocess.run(
    [sys.executable, SCRIPT], cwd=paths.BASE,
    capture_output=True, text=True, encoding="utf-8", errors="replace")
out = res.stdout or ""
err = res.stderr or ""

# 主流程完成后，把本批成品复制到 30_待打印（幂等去重）
staged = 0
try:
    staged = archive_outputs.stage_for_print()
except Exception as e:
    print("暂存待打印失败（忽略）: %s" % e)

now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
logpath = os.path.join(LOG_DIR, "运行报告_%s.log" % now)

skip = len(re.findall(r"跳过[:：]", out))
warn = len(re.findall(r"WARNING", out))
accept = len(re.findall(r"验收单", out))
pay = len(re.findall(r"付款单", out))
done = re.findall(r"全部完成!.*", out)
done_txt = done[0] if done else "（未打印完成行）"

lines = []
lines.append("=" * 70)
lines.append("运行报告  生成时间: %s" % now)
lines.append("脚本      : %s" % SCRIPT)
lines.append("数据库    : %s" % paths.DB_PATH)
lines.append("退出码    : %s" % res.returncode)
lines.append("本批归档  : %s 个文件移入 40_归档" % archived)
lines.append("待打印    : %s 个文件复制到 30_待打印" % staged)
lines.append("-" * 70)
lines.append("分类汇总  : 跳过=%d  WARNING=%d  付款单提及=%d  验收单提及=%d" % (skip, warn, pay, accept))
lines.append("最终结果  : %s" % done_txt)
lines.append("=" * 70)
lines.append(out.rstrip("\n"))
if (err or "").strip():
    lines.append("")
    lines.append("--------------------- stderr ---------------------")
    lines.append(err.rstrip("\n"))

with open(logpath, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print("运行报告已生成: %s" % logpath)
print("分类汇总: 跳过=%d  WARNING=%d  付款单提及=%d  验收单提及=%d" % (skip, warn, pay, accept))
print("最终结果: %s" % done_txt)