# -*- coding: utf-8 -*-
"""生成《附件命名规范.md》：列出 m05 上传要求的 4 个必传件命名规则 + 每个项目
该用的“事项”和 4 个目标文件名，并标注当前是否已对上。"""
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
sys.path.insert(0, str(BASE / "modules" / "m05_contract"))
from oa_common import paths  # noqa: E402
import extract_contract as ex  # noqa: E402
import fill_contract as fc  # noqa: E402

ARCH = paths.DATA / "OA附件库/文档整理/采购项目档案"
OUT = paths.DATA / "附件命名规范.md"
PROJECTS = [
    "2026-002-滤布采购", "2026-003-潜污泵采购", "2026-004-污水池堵漏",
    "2026-005-发电机排涝采购", "2026-007-细格栅维修", "2026-008-维修用品采购",
    "2026-009-垃圾清运服务", "2026-010-机械油及电气配件采购",
    "2026-011-化验室药剂及器具",
]
DOCS = ["事项审批", "比质比价报告单", "供应商报价单及资质", "采购合同"]

lines = []
lines.append("# 合同会签(m05)附件命名规范\n")
lines.append("上传到「过程文件上传」的 4 个必传件，文件名必须严格按下式命名，"
             "缺任何一个程序会报“缺少文件”并中止：\n")
lines.append("```\n事项审批（事项）.pdf\n比质比价报告单（事项）.pdf\n"
             "供应商报价单及资质（事项）.pdf\n采购合同（事项）.pdf\n```")
lines.append("\n- “事项”= 项目文件夹 `2026-NNN-XXX` 去掉编号前缀、再去掉尾部“采购/服务”"
             "（“维修”等实义词保留）。下表已按此算好。\n"
             "- 4 个文件放在该项目文件夹内任意子目录即可（程序递归找）。\n"
             "- 盖章扫描版（含“已用印/盖章”）不会被当正文附件，可另存。\n")
lines.append("\n---\n")

for proj in PROJECTS:
    d = ARCH / proj
    if not d.exists():
        continue
    # 事项：用合同路径推(合同可能在04_采购合同)
    cdir = d / "04_采购合同"
    contract = None
    if cdir.exists():
        cs = [f for f in cdir.glob("*.pdf")
              if not any(s in f.name for s in ("已用印", "盖章", "扫描"))]
        contract = cs[0] if cs else None
    kw = ex.keyword_from_path(contract) if contract else ex.keyword_from_path(str(d) + "\\x\\")
    if not kw:
        # 直接从文件夹名算
        import re
        m = re.search(r"-(.+)$", proj)
        s = m.group(1)
        for suf in ("采购合同", "采购", "服务合同", "服务"):
            if s.endswith(suf) and len(s) > len(suf) + 1:
                s = s[:-len(suf)]; break
        kw = s
    lines.append(f"## {proj}")
    lines.append(f"- 事项（关键词）：**{kw}**")
    for doc in DOCS:
        want = f"{doc}（{kw}）.pdf"
        found = list(d.rglob(want))
        mark = "✅ 已就位" if found else "❌ 需改名/补齐 → 命名为 `" + want + "`"
        lines.append(f"  - {doc}：{mark}")
    if not contract:
        lines.append("  - （注意：04_采购合同 下暂无文字版合同 PDF）")
    lines.append("")

OUT.write_text("\n".join(lines), encoding="utf-8")
print("已生成:", OUT)
print("\n".join(lines))
