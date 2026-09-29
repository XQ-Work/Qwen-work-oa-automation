# -*- coding: utf-8 -*-
"""生成《附件命名规范.md》（记录/台账/附件命名规范.md）：
总命名格式 → 材料清单与 01~06 阶段子夹 → 会签/付款必传件规则 → 逐项目就位核对。
项目清单自动扫描 记录/项目档案，不硬编码；run/flush 后可顺带重跑（非自动）。
用法：python tools/gen_naming_doc.py
"""
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
sys.path.insert(0, str(BASE / "modules" / "m05_contract"))
from oa_common import paths  # noqa: E402
import extract_contract as ex  # noqa: E402

ARCH = paths.ARCHIVE
OUT = paths.NAMING_DOC
PROJECTS = [d.name for d in sorted(ARCH.iterdir())
            if d.is_dir() and re.match(r"^\d{4}-\d{3}-", d.name)] if ARCH.is_dir() else []
DOCS = ["事项审批", "比质比价报告单", "供应商报价单及资质", "采购合同"]

L = []
L.append("# 附件命名规范\n")
L.append("## 一、总格式\n")
L.append("```\n材料名（事项名）.pdf　　可带金额：采购合同（潜污泵采购-金额：15877）.pdf\n```")
L.append("\n“事项名”= 项目文件夹 `2026-NNN-XXX` 去编号前缀、再去尾部“采购/服务”"
         "（“维修”等实义词保留）。材料名固定词表与归档位置：\n")
L.append("| 材料名 | 归入阶段夹 | 用途 |\n|---|---|---|\n"
         "| 申购单 | 01_事项审批 | 事项审批附件、合同会签第1件来源 |\n"
         "| 事项审批 | 01_事项审批 | 合同会签必传 |\n"
         "| 供应商报价单及资质 | 02_供应商报价及资质 | 合同会签必传 |\n"
         "| 比质比价报告单 | 03_比质比价报告单 | 合同会签必传 |\n"
         "| 采购合同 | 04_采购合同 | 合同会签必传（未盖章原件） |\n"
         "| 流程会签 / 采购合同（已用印）/ 发票 / 付款单 | 05_付款及验收 | 付款申请必传 |\n"
         "| 回单 | 05_付款及验收（M07匹配后复制入） | 结清凭证 |\n"
         "| 其他 | 06_其他 | 非常规附件，自由区不进核对 |\n")
L.append("\n投料口=根目录 `输入/`（平铺），系统按名自动归档；已按要求在阶段夹内的文件不必再投。\n")
L.append("\n## 二、合同会签(m05) 4 必传件\n\n缺任何一个程序报“缺少文件”并中止：\n")
L.append("```\n事项审批（事项）.pdf\n比质比价报告单（事项）.pdf\n"
         "供应商报价单及资质（事项）.pdf\n采购合同（事项）.pdf\n```")
L.append("\n- 同一材料有多份（如两份事项审批）：都按 `材料名（事项）xxx.pdf` 前缀命名即全部上传。\n"
         "- 盖章扫描版（含“已用印/盖章”）不会被当会签正文附件——那是付款侧要的东西。\n")
L.append("\n## 三、付款申请(m06) 必传件\n\n```\n流程会签（事项）.pdf\n"
         "采购合同（事项 已用印）.pdf\n发票（事项）.pdf\n付款单*（事项）.pdf\n```\n"
         "（付款单允许“付款单、发票（事项）.pdf”式合并件，前缀“付款单”即可命中）\n")
L.append("\n---\n")

for proj in PROJECTS:
    d = ARCH / proj
    cdir = d / "04_采购合同"
    contract = None
    if cdir.is_dir():
        cs = [f for f in cdir.glob("*.pdf")
              if not any(s in f.name for s in ("已用印", "盖章", "扫描"))]
        contract = cs[0] if cs else None
    kw = (ex.keyword_from_path(contract) if contract
          else ex.keyword_from_path(str(d) + "\\x\\"))
    if not kw:
        s = re.sub(r"^\d{4}-\d{3}-", "", proj)
        for suf in ("采购合同", "采购", "服务合同", "服务"):
            if s.endswith(suf) and len(s) > len(suf) + 1:
                s = s[:-len(suf)]
                break
        kw = s
    L.append(f"\n## 项目 {proj}")
    L.append(f"- 事项名（括号里填这个）：**{kw}**")
    for doc in DOCS:
        want = f"{doc}（{kw}）*.pdf"
        found = list(d.rglob(want))
        L.append(f"  - {doc}：{'✅ 已就位' if found else '❌ 待补 → 命名为 `' + want.replace('*', '') + '`，投进 输入/'}")
    for doc, pat in [("流程会签", "流程会签"), ("发票", "发票"), ("付款单", "付款单")]:
        want = f"{pat}（{kw}）*.pdf"
        found = list((d / "05_付款及验收").glob(want)) if (d / "05_付款及验收").is_dir() else []
        L.append(f"  - 付款侧·{doc}：{'✅' if found else '－（未到付款环节可忽略）'}")
    if not contract:
        L.append("  - （注意：04_采购合同 下暂无未盖章合同 PDF）")

OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text("\n".join(L) + "\n", encoding="utf-8")
print("已生成:", OUT, "| 项目数:", len(PROJECTS))
