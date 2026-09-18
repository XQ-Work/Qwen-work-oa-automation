# -*- coding: utf-8 -*-
"""批量体检：对 9 个采购项目的合同跑抽取预览 + 附件解析，离线不开浏览器。
输出每项 场景/关键词/乙方/金额 + 4 附件是否解析到。"""
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
sys.path.insert(0, str(BASE / "modules" / "m05_contract"))
from oa_common import paths  # noqa: E402
import extract_contract as ex  # noqa: E402
import fill_contract as fc  # noqa: E402

ARCH = paths.DATA / "OA附件库/文档整理/采购项目档案"
PROJECTS = {
    "2026-002-滤布采购": "04_采购合同/滤布采购合同.pdf",
    "2026-003-潜污泵采购": "04_采购合同/潜污泵采购合同.pdf",
    "2026-004-污水池堵漏": "04_采购合同/污水池堵漏施工合同.pdf",
    "2026-005-发电机排涝采购": "04_采购合同/发电机及排涝用品采购合同 .pdf",
    "2026-007-细格栅维修": "04_采购合同/服务合同（细格栅维修）.pdf",
    "2026-008-维修用品采购": "04_采购合同/采购合同（维修用品）.pdf",
    "2026-009-垃圾清运服务": "04_采购合同/服务合同（垃圾清运）.pdf",
    "2026-010-机械油及电气配件采购": "04_采购合同/采购合同（机械油及电气配件）.pdf",
    "2026-011-化验室药剂及器具": None,
}

for proj, rel in PROJECTS.items():
    print("=" * 70)
    print("项目:", proj)
    if not rel:
        print("  (无合同文件，跳过)")
        continue
    p = ARCH / proj / rel
    if not p.exists():
        print("  合同文件不存在:", rel)
        continue
    text = ex.read_text(p)
    items = ex.read_items(p)
    rec = ex.parse(text, items)
    scene = ex.detect_scene(text, p.stem)
    kw = ex.keyword_from_name(p.stem) or (rec["明细"][0]["名称"] if rec["明细"] else "")
    print(f"  场景={scene}  关键词(事项)={kw!r}")
    print(f"  甲方={rec['甲方']!r}  乙方={rec['乙方']!r}  金额={rec['含税总价']!r}")
    # 附件解析（用抽到的 kw）
    try:
        got = fc.resolve_attachments(kw, contract_path=str(p))
        for k, v in got.items():
            print(f"    [有] {k} -> {v.name}")
    except FileNotFoundError as e:
        print("    [缺件]", e)
