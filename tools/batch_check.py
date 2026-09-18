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
PROJECTS = [
    "2026-002-滤布采购", "2026-003-潜污泵采购", "2026-004-污水池堵漏",
    "2026-005-发电机排涝采购", "2026-007-细格栅维修", "2026-008-维修用品采购",
    "2026-009-垃圾清运服务", "2026-010-机械油及电气配件采购",
    "2026-011-化验室药剂及器具",
]


def find_contract(proj):
    """在 04_采购合同 里找文字版合同 PDF(排除 已用印/盖章/扫描)。"""
    d = ARCH / proj / "04_采购合同"
    if not d.exists():
        return None
    cands = [f for f in d.glob("*.pdf")
             if not any(s in f.name for s in ("已用印", "盖章", "扫描"))]
    return cands[0] if cands else None


for proj in PROJECTS:
    print("=" * 70)
    print("项目:", proj)
    p = find_contract(proj)
    if not p:
        print("  (04_采购合同 下无文字版合同 PDF，跳过)")
        continue
    text = ex.read_text(p)
    items = ex.read_items(p)
    rec = ex.parse(text, items)
    scene = ex.detect_scene(text, p.stem)
    kw = ex.keyword_from_path(p) or (rec["明细"][0]["名称"] if rec["明细"] else "")
    print(f"  场景={scene}  关键词(事项)={kw!r}")
    print(f"  甲方={rec['甲方']!r}  乙方={rec['乙方']!r}  金额={rec['含税总价']!r}")
    # 附件解析（用抽到的 kw）
    try:
        got = fc.resolve_attachments(kw, contract_path=str(p))
        for k, v in got.items():
            print(f"    [有] {k} -> " + " + ".join(x.name for x in v))
    except FileNotFoundError as e:
        print("    [缺件]", e)
