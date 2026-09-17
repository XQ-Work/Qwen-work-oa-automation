# -*- coding: utf-8 -*-
"""合同要素抽取器：读一份合同 PDF/docx，抽出 甲乙方、金额、明细、付款方式，
按固定模板拼「摘要」，预览或回写《合同发起台账.xlsx》。

摘要模板（用户约定）：
    根据事项审批及比质比价报告与{乙方全称}签订{新合同名称}，详情如下：
    1.此次采购{名称} {型号} 单价：{单价}元 数量：{数量}{单位} 共计：{含税总价}元。
    2.{货款支付条款原文}
新合同名称 = {关联关键词} + 采购合同/服务合同（沿用 m05 规则）。

用法：
    python extract_contract.py "合同.pdf"          # 预览（不回写）
    python extract_contract.py "合同.pdf" --commit  # 新增一行到台账，状态=待确认
"""
import argparse
import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from oa_common import paths  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

LEDGER = paths.DATA / "合同发起台账.xlsx"
SHEET = "合同台账"
# 台账列（与 m05 的 L 对齐）
COL = {2: "场景", 3: "合同名称", 4: "金额", 5: "甲方", 6: "乙方",
       7: "关联关键词", 8: "拟签日期", 9: "摘要覆盖", 10: "合同路径", 11: "状态"}


def read_text(path: Path) -> str:
    """合同全文（用于甲乙方/总价/付款条款的正则抽取）。"""
    suf = path.suffix.lower()
    if suf == ".pdf":
        import pymupdf
        d = pymupdf.open(path)
        return "\n".join(d[p].get_text() for p in range(d.page_count))
    if suf == ".docx":
        import docx
        doc = docx.Document(path)
        parts = [par.text for par in doc.paragraphs]
        for t in doc.tables:  # 甲乙方、含税总价常在表格里
            for row in t.rows:
                parts.append(" ".join(c.text for c in row.cells))
        return "\n".join(parts)
    raise SystemExit(f"暂不支持 {suf} 格式（当前支持 .pdf/.docx）")


def _clean(cell):
    return re.sub(r"\s+", "", (cell or ""))


def read_items(path: Path) -> list:
    """用版面表格识别抽价格明细行。按表头动态定位列（各合同列序不同、
    部分无单价列）。返回 [{名称,型号,单价,单位,数量}]，缺列则空串。"""
    suf = path.suffix.lower()
    grids = []
    if suf == ".pdf":
        import pymupdf
        d = pymupdf.open(path)
        for pno in range(min(3, d.page_count)):
            for tb in d[pno].find_tables().tables:
                grids.append(tb.extract())
    elif suf == ".docx":
        import docx
        for t in docx.Document(path).tables:
            grids.append([[c.text for c in row.cells] for row in t.rows])
    items = []
    for g in grids:
        if not g:
            continue
        header = g[0]
        idx = {}
        for i, h in enumerate(header):
            h = h or ""
            if "序号" in h:
                idx["seq"] = i
            elif "名称" in h:
                idx["name"] = i
            elif "型号" in h or "规格" in h:
                idx["model"] = i
            elif "单价" in h or "价格" in h:
                idx["price"] = i
            elif "单位" in h:
                idx["unit"] = i
            elif "数量" in h:
                idx["qty"] = i
        if "name" not in idx:  # 不是明细表
            continue
        for row in g[1:]:
            name = _clean(row[idx["name"]] if idx["name"] < len(row) else "")
            if "总价" in name or "含税" in name or not name:
                continue
            def get(k):
                j = idx.get(k)
                return _clean(row[j]) if j is not None and j < len(row) else ""
            items.append({"名称": name, "型号": get("model"), "单价": get("price"),
                          "单位": get("unit"), "数量": get("qty")})
        if items:
            break
    return items


def _num(s):
    s = re.sub(r"[^\d.]", "", s or "")
    return s


def _fmt_price(s):
    """4195.00 -> 4195，8390.00 保留两位用于总价。"""
    v = float(s)
    return str(int(v)) if v == int(v) else f"{v:.2f}"


def parse(text: str, items: list = None) -> dict:
    nl = [l for l in (x.strip() for x in text.splitlines()) if l]
    rec = {"甲方": "", "乙方": "", "含税总价": "", "明细": items or [], "付款": ""}

    # 甲乙方：找含“（买方/卖方）”的行，取其后候选，只认“像公司名”的，
    # 挡掉日期/省市地点（docx 表格拼接常把甲方：后面的日期误当名称）
    def _is_org(s):
        s = s.strip()
        if not (2 <= len(s) <= 30) or "：" in s or ":" in s:
            return False
        if re.search(r"[【】\d年月日省市县路号]", s):
            # 地点/日期特征；公司名极少含【】或裸“省/市/路”
            if not any(k in s for k in ("公司", "有限")):
                return False
        return any(k in s for k in
                   ("公司", "有限", "厂", "中心", "经营部", "商行", "门市",
                    "合作社", "集团", "商店", "站", "所", "部"))

    def party(keys):
        for i, l in enumerate(nl):
            if any(k in l for k in keys):
                tail = l.split("：")[-1].split(":")[-1].strip()
                cand = [tail] + nl[i + 1:i + 3]
                for c in cand:
                    if _is_org(c):
                        return c
        return ""
    rec["甲方"] = party(["甲方（买方", "甲方(买方", "甲方（购买", "甲方：", "甲方:"])
    rec["乙方"] = party(["乙方（卖方", "乙方(卖方", "乙方（供货", "乙方：", "乙方:"])

    # 含税总价：锚定“含税总价为”，跳过正文“合同含税总价已经涵盖…”模板句；
    # “不含税总价”是超串，用 (?<!不) 挡掉
    m = re.search(r"(?<!不)含税总价为[^\d]{0,4}([\d][\d,]*\.?\d*)", text)
    if m:
        v = float(_num(m.group(1)))
        rec["含税总价"] = f"{v:.2f}"  # 总价一律保留两位，匹配“共计：8390.00”

    # 付款方式：从“货款支付”起到第一个句号
    flat = re.sub(r"\s+", "", text)
    m = re.search(r"(货款支付[：:].*?。|付款方式[：:].*?。|支付方[。]?)", flat)
    if m:
        rec["付款"] = m.group(1)
    return rec


def build_summary(rec: dict, scene: str, keyword: str) -> str:
    cname = keyword + ("服务合同" if "服务" in scene else "采购合同")
    title = f"根据事项审批及比质比价报告与{rec['乙方']}签订{cname}，详情如下："
    items = rec["明细"]
    total = rec["含税总价"]

    def _seg(it):
        model = f" {it['型号']}" if it["型号"] else ""
        price = (f" 单价：{_fmt_price(_num(it['单价']))}元" if it["单价"] else "")
        qty = f" 数量：{it['数量']}{it['单位']}" if it["数量"] else ""
        return f"{it['名称']}{model}{price}{qty}"

    if len(items) == 1:
        p1 = f"1.此次采购{_seg(items[0])} 共计：{total}元。"
    elif len(items) > 1:
        seg = "、".join(_seg(x) for x in items)
        p1 = f"1.此次采购{seg}，共计：{total}元。"
    else:
        p1 = f"1.共计：{total}元。"
    p2 = "2." + (rec["付款"] or "货款支付按合同约定执行。")
    return f"{title}\n{p1}\n{p2}"


def detect_scene(text: str, stem: str = "") -> str:
    if "服务合同" in text[:400] or "服务合同" in stem:
        return "服务"
    return "货物"


# 文件名括号里常混入的状态词，抽关键词时剔除
_STATUS = ["已用印", "用印", "盖章", "定稿", "终稿", "初稿", "正本", "副本",
           "扫描件", "已签", "未签"]


def keyword_from_name(stem: str) -> str:
    m = re.search(r"[（(]([^（）()]+)[）)]", stem)
    raw = m.group(1).strip() if m else ""
    for s in _STATUS:
        raw = raw.replace(s, "")
    return raw.strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--commit", action="store_true", help="新增一行回写台账")
    args = ap.parse_args()

    path = Path(args.file)
    if not path.exists():
        raise SystemExit(f"文件不存在：{path}")
    text = read_text(path)
    items = read_items(path)
    rec = parse(text, items)
    scene = detect_scene(text, path.stem)
    keyword = keyword_from_name(path.stem) or \
        (rec["明细"][0]["名称"] if rec["明细"] else "")
    summary = build_summary(rec, scene, keyword)

    # 扫描件守卫：文字层几乎为空且啥都没抽到 -> 拒绝，提示走视觉兜底
    if len(text.strip()) < 150 and not rec["乙方"] and not rec["含税总价"]:
        print(f"[!] 疑似扫描件（无文字层），本合同暂不支持自动抽取，"
              f"请人工填台账或等视觉兜底版本：{path.name}")
        return

    print("=" * 60)
    print(f"文件：{path.name}")
    print(f"场景：{scene}   关联关键词：{keyword}")
    print(f"甲方：{rec['甲方']}")
    print(f"乙方：{rec['乙方']}")
    print(f"金额（含税总价）：{rec['含税总价']}")
    print(f"明细：{rec['明细']}")
    print("-" * 60)
    print("生成的【摘要】：")
    print(summary)
    print("=" * 60)

    if not args.commit:
        print("（预览模式，未回写。加 --commit 才写入台账）")
        return

    wb = load_workbook(LEDGER)
    ws = wb[SHEET]
    r = ws.max_row + 1
    vals = {2: scene, 3: keyword + ("服务合同" if scene == "服务" else "采购合同"),
            4: rec["含税总价"], 5: rec["甲方"], 6: rec["乙方"], 7: keyword,
            9: summary, 10: str(path), 11: "待确认"}
    for c, v in vals.items():
        if v:
            ws.cell(row=r, column=c, value=v)
    wb.save(LEDGER)
    print(f"[ok] 已新增台账第{r}行（状态=待确认，请核对后跑 fill_contract）")


if __name__ == "__main__":
    main()
