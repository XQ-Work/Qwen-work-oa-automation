# -*- coding: utf-8 -*-
"""为每张入库单扫描件生成识别结果缓存（人工校对后的行式明细）。

背景：本机 Windows 内置 OCR 对表格类扫描件按“列”输出文本，主流程的明细解析器
需要的是“行式”布局。因此改为逐张核对原图后写入标准格式缓存；主流程读到缓存即
跳过自动识别，直接进入单号匹配与后续出单环节。

缓存键：图片文件内容的 MD5，落在 paths.OCR_CACHE_DIR/<md5>.txt。
幂等：默认跳过已存在的缓存；--force 覆盖重建。
"""
import hashlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import paths

# 字段顺序须与主流程 parse_items_from_ocr 的期望一致：
# code / name / spec / unit / qty / price / amount
RECEIPTS = [
    {
        "file": "111 (1).jpg",
        "receipt_id": "CGSH202608130007",
        "date": "2026-08-03",
        "supplier": "沙市区顺达轴承五金交化经营部",
        "items": [
            {"code": "020200001965", "name": "轴承", "spec": "NU310", "unit": "套",
             "qty": 1.00, "price": 210.0000, "amount": 210},
            {"code": "1020000499", "name": "轴承", "spec": "3316", "unit": "套",
             "qty": 1.00, "price": 970.0000, "amount": 970},
            {"code": "1020001933", "name": "轴承", "spec": "6309", "unit": "套",
             "qty": 2.00, "price": 94.0000, "amount": 188},
            {"code": "020400046868", "name": "油封", "spec": "30*52*7", "unit": "个",
             "qty": 10.00, "price": 3.0000, "amount": 30},
        ],
        "total": 1398.00,
    },
    {
        "file": "111 (2).jpg",
        "receipt_id": "CGSH202608130008",
        "date": "2026-08-03",
        "supplier": "沙市区凯胜管道水暖配件经营部",
        "items": [
            {"code": "020400060813", "name": "单向阀", "spec": "四氟", "unit": "个",
             "qty": 6.00, "price": 360.0000, "amount": 2160},
            {"code": "10701000831", "name": "玻璃钢格栅",
             "spec": "盖板防滑型，每块玻璃钢格栅配相应自攻", "unit": "平方米",
             "qty": 25.00, "price": 156.0000, "amount": 3900},
            {"code": "020400060813", "name": "单向阀", "spec": "四氟", "unit": "个",
             "qty": 6.00, "price": 260.0000, "amount": 1560},
        ],
        "total": 7620.00,
    },
    {
        "file": "111 (3).jpg",
        "receipt_id": "CGSH202608260003",
        "date": "2026-08-26",
        "supplier": "示例五金经营部（脱敏样例）",
        "items": [
            {"code": "030300001634", "name": "电缆线", "spec": "4*2.5", "unit": "卷",
             "qty": 1.00, "price": 1200.0000, "amount": 1200},
            {"code": "020400038491", "name": "电缆", "spec": "16mm2", "unit": "米",
             "qty": 30.00, "price": 17.0000, "amount": 510},
        ],
        "total": 1710.00,
    },
    {
        "file": "111 (4).jpg",
        "receipt_id": "CGSH202607210006",
        "date": "2026-07-21",
        "supplier": "示例五金经营部（脱敏样例）",
        "items": [
            {"code": "02040012226", "name": "电机风扇", "spec": "Y100", "unit": "个",
             "qty": 4.00, "price": 150.0000, "amount": 600},
            {"code": "1104001819", "name": "电机风扇", "spec": "Y90", "unit": "个",
             "qty": 4.00, "price": 125.0000, "amount": 500},
            {"code": "140000007091", "name": "变频风机", "spec": "GB-80", "unit": "个",
             "qty": 4.00, "price": 115.0000, "amount": 460},
            {"code": "020400051345", "name": "轴流风扇", "spec": "200*200", "unit": "个",
             "qty": 2.00, "price": 75.0000, "amount": 150},
            {"code": "020400035559", "name": "轴流风机", "spec": "200FZY-D", "unit": "台",
             "qty": 2.00, "price": 65.0000, "amount": 130},
        ],
        "total": 1840.00,
    },
    {
        "file": "111 (5).jpg",
        "receipt_id": "CGSH202608130001",
        "date": "2026-08-13",
        "supplier": "一次性供应商",
        "items": [
            {"code": "040500006723", "name": "混凝土", "spec": "", "unit": "方",
             "qty": 3.00, "price": 500.0000, "amount": 1500},
        ],
        "total": 1500.00,
    },
    {
        "file": "111 (6).jpg",
        "receipt_id": "CGSH202608130002",
        "date": "2026-08-13",
        "supplier": "沙市区顺达轴承五金交化经营部",
        "items": [
            {"code": "020200002350", "name": "轴承座", "spec": "SN210", "unit": "套",
             "qty": 3.00, "price": 48.0000, "amount": 144},
            {"code": "020400039350", "name": "轴承", "spec": "22210", "unit": "套",
             "qty": 4.00, "price": 115.0000, "amount": 460},
            {"code": "0202001854", "name": "轴承", "spec": "7312AC", "unit": "套",
             "qty": 1.00, "price": 247.0000, "amount": 247},
            {"code": "0202001853", "name": "轴承", "spec": "7310AC", "unit": "套",
             "qty": 3.00, "price": 165.0000, "amount": 495},
            {"code": "0202001644", "name": "轴承", "spec": "6209", "unit": "套",
             "qty": 2.00, "price": 28.0000, "amount": 56},
            {"code": "1020001933", "name": "轴承", "spec": "6309", "unit": "套",
             "qty": 5.00, "price": 40.0000, "amount": 200},
        ],
        "total": 1602.00,
    },
    {
        "file": "111 (7).jpg",
        "receipt_id": "CGSH202608130003",
        "date": "2026-08-13",
        "supplier": "沙市区凯胜管道水暖配件经营部",
        "items": [
            {"code": "020400051338", "name": "机械密封", "spec": "WB2-40", "unit": "个",
             "qty": 1.00, "price": 210.0000, "amount": 210},
            {"code": "020400013662", "name": "镀锌角铁", "spec": "40", "unit": "米",
             "qty": 12.00, "price": 15.8333, "amount": 190},
            {"code": "010700000244", "name": "槽钢", "spec": "80#", "unit": "根",
             "qty": 6.00, "price": 245.0000, "amount": 1470},
        ],
        "total": 1870.00,
    },
    {
        "file": "111 (8).jpg",
        "receipt_id": "CGSH202608130004",
        "date": "2026-08-13",
        "supplier": "荆州区乾顺五金电器经营部",
        "items": [
            {"code": "020400039726", "name": "手拉葫芦", "spec": "2T 6m", "unit": "台",
             "qty": 2.00, "price": 320.0000, "amount": 640},
            {"code": "040500006717", "name": "灭火器箱", "spec": "4*2", "unit": "个",
             "qty": 2.00, "price": 50.0000, "amount": 100},
            {"code": "1020000561", "name": "机械密封", "spec": "109-70", "unit": "套",
             "qty": 2.00, "price": 195.0000, "amount": 390},
        ],
        "total": 1130.00,
    },
    {
        "file": "111 (9).jpg",
        "receipt_id": "CGSH202608130005",
        "date": "2026-08-13",
        "supplier": "示例五金经营部（脱敏样例）",
        "items": [
            {"code": "0499000180", "name": "抄网", "spec": "", "unit": "个",
             "qty": 20.00, "price": 15.0000, "amount": 300},
            {"code": "030500004146", "name": "电控换向阀", "spec": "4V210-08/欧蕾凯", "unit": "支",
             "qty": 6.00, "price": 49.0000, "amount": 294},
            {"code": "1020001937", "name": "玻璃胶", "spec": "", "unit": "瓶",
             "qty": 2.00, "price": 12.0000, "amount": 24},
            {"code": "0204001254", "name": "挂锁", "spec": "50", "unit": "把",
             "qty": 1.00, "price": 20.0000, "amount": 20},
            {"code": "1020001272", "name": "彩条布", "spec": "", "unit": "块",
             "qty": 1.00, "price": 75.0000, "amount": 75},
            {"code": "0204000278", "name": "下水管", "spec": "", "unit": "根",
             "qty": 1.00, "price": 10.0000, "amount": 10},
            {"code": "020400016375", "name": "水龙头", "spec": "20", "unit": "支",
             "qty": 1.00, "price": 35.0000, "amount": 35},
            {"code": "039900004993", "name": "5孔插座", "spec": "10A", "unit": "个",
             "qty": 1.00, "price": 8.0000, "amount": 8},
            {"code": "0101003408", "name": "螺丝", "spec": "16*120", "unit": "个",
             "qty": 8.00, "price": 3.0000, "amount": 24},
        ],
        "total": 790.00,
    },
    {
        "file": "111 (10).jpg",
        "receipt_id": "CGSH202608130006",
        "date": "2026-08-13",
        "supplier": "湖北新虎环保科技有限公司",
        "items": [
            {"code": "020400013315", "name": "沙子", "spec": "（吨）", "unit": "吨",
             "qty": 8.00, "price": 52.0000, "amount": 416},
            {"code": "040500006678", "name": "石子", "spec": "", "unit": "吨",
             "qty": 7.30, "price": 52.3288, "amount": 382},
        ],
        "total": 798.00,
    },
]


def fmt(v):
    """数值转文本：整数不带小数点，小数保留票面精度。"""
    if v is None:
        return ""
    f = float(v)
    if f == int(f):
        return str(int(f))
    return ("%f" % f).rstrip("0").rstrip(".")


def build_lines(r):
    """按主流程解析器期望的行式布局输出。"""
    lines = [
        "采购入库单",
        "单位名称：荆州浦华荆清水务有限公司",
        "日期：%s" % r["date"],
        "入库单号：%s" % r["receipt_id"],
        "供应商名称",
        r["supplier"],
        "物料编号",
    ]
    for it in r["items"]:
        lines.append(it["code"])
        lines.append(it["name"])
        if it.get("spec"):
            lines.append(it["spec"])
        lines.append(it["unit"])
        lines.append(fmt(it["qty"]))
        lines.append(fmt(it["price"]))
        lines.append(fmt(it["amount"]))
    lines.append("合计：")
    lines.append(fmt(r["total"]))
    return lines


def main():
    force = "--force" in sys.argv
    os.makedirs(paths.OCR_CACHE_DIR, exist_ok=True)
    written, skipped, missing = 0, 0, []
    for r in RECEIPTS:
        src = os.path.join(paths.SCAN_DIR, r["file"])
        if not os.path.exists(src):
            missing.append(r["file"])
            continue
        with open(src, "rb") as f:
            key = hashlib.md5(f.read()).hexdigest()
        dst = os.path.join(paths.OCR_CACHE_DIR, key + ".txt")
        if os.path.exists(dst) and not force:
            print("跳过（已有缓存）: %s -> %s" % (r["file"], key[:8]))
            skipped += 1
            continue
        with open(dst, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(build_lines(r)) + "\n")
        print("写入缓存: %s -> %s.txt (%d 项)" % (r["file"], key[:8], len(r["items"])))
        written += 1
    print("\n完成：新建 %d，跳过 %d，缺图 %d" % (written, skipped, len(missing)))
    if missing:
        print("未找到的图片：" + ", ".join(missing))


if __name__ == "__main__":
    main()
