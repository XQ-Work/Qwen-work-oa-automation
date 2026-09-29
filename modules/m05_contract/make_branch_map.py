# -*- coding: utf-8 -*-
"""生成《合同分支映射表.xlsx》：场景 -> 级联下拉选择映射，全列带实测选项下拉
用户填完后 m05 填单脚本按此表自动决策下拉路径"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

import sys as _s; _s.path.insert(0, r"D:\自动备份\Qwen work\自动化（OA流程）\系统\oa-automation"); from oa_common import paths as _p
OUT = str(_p.BRANCH_MAP)

# ===== 实测选项（probe_cascade 两轮）=====
OPTS = {
    "①事项类型": ["货物类采购事项", "服务类采购事项"],
    "水务板块事项审批": [
        "项目公司非长期固定采购生产物资", "项目公司设备设施购置等采购类合同",
        "项目公司工程施工类合同", "项目公司技术服务类合同（可研、环评、勘测、设计、监理、代理等）",
        "项目公司长期采购合同（框架协议）", "项目公司工程费率合同",
        "项目公司自来水报装项目合同", "项目公司日常生产经营用水、用电等国家固定格式合同",
        "项目公司人力资源类事项-合同", "项目公司行政办公类事项-合同",
        "项目公司法务类事项-合同", "项目公司财务类事项-合同",
        "项目公司工程其他费用类-合同"],
    "水务项目状态": ["在建", "运营", "轻资产项目"],
    "水务项目种类": ["污水", "净水", "轻资产项目"],
    "项目状态": ["2.1在建", "2.2运营", "2.3日常经营",
            "2.4特殊业务合同审批（贸 易、融资租赁）", "2.5特殊事项(咨询内控后使用)",
            "2.6特殊事项（投资顾问协议）-水务", "2.7融资合同"],
    "项目类型": ["水务-建筑工程", "水务-安装工程", "水务-设备采购及安装", "水务-服务类",
            "水务-生产物资", "水务-技改", "水务-日常经营",
            "水务-特殊业务合同审批（贸 易、融资租赁）", "水务-特殊事项（投资顾问协议）",
            "建筑工程", "建筑工程-水务", "安装工程", "安装工程-水务",
            "设备采购及安装", "设备采购及安装-水务", "服务类", "服务类-水务", "水务-融资合同"],
    "合同内容": ["工程类", "设备采购及安装", "日常运营及其他", "大宗物资",
            "贸易或融资租赁（特殊业务）", "特殊事项(权限指引内的特殊事项)"],
    "合同所属类型": ["销售", "采购", "设备", "施工", "安装", "租赁", "保密", "股权", "维修"],
    "合同类型": ["工程承包", "建设工程（分包）", "设备采购（外协加工）", "工程设计/咨询外包",
            "对外承揽工程设计/技术咨询", "融资类", "行政/人力资源", "招标文件", "销售类"],
    "是否为采购类合同": ["是", "否"],
    "人员所属": ["项目公司在建项目部/建设管理部项目部", "其他人员"],
}

HEADERS = ["场景名*", "①事项类型", "水务板块事项审批", "水务项目状态", "水务项目种类",
           "项目状态", "项目类型", "合同内容", "合同所属类型", "合同类型",
           "是否为采购类合同", "人员所属", "历史参照", "备注"]

PRE = [
    ["货物-生产物资/设备", "货物类采购事项", "项目公司非长期固定采购生产物资", "运营", "污水",
     "2.2运营", "水务-生产物资", "设备采购及安装", "采购", "", "是", "其他人员",
     "变频器/潜污泵/滤布/机械油/化验室药剂", "合同类型列请补"],
    ["服务-维修/外包", "服务类采购事项", "", "运营", "污水",
     "2.3日常经营", "水务-日常经营", "日常运营及其他", "维修", "", "是", "其他人员",
     "细格栅维修/垃圾清运", "水务板块事项审批列请选"],
    ["施工-工程类", "", "", "", "",
     "", "", "工程类", "施工", "", "", "",
     "污水池堵漏", "整行请按实际补"],
]

YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
ws = wb.active
ws.title = "分支映射"
gd = wb.create_sheet("选项库")

# 选项库 sheet（下拉引用源）
col = 1
ranges = {}
for name, opts in OPTS.items():
    c = gd.cell(row=1, column=col, value=name)
    c.font = Font(name="微软雅黑", size=10, bold=True)
    for i, o in enumerate(opts, 2):
        gd.cell(row=i, column=col, value=o)
    letter = get_column_letter(col)
    ranges[name] = f"选项库!${letter}$2:${letter}${len(opts)+1}"
    gd.column_dimensions[letter].width = 30
    col += 1
gd.sheet_state = "hidden"

for ci, h in enumerate(HEADERS, 1):
    c = ws.cell(row=1, column=ci, value=h)
    c.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
    ws.column_dimensions[get_column_letter(ci)].width = 24
ws.row_dimensions[1].height = 28

for r, row in enumerate(PRE, 2):
    for ci, v in enumerate(row, 1):
        c = ws.cell(row=r, column=ci, value=v)
        c.font = Font(name="微软雅黑", size=10, color="808080", italic=True)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        c.border = BORDER

for r in range(5, 25):
    for ci in range(1, len(HEADERS) + 1):
        c = ws.cell(row=r, column=ci)
        c.border = BORDER
        c.font = Font(name="微软雅黑", size=10)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        if ci <= 12:
            c.fill = YELLOW

# 挂下拉（列2..12 对应 OPTS 顺序）
col_map = {2: "①事项类型", 3: "水务板块事项审批", 4: "水务项目状态", 5: "水务项目种类",
           6: "项目状态", 7: "项目类型", 8: "合同内容", 9: "合同所属类型",
           10: "合同类型", 11: "是否为采购类合同", 12: "人员所属"}
for ci, key in col_map.items():
    dv = DataValidation(type="list", formula1="=" + ranges[key], allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{get_column_letter(ci)}2:{get_column_letter(ci)}24")

ws.freeze_panes = "B2"

tip = wb.create_sheet("怎么填")
tip.column_dimensions["A"].width = 100
lines = [
    ("合同分支映射表——填法", True),
    ("", False),
    ("一行 = 一种签约场景。以后合同会签发起时，脚本按『①事项类型+场景』查这张表，", True),
    ("自动完成 12 个级联下拉的选择，不再需要人工判断。", True),
    ("", False),
    ("■ 场景名：起个你自己认得的名字（如 货物-生产物资、服务-检测、施工-堵漏）", False),
    ("■ 灰色三行是预填参考（来自历史过审实例），确认后改成正式行或直接在其下重写", False),
    ("■ 下拉选项全部是系统实测值，直接选即可；『项目类型』选项随『项目状态』联动，", False),
    ("   若某组合实际不存在，脚本填单时会报错反馈，再调整", False),
    ("■ 拿不准的列留空，脚本会停在草稿由你补", False),
    ("", False),
    ("注：『水务板块事项审批』的可选分支可能随前置选择（水务项目状态等）变化，", False),
    ("本表选项按实测快照生成；若你发现某场景下系统里没有对应分支，告诉我重测。", False),
]
for i, (t, b) in enumerate(lines, 1):
    c = tip.cell(row=i, column=1, value=t)
    c.font = Font(name="微软雅黑", size=11 if b else 10, bold=b)
    c.alignment = Alignment(wrap_text=True, vertical="top")

wb.save(OUT)
print("saved:", OUT)
