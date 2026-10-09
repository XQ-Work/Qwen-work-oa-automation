# -*- coding: utf-8 -*-
"""重建《归档说明.docx》：与卡片式归档台账（用户定稿）同步。跑完即删本脚本。"""
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()
st = doc.styles["Normal"]
st.font.name = "微软雅黑"; st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

def run(txt, size=11, bold=False, color=None, center=False):
    p = doc.add_paragraph(); r = p.add_run(txt)
    r.bold = bold; r.font.size = Pt(size)
    if color: r.font.color.rgb = RGBColor.from_string(color)
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def _shd(color):
    e = OxmlElement("w:shd"); e.set(qn("w:fill"), color); return e

def table(rows):
    t = doc.add_table(rows=len(rows), cols=len(rows[0])); t.style = "Table Grid"
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.cell(i, j); c.text = ""
            r = c.paragraphs[0].add_run(v); r.font.size = Pt(10.5); r.bold = (i == 0)
            if i == 0:
                r.font.color.rgb = RGBColor.from_string("FFFFFF")
                c._tc.get_or_add_tcPr().append(_shd("4472C4"))

run("归 档 说 明", 20, True, "1F3864", center=True)
run("归档分“双轨”：跑任务的轨（记录/项目档案）存过程全量，办结后原地不动；", 12)
run("你看档案的轨（归档）在办结时组装一份定稿册，按 2000 元限值自动分两架：", 12)
table([("架", "放什么", "例子"),
       ("归档/日常采购", "申购金额低于 2000 元（不走合同）", "车辆保险、小额配件"),
       ("归档/采购合同", "申购金额 2000 元及以上（走过合同会签）", "潜污泵采购、化验室药剂")])
run("")
run("定稿册怎么排（一项目一夹，夹名带〔办结日期〕）", 14, True, "4472C4")
run("主件按业务顺序编号 01、02…（真实/生效凭证）：事项审批→比质比价→报价单→采购合同(已用印)→流程会签→发票→银行回单。")
run("验收单、付款单是系统按台账生成的送签物料、非真实凭证，单列进册内“生成物料（非真实凭证）”子夹。")
run("缺哪类就不占号，自动连续。整套台账 + 进度总览同时刷一份快照进 归档/台账快照。")
run("")
run("归档时自动填好一张“项目卡片”（记在 归档台账.xlsx，两个地方各存一份）", 14, True, "4472C4")
table([("表", "列", "怎么填"),
       ("采购合同", "编号/创建时间/事项名称/事项类别/供应商/合同金额/税率/签订时间/应付/已付/未付/付款时间/其他",
        "系统自动：创建时间=事项审批发起日；签订时间=归档当天；税率从合同正文抓取；已付/未付/付款时间按付款台账+回单核算（付款时间可多次，顿号分隔）。个别抓不到的（如盖章件上的补充信息）留空手填。"),
       ("日常采购", "编号/创建时间/审批状态/审批日期/入库单单号/付款/供应商/其他",
        "系统自动：审批状态与日期取事项审批台账；付款=金额（一单一次付，没有二次付款）；供应商和入库单单号按金额反查单据库自动带出（同名多结果不猜、留空）。")])
run("")
run("三条约定", 14, True, "4472C4")
run("1.  什么时候归：进度总览该项目流程三列全“已归档”、附件无红格、钱账核平，三条齐才归。")
run("2.  怎么归：跟助理说“归档 2026-005”即可（自动选架、组定稿册、写卡片、刷快照、总览变灰）；放错了说“取出 2026-005”，定稿册移进回收、卡片行同步删，记录/项目档案始终不受影响。")
run("3.  查历史：直接开 归档/日常采购 或 采购合同，一个项目一个定稿册；要看全部过程件去 记录/项目档案。")
run("")
run("会不会误删：全系统没有删除功能，归档只是复制组册（记录原样留着），取出也只是把册挪进回收；台账想改口径，直接在 归档台账.xlsx 上改，系统跟着学。", 10, color="808080")
import pathlib as _pl
ROOT = _pl.Path(__file__).resolve().parents[3]
doc.save(str(ROOT / "归档/归档说明.docx"))
print("归档说明.docx 已重建")
