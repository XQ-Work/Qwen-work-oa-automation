// OA 自动化 · 现状与说明（2026-09-29 五层版）  用法: node tools/gen_status_doc.js "../docs/OA自动化_现状与说明.docx"
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow,
  TableCell, WidthType, ShadingType,
} = require("docx");

const FONT = "Microsoft YaHei";
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, ...o });
const H = (t, lvl) => new Paragraph({ heading: lvl, children: [run(t, { bold: true })], spacing: { before: 200, after: 100 } });
const P = (t, o = {}) => new Paragraph({ children: [run(t, o.r || {})], spacing: { after: o.gap ? 200 : 80 }, ...(o.p || {}) });
const bullet = (t) => new Paragraph({ children: [run(t)], bullet: { level: 0 }, spacing: { after: 40 } });
function cell(text, { bold = false, w = 2200, fill = null } = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: fill ? { type: ShadingType.CLEAR, fill } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [run(text, { bold })], spacing: { after: 0 } })],
  });
}
function table(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, { bold: true, w: widths[i], fill: "DDEBF7" })) }),
      ...rows.map((r) => new TableRow({ children: r.map((c, i) => cell(c, { w: widths[i] })) })),
    ],
  });
}

const children = [
  new Paragraph({ heading: HeadingLevel.TITLE, children: [run("OA 流程自动化 · 现状与说明", { bold: true })], spacing: { after: 80 } }),
  P("2026-09-29 · 数据清空后从零运行 · 工作区五层：输入 / 工作 / 记录 / 系统 / 交付 / 归档", { r: { color: "808080" } }),

  H("一、总览", HeadingLevel.HEADING_1),
  table(
    ["部件", "职责", "入口"],
    [
      ["M01 事项审批", "070801/475401 填单存草稿+归档回写", "modules/m01_shxiang/fill_oa.py --row N --save"],
      ["M02 申购单", "需求台账→申购单（自动落输入口+交付区）", "modules/m02_sggen/make_sg.py --ledger"],
      ["M03 比质比价", "报价SQLite→报告单生成（含交付副本）", "modules/m03_bizhijie/bizhijie.py"],
      ["M03b 供应商对比", "企业画像库+横向对比报告", "modules/m03b_supplier_info"],
      ["M04 供应商新增", "444401 建档填单", "modules/m04_supplier/fill_supplier.py"],
      ["M05 合同会签", "070302/476402 抽取要素→填单→3行4必传附件", "extract_contract.py / fill_contract.py"],
      ["M06 付款申请", "071501/486901 发票解析→填单→附件", "intake_payment.py / fill_payment.py"],
      ["M07 回单归档", "OCR台账→双泳道匹配→回单进05夹", "modules/m07_receipt"],
      ["单据流水线", "入库单→采购台账→验收单/付款单（合而不重构并入）", "danju/adapter.py run"],
      ["导流器", "投料分流归档+排队，攒批发草稿", "intake_all.py（8 子命令）"],
      ["金额链", "项目编号主键，全链金额留痕+对账", "python -m oa_common.moneyline report"],
      ["文档索引", "投料即建档，FTS5全文+OCR+错投报警", "python -m oa_common.docindex"],
      ["进度总览", "流程3列+附件10列矩阵+速查页，自动重建", "python -m oa_common.dashboard"],
      ["交付区", "成品转PDF、按事项合并一册、日期前缀", "自动生成；intake_all.py deliver"],
      ["归档区", "2000分架+卡片台账，archive/--undo", "intake_all.py archive"],
    ],
    [1800, 4200, 3000]
  ),

  H("二、使用动线（用户只有三个动作）", HeadingLevel.HEADING_1),
  bullet("投料：文件按「材料名（事项名）.pdf」丢进 输入/（回单/入库单/电子发票免改名）"),
  bullet("收料：让助理跑 intake_all.py run —— 分流、归位、建档、排队，不开浏览器"),
  bullet("发草稿：intake_all.py flush —— 攒批开浏览器逐个存草稿，requestid 自动回写；提交永远人工"),
  P("其余全自动：进度总览/交付区/金额链/归档卡片随 run、flush、archive 同步刷新。", { r: { color: "808080" } }),

  H("三、进度总览怎么读", HeadingLevel.HEADING_1),
  bullet("流程三列（事项审批/合同会签/付款流程）：未完成→已完成→已归档"),
  bullet("附件十列矩阵：列头即应交清单，绿=已上传（档案有实物）、红=未上传、—=还没到这个环节"),
  bullet("灯：绿=不用管，黄=等你投料或等回单，红=出错必须人工（超付/草稿失败/名实存疑）"),
  bullet("第二个工作表「速查」：命名规范、必传件、范例、口径全在里面"),

  H("四、台账与数据", HeadingLevel.HEADING_1),
  bullet("发起台账5张在 记录/台账/（现为清空后表头版），项目编号列贯穿，机器读写"),
  bullet("台账双存规矩：主份在 记录/台账/，副本随功能夹（交付台账在交付/，归档台账在归档/）"),
  bullet("归档台账=卡片式双sheet（采购合同13列 / 日常采购8列），归档自动填行，税率/供应商自动抓"),
  bullet("数据库：主档(金额链)+文档索引(全文/OCR缓存) 在 记录/数据库/，永不入 git"),

  H("五、挂起与待办", HeadingLevel.HEADING_1),
  bullet("挂起：单据三接线(入库列/付款关联入库单/追付款单)、OA端回读、尾差全文断案、服务类文本一致性、多用户配置化"),
  bullet("待办：推送GitHub私有仓库(等地址)、OA服务器7个测试草稿手动删、新电脑实跑迁移包、首单实盘全流程"),

  H("六、铁律", HeadingLevel.HEADING_1),
  bullet("只到草稿，提交人工；密码只进 Windows 凭据管理器"),
  bullet("数据代码分离：config.json 定位一切，迁移只改 data_root + 重建 junction"),
  bullet("全系统无删除功能：错放就挪位，历史进回收站或归档"),
  bullet("改台账结构=用户说了算：系统在 记录/台账 上直接改，系统跟着学（代码有回归测试保底）"),
];

const doc = new Document({ sections: [{ properties: {}, children }] });
const out = process.argv[2] || require("path").join(__dirname, "../../docs/OA自动化_现状与说明.docx");
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(out, b); console.log("written:", out, b.length, "bytes"); });
