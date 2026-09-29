# 单据类一键自动化（Workflow）

基于 OCR 识别入库单扫描件，自动生成采购台账、验收单、付款单等单据的 Python 自动化流程。

## 功能
- **文件直接丢进 `10_输入` 根目录即可**：程序自动按文件名/扩展名/PDF 内容分拣归位（无法判定的进 `10_输入\待确认` 并提示），也可继续手动放子目录
- **PDF 版入库单免 OCR**：带文字层的 PDF 直接提取结构化数据入库（零识别误差，见 `80_工具\parse_pdf_receipt.py`）；无文字层的图片型 PDF 提示转人工，JPG/PNG 扫描件仍走 OCR
- OCR 识别入库单扫描件 → 匹配/写入数据库
- 自动生成采购台账（xls，按单据自身日期归入对应月份页）
- 生成验收单（Word + PDF）与付款单（xlsx + PDF）；**验收单明细超 17 项自动续页**，每页抬头/表头/签字栏完整（`80_工具\accept_pages.py`）
- **OCR 人工校对用侧车**：识别有错的图，在旁边放同名 `<图名>.jpg.ocr.txt`（行式明细），重跑自动采用并转存缓存，无需手算 MD5
- **付款单由数据库未付款入库单驱动**（供应商/金额/物料全取自 DB，按 status=paid 防重复出单）；从发票 PDF 自动提取供应商、金额、用途仅作**辅助依据**（用途线索+凭证归档），无发票也能出单
- 双击 `一键生成.bat` 走报告链路：旧批成品自动归档 `40_归档`、完整输出落 `70_日志`、成品暂存 `30_待打印`
- **参考件自动归档**：申购单/事项审批与发票一对一，被消费（对应发票已处理/入库单已付款）后自动移入 `40_归档`，输入目录不堆积

## 目录结构
```
10_输入   待处理单据（入库单扫描件/发票/申购单/事项审批）
20_输出   成品（台账/验收单 Word/付款单）
30_待打印 待打印 PDF
40_归档   历史文件按类型/日期分层
50_模板   只读模板
60_数据   SQLite 数据库（单据.db）
70_日志   运行日志
80_工具   通用工具模块（本地OCR、DB维护、台账重建、PDF转换等）
90_临时   工作缓存（OCR识别缓存、解包中间件、备份）
```

## 核心脚本
- `一键生成.py` — 主流程
- `db.py` — SQLite 数据访问
- `naming.py` — 统一命名规范
- `paths.py` — 路径配置（根目录按本文件位置自适应，也可用环境变量 `DJ_BASE` 覆盖）

## 80_工具 说明
- `ocr_local.ps1` — 本地识别（Windows 内置中文 OCR；人工校对优先：**侧车 `<图名>.jpg.ocr.txt`** 或 `90_临时\ocr_cache\<图片MD5>.txt`）
- `parse_pdf_receipt.py` — PDF 版入库单文字层解析（零识别误差，免 OCR）
- `accept_pages.py` — 验收单分页渲染（明细超 17 项自动续页，每页完整保留抬头/表头/签字栏）
- `payment_ledger.py` — 付款台账写入（按入库单 DB 明细逐行展开）
- `build_ocr_cache.py` — 从人工核对数据批量生成识别缓存
- `seed_db.py` — 把核对过的入库单数据写入数据库（幂等）
- `rebuild_outputs.py` — 以数据库为唯一数据源重建台账与验收单成品
- `to_pdf.py` — 验收单 docx 转 PDF（Word 组件；docx 留在 20_输出\验收单，PDF 输出到 30_待打印）
- `manage_db.py` — 查库/补录工具（list/show/add/usage/total）
- `run_report.py` — 带运行报告包裹主脚本

## 本地部署（Windows）
1. 安装依赖：`python -m pip install -r requirements.txt`（需 Python 3 + Word/Excel 桌面组件用于转 PDF）
2. 建目录骨架：`python -c "import paths; paths.ensure_dirs()"`
3. 建库：`python init_db.py`（新机器无历史 JSON 时得到空库，属正常）
4. 模板放入 `50_模板`：3 个表单模板（采购物资台帐.xls / 验收单模板.docx / 付款单模板.xlsx）随仓库分发；`付款单数据库.xlsx` 含供应商收款账户信息，不入公开仓库，需从本机拷贝
5. 扫描件直接丢 `10_输入` 根目录（自动分拣），运行：双击 `一键生成.bat`（含批归档+运行日志），或 `python 一键生成.py`（直跑）

## 依赖
Python 3 + lxml、xlrd、xlutils、xlwt、pdfplumber、pywin32、pandas、openpyxl；
OCR 使用 Windows 内置中文识别引擎（zh-Hans-CN），详见 `80_工具\ocr_local.ps1`。