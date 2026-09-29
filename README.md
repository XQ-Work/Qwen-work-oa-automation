# oa-automation

企业 OA 采购流程自动化模块集（泛微 e-cology）。**仓库只存代码与模板样例，业务数据一律在 `config.json` 指向的 data_root**，凭据存 Windows 凭据管理器，绝不入库。

## 工作区四层（v4：眼睛层全中文，代码层保英文）

```
{工作区根}/
├── 输入/     唯一投料口（平铺不建子文件夹）：首选「材料名（事项名）」精确路由；
│             未按命名的按文件名特征兜底（dzfp_/入库单/申购单_→单据；回单/流水→银行回单），认不出→工作/复核/收件异常
├── 工作/     单据/(入库单→验收单→付款单流水线；80_工具与50_模板 junction 到系统层) · 银行回单/(输入·待确认·已完成)
│             复核/{收件异常, 回单核对}  ← 一切需人工过目的文件统一入口，新异常类型=加子夹
├── 记录/     台账/(5张发起台账+分支映射+命名规范) · 数据库/(主库+文档索引) · 项目档案/(2026-NNN) · 回单影像→工作/银行回单/已完成
（根级：进度总览.xlsx——dashboard 输出，config 键 board_file）
├── 系统/     oa-automation(全部代码) · docs(文档) · 模板/{01_申购单,02_比质比价,03_单据类,04_流程表单参考} · 运行(edge_profile+state)
├── 交付/     成品副本区，打印送签专用（原件仍走原链路）：申购单/ · 比质比价报告单/ · 付款单/ · 验收单/
└── 归档/     办结项目按2000限值自动分架：日常采购/(<2000) · 采购合同/(≥2000)，夹内01~06同构
              归档：python intake_all.py archive --project 2026-NNN [--undo]；总览灰显垫底；归档台账双存
              m02/m03 生成即自动多存一份；单据成品随 reflow/flush 同步；手动补扫：python intake_all.py deliver
```

命名规矩：眼睛层目录用两字业务名词、不中英混搭；程序内部标识（模块名/edge_profile）保持英文。
所有路径由 config.json 解析（含 template_sg/template_biz 模板分区键），代码零硬编码；换机器改 data_root 一处 + 重建 3 个 junction（80_工具、50_模板、回单影像）。

## 五条联通公约（所有模块共同遵守）

1. 一事一夹：`{data_root}/flow/文档整理/采购项目档案/年-序号-事项名/`（01~06 环节子夹）
2. 文件命名：`文档类型（事项名-金额：X）.pdf`，事项名是全链关联键；上传 OA 时自动去掉金额段
3. 交接面：模块间不传参数，只经 收件箱（进）/ 档案夹（归档）/ 台账状态列（记录）
4. 台账状态机：待确认 → 已保存草稿 → 已提交
5. 人工卡点：工具只保存到草稿，提交永远由人做

## 模块

| 模块 | 功能 | 入口 |
|---|---|---|
| modules/m01_shxiang | 事项审批一条龙（070801）：台账→登录→填表→附件→存草稿→归档回写 | `python modules/m01_shxiang/fill_oa.py --row N --save`；收件箱识别 `intake.py scan/add` |
| modules/m02_sggen | 申购单生成：需求台账→套模板→落收件箱 | `python modules/m02_sggen/make_sg.py --ledger` |
| modules/m03_bizhijie | 比质比价报告单生成：SQLite 报价数据→Word 模板填充→归档 03 夹 | `python modules/m03_bizhijie/bizhijie.py add-quote/gen` |
| modules/m03b_supplier_info | 供应商信息查询对比报告：企业画像入 SQLite（跨事项复用），生成横向对比 Word 表（行=查询项，列=供应商） | `python modules/m03b_supplier_info/supplier_info.py set/gen` |
| danju | 单据类自动化（并入主包，内部逻辑未动）：入库单扫描件/发票→采购台账、验收单、付款单；数据根钩子 DJ_BASE 或 config.json `danju_base` | `python danju/adapter.py status\|receipts\|invoices\|run`（run 转调原 run.py） |
| intake_all（导流器） | 收件箱分流：`材料名（事项名）` 命名→立即归档进项目夹+解析入台账→待发动作排队；flush 时攒批开浏览器统一发草稿 | `python intake_all.py scan\|run\|flush [--only m05,m06]\|status\|verify [--项目]\|reflow [--dry-run]`；自测 `run --selftest --inbox 目录` |
| docindex（文档索引） | 投料即建档：入档文件抽正文进 FTS5 全文库（记录/db/文档索引.db），错投检测（名实不符当场⚠）；二期OCR：无文字层PDF逐页渲染+图片走Windows内置OCR（按MD5缓存，run 收料尾部自动补识别） | `python -m oa_common.docindex stat\|search 关键词 [编号]\|add 文件\|rescan\|ocr [--limit N]` |
| dashboard（进度总览） | 全链监控：一行一项目，流程三列+附件十列矩阵（已上传/未上传红绿底）、下一步、红黄绿灯；第2页速查手册；run/flush 后自动重建 | `python -m oa_common.dashboard` → 记录/进度总览.xlsx |
| moneyline（金额链主档） | 项目编号为主键的金额链：预估/审批/合同金额、付款性质拆分、累计已付与未执行差额、超付告警（三表在 oa_automation.db） | `python -m oa_common.moneyline report\|show <code>`；台账历史行回填 `python tools/backfill_projcode.py` |
| oa_common | 路径/凭据/SQLite 公共库/项目夹定位(archive)（keyring 服务名 qwenwork-oa-tus-sound；库文件 {data_root}/oa_automation.db） | 内部引用 |

## 环境准备

```bash
pip install --user playwright openpyxl keyring   # Playwright 用系统 Edge(channel=msedge)
cp config.example.json config.json               # 改 data_root 指向你的数据目录
# 首次运行 fill_oa 会弹窗录入一次 OA 凭据（存 Windows 凭据管理器）
```

## 已知约束

- 泛微表单 iframe 的 JS 上下文异常：读结构用 CDP `frame.content()` 全帧普查，填值用 locator（详见 m01 注释）
- 会话 cookie 不跨浏览器重启，每次运行自动登录；域名拥堵自动切备用 IP
- 事项审批金额 ≤2000 走小额通用句式（缘由三选一：维修/日常生产/化验室日常化验），>2000 走详述句式并触发后续合同/付款链

## 路线图

- [x] m01 事项审批一条龙
- [x] m02 申购单生成
- [x] m03 比质比价报告单生成（吸收 TRAE 旧方案，输出与历史成品逐格一致）
- [ ] m04 供应商新增一条龙
- [ ] m05 合同会签一条龙
- [ ] m06 付款申请一条龙
- [ ] orchestrator 全链编排
