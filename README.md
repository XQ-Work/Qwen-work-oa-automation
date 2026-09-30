# Qwen-work · OA 采购流程自动化

企业 OA（泛微 e-cology）采购全链路自动化：**申购 → 比质比价 → 事项审批 → 合同会签 → 付款申请 → 银行回单 → 归档**。
工具只生成草稿，提交永远人工；密码只存 Windows 凭据管理器。

> **仓库纪律：只存代码与脱敏示例。** 业务数据（台账/数据库/项目档案/交付/归档）全部在 `config.json → data_root` 指向的工作区，物理上与代码分离，永不入库。

## 五层工作区

```
{data_root}/
├── 输入/    唯一投料口（平铺）。材料名（事项名）.pdf 精确路由；
│            未按命名的按词头兜底（dzfp_/入库单→单据、回单/流水→银行回单），认不出→工作/复核/收件异常
├── 工作/    单据(入库→验收→付款) · 银行回单(扫描→核对→确认) · 复核(一切人工异常件)
├── 记录/    台账(发起台账+台账快照) · 数据库(金额链主档+文档索引) · 项目档案(2026-NNN，01~06阶段子夹)
├── 系统/    oa-automation(本仓库) · docs · 模板(01申购单~04流程表单) · 运行(浏览器登录态+state)
├── 交付/    成品PDF副本区：按事项自动合并成册、日期前缀，打印送签用
└── 归档/    办结收纳：日常采购(<2000) / 采购合同(≥2000) 自动分架 + 卡片式归档台账
根级         进度总览.xlsx —— 一行一项目：流程3列+附件10列矩阵+红黄绿灯，自动重建
```

## 模块地图

| 编号 | 功能 | OA流程/workflowid | 入口 |
|---|---|---|---|
| m01 | 申购单生成 | — | `modules/m01_sggen/make_sg.py --ledger` |
| m02 | 事项审批一条龙 | 070801 / 475401 | `modules/m02_shxiang/fill_oa.py --row N --save` |
| m03 | 比质比价报告单 | — | `modules/m03_bizhijie/bizhijie.py` |
| m03b | 供应商对比报告 | — | `modules/m03b_supplier_info/` |
| m04 | 供应商新增 | 444401 | `modules/m04_supplier/fill_supplier.py` |
| m05 | 合同会签（抽取+4必传附件） | 070302 / 476402 | `extract_contract.py` → `fill_contract.py` |
| m06 | 付款申请（发票解析+附件） | 071501 / 486901 | `intake_payment.py` → `fill_payment.py` |
| m07 | 银行回单双泳道归档 | — | `modules/m07_receipt/match_archive.py` |
| 单据 | 入库→验收单/付款单（合而不重构） | — | `danju/adapter.py run` |
| 导流 | 收料归档/攒批发草稿/交付/归档 | — | `intake_all.py run\|flush\|archive\|…` |
| 金额链 | 项目编号主键全链留痕对账 | — | `python -m oa_common.moneyline report` |
| 索引 | 投料即建档 FTS5全文+OCR+错投报警 | — | `python -m oa_common.docindex search 词` |
| 总览 | 进度总览+速查手册自动生成 | — | `python -m oa_common.dashboard` |

## 日常命令速查

```bash
python intake_all.py run            # 收料：分流+归档+建档+排队（不开浏览器）
python intake_all.py flush          # 攒批发草稿（开浏览器，失败留队可重试）
python intake_all.py status         # 队列/收件箱/异常一览
python intake_all.py verify --project 2026-005   # 单项目齐套核对
python intake_all.py archive --project 2026-005  # 办结归档（--undo 取出）
python -m oa_common.moneyline report             # 金额对账
python -m oa_common.docindex search 质保金        # 全文检索附件内容
python tools/env_doctor.py          # 换机自检（--fix 重建junction）
```

## 新机器上手

1. 拷工作区（robocopy 加 `/XJ` 防 junction 双拷）→ 2. 改 `config.json` 的 `data_root` → 3. `pip install -r requirements.txt` → 4. `python tools/env_doctor.py --fix` → 5. 首单验证。详见 `docs/换机迁移手册.md`。

## 五条铁律

1. 只到草稿，提交人工；2. 密码只进 Windows 凭据管理器；3. 数据与代码分离，迁移只改一行；4. 全系统零删除——错放就挪位，历史进回收站/归档；5. 台账结构用户说了算——用户在 `记录/台账` 改列，代码跟随。
