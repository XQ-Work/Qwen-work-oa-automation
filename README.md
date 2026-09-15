# oa-automation

企业 OA 采购流程自动化模块集（泛微 e-cology）。**仓库只存代码与模板样例，业务数据一律在 `config.json` 指向的 data_root**，凭据存 Windows 凭据管理器，绝不入库。

## 五条联通公约（所有模块共同遵守）

1. 一事一夹：`{data_root}/OA附件库/文档整理/采购项目档案/年-序号-事项名/`（01~06 环节子夹）
2. 文件命名：`文档类型（事项名-金额：X）.pdf`，事项名是全链关联键；上传 OA 时自动去掉金额段
3. 交接面：模块间不传参数，只经 收件箱（进）/ 档案夹（归档）/ 台账状态列（记录）
4. 台账状态机：待确认 → 已保存草稿 → 已提交
5. 人工卡点：工具只保存到草稿，提交永远由人做

## 模块

| 模块 | 功能 | 入口 |
|---|---|---|
| modules/m01_shxiang | 事项审批一条龙（070801）：台账→登录→填表→附件→存草稿→归档回写 | `python modules/m01_shxiang/fill_oa.py --row N --save`；收件箱识别 `intake.py scan/add` |
| modules/m02_sggen | 申购单生成：需求台账→套模板→落收件箱 | `python modules/m02_sggen/make_sg.py --ledger` |
| oa_common | 路径/凭据公共库（keyring 服务名 qwenwork-oa-tus-sound） | 内部引用 |

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
- [ ] m03 比质比价报告单生成
- [ ] m04 供应商新增一条龙
- [ ] m05 合同会签一条龙
- [ ] m06 付款申请一条龙
- [ ] orchestrator 全链编排
