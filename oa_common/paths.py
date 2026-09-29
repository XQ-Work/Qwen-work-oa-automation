# -*- coding: utf-8 -*-
"""集中路径：仓库只存代码，业务数据全部由 config.json 的 data_root 定位（v2 彻底分层）。

目录约定（data_root 下）：
  common/            共享库与跨模块文件（oa_automation.db、合同分支映射表、附件命名规范）
  mNN_模块名/        各模块专属数据（台账等）
  输入/工作/记录/系统   四层见 README（v4）
  danju/             单据类数据根（paths.py DJ_BASE 钩子同源）
  pilot/             Edge 持久化配置与运行状态截图
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
DATA = Path(_cfg["data_root"])
DOCS = Path(_cfg.get("docs_root", DATA.parent / "docs"))


def _p(*rel):
    return DATA.joinpath(*[r for r in rel if r])


def _abs(p):
    q = Path(p)
    return q if q.is_absolute() else DATA / p


DB = _abs(_cfg["db"])
DOCS_DB = _abs(_cfg["docs_db"])
BRANCH_MAP = _abs(_cfg["branch_map"])
NAMING_DOC = _abs(_cfg["naming_doc"])
LEDGER_OA = _abs(_cfg["ledger_oa"])
LEDGER_SG = _abs(_cfg["ledger_sg"])
LEDGER_SUPPLIER = _abs(_cfg["ledger_supplier"])
LEDGER_CONTRACT = _abs(_cfg["ledger_contract"])
LEDGER_PAY = _abs(_cfg["ledger_pay"])
RECEIPT_DIR = _abs(_cfg["receipt_dir"])
DANJU_BASE = _abs(_cfg.get("danju_base", "danju"))
INBOX = _p(*_cfg["inbox_rel"].split("/"))
TPL = _p(*_cfg["template_rel"].split("/"))
TPL_SG = _abs(_cfg["template_sg"])
TPL_BIZ = _abs(_cfg["template_biz"])
ARCHIVE = _p(*_cfg["archive_rel"].split("/"))
EXCEPT = _p(*_cfg["except_rel"].split("/"))
DELIVER = _p(*_cfg["deliver_rel"].split("/"))
BOARD = _abs(_cfg["board_file"])
ARCHIVED_SMALL = _abs(_cfg["archive_small_rel"])
ARCHIVED_LARGE = _abs(_cfg["archive_large_rel"])
ARCHIVE_LOGS = [_abs(_cfg["archivelog_ledger"]), _abs(_cfg["archivelog_copy"])]
PROFILE = _p(*_cfg["profile_rel"].split("/"))
STATE = _p(*_cfg["state_rel"].split("/"))
STATE.mkdir(parents=True, exist_ok=True)
