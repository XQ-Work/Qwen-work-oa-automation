# -*- coding: utf-8 -*-
"""集中路径：仓库只存代码，业务数据全部由 config.json 的 data_root 定位。"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
DATA = Path(_cfg["data_root"])

def _p(*rel):
    return DATA.joinpath(*[r for r in rel if r])

LEDGER_OA = DATA / _cfg["ledger_oa"]
LEDGER_SG = DATA / _cfg["ledger_sg"]
INBOX = _p(*_cfg["inbox_rel"].split("/"))
TPL = _p(*_cfg["template_rel"].split("/"))
STAGING = _p(*_cfg["staging_rel"].split("/"))
ARCHIVE = _p(*_cfg["archive_rel"].split("/"))
PROFILE = _p(*_cfg["profile_rel"].split("/"))
STATE = _p(*_cfg["state_rel"].split("/"))
STATE.mkdir(parents=True, exist_ok=True)
